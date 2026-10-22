#!/usr/bin/env python3
"""Template-bound paragraph writer for reviewed merge adapters.

Supports plain paragraphs, inline pictures and native Office Math. Complex
tables, fields, lists, notes, revisions, shape groups and floating objects
need a task-specific Documents adapter; this helper refuses them instead of
silently flattening them. It never imports donor styles.
"""
from __future__ import annotations
import copy
import posixpath
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree as E

from merge_integrity import Package, IntegrityError, W, M, R, REL, CT, WP, A, q, block_id, rels_part, local
from style_firewall import detach_normalized_fragment, apply_paragraph_prototype


def xml_bytes(root):
    return E.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def find_unit(package, unit_id):
    for part in package.files:
        if not part.endswith(".xml"):
            continue
        for node in package.root(part).iter():
            if node.tag in {q("p"), q("tbl")} and block_id(part, node) == unit_id:
                return part, node
    raise IntegrityError("unknown unit: " + unit_id)


class TemplateWriter:
    def __init__(self, template):
        self.template = Package(template)
        self.files = dict(self.template.files)
        self.root = copy.deepcopy(self.template.root("word/document.xml"))
        rel_name = rels_part("word/document.xml")
        self.rels = copy.deepcopy(self.template.root(rel_name)) if rel_name in self.files else E.Element("{" + REL + "}Relationships", nsmap={None: REL})
        self.types = copy.deepcopy(self.template.root("[Content_Types].xml"))
        self.imported = {}
        self.counter = 0
        self.trace = []
        self.protected_paths = {self.template.path}
        # Freeze original anchors as node references before any insertion.
        self.anchors = {block_id("word/document.xml", n): n for n in self.root.iter() if n.tag in {q("p"), q("tbl")}}

    def role(self, prototype):
        part, paragraph = find_unit(self.template, prototype)
        if paragraph.tag != q("p") or part != "word/document.xml":
            raise IntegrityError("paragraph prototype must come from template main story")
        ppr = paragraph.find(q("pPr"))
        runs = paragraph.findall(q("r"))
        rpr = runs[0].find(q("rPr")) if runs else None
        return copy.deepcopy(ppr), copy.deepcopy(rpr)

    def _import_part(self, source, name):
        key = (source.sha256, name)
        if key in self.imported:
            return self.imported[key]
        if name in {"word/styles.xml", "word/numbering.xml", "word/settings.xml", "word/fontTable.xml"} or name.startswith("word/theme/"):
            raise IntegrityError("payload graph reaches a global formatting part")
        self.counter += 1
        suffix = Path(name).suffix
        destination = f"word/merge_payload_{self.counter:05d}{suffix}"
        while destination in self.files:
            self.counter += 1
            destination = f"word/merge_payload_{self.counter:05d}{suffix}"
        self.imported[key] = destination
        self.files[destination] = source.files[name]
        child_rels = source.relationships(name)
        if child_rels:
            root = copy.deepcopy(source.root(rels_part(name)))
            for rel in root:
                if rel.get("TargetMode") == "External":
                    continue
                old = rel.get("Target", "")
                resolved = old.lstrip("/") if old.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(name), old))
                copied = self._import_part(source, resolved)
                rel.set("Target", posixpath.relpath(copied, posixpath.dirname(destination)))
            self.files[rels_part(destination)] = xml_bytes(root)
        types = source.root("[Content_Types].xml")
        candidates = [n.get("ContentType") for n in types if (local(n.tag) == "Override" and n.get("PartName") == "/" + name) or (local(n.tag) == "Default" and n.get("Extension") == name.rsplit(".", 1)[-1])]
        if not candidates:
            raise IntegrityError("source payload lacks content type")
        E.SubElement(self.types, "{" + CT + "}Override", PartName="/" + destination, ContentType=candidates[0])
        return destination

    def _relationships(self, staged, source, owner):
        mapping = {}
        existing = {n.get("Id") for n in self.rels}
        for node in staged.iter():
            for attr, rid in list(node.attrib.items()):
                if not attr.startswith("{" + R + "}") or local(attr) not in {"id", "embed", "link"}:
                    continue
                if rid not in mapping:
                    old = source.relationships(owner).get(rid)
                    if old is None:
                        raise IntegrityError("missing source relationship")
                    rel = copy.deepcopy(old)
                    if rel.get("TargetMode") != "External":
                        target = rel.get("Target", "")
                        resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(owner), target))
                        rel.set("Target", posixpath.relpath(self._import_part(source, resolved), "word"))
                    i = 1
                    while f"rIdMerge{i}" in existing:
                        i += 1
                    new = f"rIdMerge{i}"
                    existing.add(new)
                    rel.set("Id", new)
                    mapping[rid] = new
                    self.rels.append(rel)
                node.set(attr, mapping[rid])

    def write_paragraph(self, source, source_unit, *, anchor, prototype, position="before"):
        self.protected_paths.add(source.path)
        part, original = find_unit(source, source_unit)
        if original.tag != q("p"):
            raise IntegrityError("use a technical-table adapter for table units")
        forbidden = {q(n) for n in ("numPr", "fldSimple", "fldChar", "instrText", "footnoteReference", "endnoteReference", "commentRangeStart", "commentRangeEnd", "commentReference", "ins", "del", "moveFrom", "moveTo", "pict", "object", "tbl", "sectPr", "bookmarkStart", "bookmarkEnd", "hyperlink")}
        if any(n.tag in forbidden or n.tag == "{" + WP + "}anchor" for n in original.iter()):
            raise IntegrityError("paragraph requires the documented task-specific complex-object adapter")
        if len(list(original.iter(q("p")))) != 1:
            raise IntegrityError("nested text boxes require a complex-object adapter")
        if position not in {"before", "after"} or anchor not in self.anchors:
            raise IntegrityError("invalid frozen template anchor/position")
        ppr, rpr = self.role(prototype)
        if ppr is not None and ppr.find(q("sectPr")) is not None:
            raise IntegrityError("role prototype must not carry a section break")
        formulas = [copy.deepcopy(n) for n in original.iter("{" + M + "}oMath")]
        staged, trace = detach_normalized_fragment(original)
        for node, frozen in zip(list(staged.iter("{" + M + "}oMath")), formulas):
            node.getparent().replace(node, frozen)
        apply_paragraph_prototype(staged, ppr)
        for run in staged.findall(q("r")):
            semantic = run.find(q("rPr"))
            extra = [copy.deepcopy(n) for n in semantic if local(n.tag) in {"vertAlign", "lang", "noProof"}] if semantic is not None else []
            if semantic is not None:
                run.remove(semantic)
            new = copy.deepcopy(rpr) if rpr is not None else E.Element(q("rPr"))
            for child in extra:
                old = new.find(child.tag)
                if old is not None:
                    new.remove(old)
                new.append(child)
            if len(new) or new.attrib:
                run.insert(0, new)
        self._relationships(staged, source, part)
        current_ids = [int(n.get("id")) for n in self.root.iter("{" + WP + "}docPr") if (n.get("id") or "").isdigit()]
        next_id = max(current_ids, default=0) + 1
        for node in staged.iter("{" + WP + "}docPr"):
            node.set("id", str(next_id))
            next_id += 1
        reference = self.anchors[anchor]
        parent = reference.getparent()
        parent.insert(parent.index(reference) + (position == "after"), staged)
        self.trace.append({**trace, "source_unit": source_unit, "source_sha256": source.sha256, "prototype": prototype, "imported_styles": []})
        return staged

    def target_id(self, node):
        return block_id("word/document.xml", node)

    def save(self, output):
        output = Path(output).resolve()
        if output in self.protected_paths:
            raise IntegrityError("writer does not overwrite template or source files")
        self.files["word/document.xml"] = xml_bytes(self.root)
        self.files[rels_part("word/document.xml")] = xml_bytes(self.rels)
        self.files["[Content_Types].xml"] = xml_bytes(self.types)
        with ZipFile(output, "w", ZIP_DEFLATED) as z:
            for name, data in self.files.items():
                z.writestr(name, data)
