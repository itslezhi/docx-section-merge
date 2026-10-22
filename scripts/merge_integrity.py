#!/usr/bin/env python3
"""Frozen-source inventory and fail-closed verification for merge plan v3.

No source range, fingerprint, or media total supplied by the plan is trusted
as proof. Units are re-enumerated from the input packages on every check.
This is a content gate; formatting, rendering and semantic synthesis retain
separate, hash-bound review gates in merge_pipeline.py.
"""
from __future__ import annotations

import copy
import hashlib
import json
import posixpath
import re
from collections import Counter
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as E

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"w": W, "m": M, "r": R}
PLAN_VERSION = 3
RETENTIONS = {"preserve", "exact_duplicate", "replace_shadow", "authorized_exclusion", "template_condensed", "container_layout"}
TEXT_CHANGE_KINDS = {"heading_number", "caption_number", "formula_number", "citation", "cross_reference", "structural_label"}
STORY = re.compile(r"^word/(document|header[^/]*|footer[^/]*|footnotes|endnotes|comments)\.xml$")


class IntegrityError(ValueError):
    pass


def digest(value):
    if not isinstance(value, bytes):
        value = json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode()
    return hashlib.sha256(value).hexdigest()


def q(name):
    return "{" + W + "}" + name


def local(tag):
    return tag.rsplit("}", 1)[-1]


def path_of(node):
    """Same-tag sibling ordinals, always scoped to a frozen package hash."""
    path = []
    while node.getparent() is not None:
        parent = node.getparent()
        peers = [n for n in parent if n.tag == node.tag]
        path.append([node.tag, peers.index(node) + 1])
        node = parent
    return list(reversed(path))


def block_id(part, node):
    return part + "#" + digest(path_of(node))[:24]


def rels_part(part):
    folder, name = posixpath.split(part)
    return posixpath.join(folder, "_rels", name + ".rels") if part else "_rels/.rels"


class Package:
    def __init__(self, path):
        self.path = Path(path).resolve()
        self.sha256 = digest(self.path.read_bytes())
        with ZipFile(self.path) as z:
            if len(z.namelist()) != len(set(z.namelist())):
                raise IntegrityError("duplicate ZIP members")
            self.files = {n: z.read(n) for n in z.namelist()}
        self.roots = {}
        self._part_digests = {}
        if "word/document.xml" not in self.files:
            raise IntegrityError("missing word/document.xml")
        self.root("word/document.xml")

    def root(self, part):
        if part not in self.roots:
            if part not in self.files:
                raise IntegrityError("missing package part: " + part)
            parser = E.XMLParser(resolve_entities=False, no_network=True)
            self.roots[part] = E.fromstring(self.files[part], parser)
        return self.roots[part]

    def relationships(self, part):
        name = rels_part(part)
        if name not in self.files:
            return {}
        items = list(self.root(name))
        ids = [n.get("Id") for n in items]
        if None in ids or len(ids) != len(set(ids)):
            raise IntegrityError("missing/duplicate relationship ID in " + name)
        return {n.get("Id"): n for n in items}

    def relationship(self, owner, rid, stack=()):
        rel = self.relationships(owner).get(rid)
        if rel is None:
            raise IntegrityError(f"dangling relationship {owner}:{rid}")
        target = rel.get("Target", "")
        if rel.get("TargetMode") == "External":
            return [rel.get("Type"), "external", target]
        path = posixpath.normpath(posixpath.join(posixpath.dirname(owner), target)) if not target.startswith("/") else target[1:]
        if path.startswith("../") or path not in self.files:
            raise IntegrityError(f"missing relationship target {owner}:{rid}")
        return [rel.get("Type"), "internal", self.part_digest(path, stack)]

    def part_digest(self, part, stack=()):
        if part in stack:
            # Relationships can be cyclic; preserve edge type and cycle depth.
            return ["cycle", len(stack) - stack.index(part)]
        if part in self._part_digests and not stack:
            return self._part_digests[part]
        chain = (*stack, part)
        raw = self.files[part]
        if part.endswith(".xml"):
            value = self.canonical(self.root(part), part, chain, object_mode=False)
        else:
            value = digest(raw)
        dependencies = sorted((self.relationship(part, rid, chain) for rid in self.relationships(part)), key=lambda x: json.dumps(x, sort_keys=True))
        result = digest([value, dependencies])
        if not stack:
            self._part_digests[part] = result
        return result

    def canonical(self, node, part, stack=(), object_mode=False):
        """Expanded-name tree, with relationship identities resolved in place."""
        if not isinstance(node.tag, str):
            return None
        # Outside-object layout is checked separately. Keep crop, transforms,
        # technical labels and all formula structure. Only display-size/ID
        # metadata known to be mechanical may vary here.
        picture_extent = node.tag == f"{{{A}}}ext" and any(local(a.tag) == "pic" and "drawingml/2006/picture" in a.tag for a in node.iterancestors())
        if object_mode and (node.tag == f"{{{WP}}}extent" or picture_extent):
            cx, cy = int(node.get("cx", "0")), int(node.get("cy", "0"))
            if cx <= 0 or cy <= 0:
                raise IntegrityError("invalid picture extent")
            return [node.tag, "aspect_ratio", round(cx / cy, 5)]
        if object_mode and node.tag in {f"{{{WP}}}positionH", f"{{{WP}}}positionV", f"{{{WP}}}effectExtent"}:
            return [node.tag, "layout-reviewed-separately"]
        attrs = []
        for key, value in sorted(node.attrib.items()):
            if key.startswith("{" + R + "}") and local(key) in {"id", "embed", "link"}:
                value = self.relationship(part, value, stack)
            elif object_mode and node.tag in {f"{{{WP}}}docPr", f"{{{A}}}cNvPr"} and local(key) in {"id", "name"}:
                continue
            elif key.startswith("{" + W + "}") and local(key).startswith("rsid"):
                continue
            attrs.append([key, value])
        content = node.text if node.text and (local(node.tag) in {"t", "instrText", "delText"} or node.text.strip()) else ""
        children = [self.canonical(c, part, stack, object_mode) for c in node]
        return [node.tag, attrs, content, [c for c in children if c is not None]]

    def note_identity(self, tag, nid):
        stem = {"footnoteReference": "footnote", "endnoteReference": "endnote"}[tag]
        part = f"word/{stem}s.xml"
        hits = self.root(part).xpath(f"./w:{stem}[@w:id=$nid]", namespaces=NS, nid=nid)
        if len(hits) != 1:
            raise IntegrityError("missing/duplicate note " + nid)
        node = copy.deepcopy(hits[0])
        node.attrib.pop(q("id"), None)
        return self.canonical(node, part)

    def stream(self, paragraph, part):
        result = []

        def emit(kind, value, props=None):
            token = [kind, value] if props is None else [kind, value, props]
            if kind == "text" and result and result[-1][0] == "text" and result[-1][2] == props:
                result[-1][1] += value
            else:
                result.append(token)

        def walk(node, semantics=()):
            if node.tag == f"{{{M}}}oMath":
                emit("formula", self.canonical(node, part))
                return
            if node.tag in {q("drawing"), q("pict"), q("object")}:
                emit("object", self.canonical(node, part, object_mode=True))
                return
            if node is not paragraph and node.tag == q("p"):
                return  # nested paragraphs are inventoried separately
            if node.tag in {q("pPr"), q("rPr")}:
                return
            if node.tag == q("r"):
                props = node.find(q("rPr"))
                semantics = tuple((local(n.tag), tuple(sorted(n.attrib.items()))) for n in props if local(n.tag) in {"vertAlign", "position", "sym"}) if props is not None else ()
            if node.tag in {q("t"), q("delText")}:
                emit("text" if node.tag == q("t") else "deleted_text", node.text or "", list(semantics))
                return
            if node.tag in {q("tab"), q("br"), q("cr")}:
                emit("break", [node.tag, sorted(node.attrib.items())])
            if node.tag == q("sym"):
                emit("symbol", sorted(node.attrib.items()))
            if node.tag in {q("instrText"), q("fldChar"), q("fldSimple")}:
                emit("field", [local(node.tag), node.text or "", sorted(node.attrib.items())])
            if node.tag in {q("ins"), q("del"), q("moveFrom"), q("moveTo")}:
                emit("revision", [local(node.tag), sorted(node.attrib.items())])
            if node.tag in {q("footnoteReference"), q("endnoteReference")}:
                emit("note", [local(node.tag), self.note_identity(local(node.tag), node.get(q("id"), ""))])
            if node.tag == q("hyperlink"):
                rid = node.get("{" + R + "}id")
                emit("hyperlink", self.relationship(part, rid) if rid else ["bookmark", node.get(q("anchor"))])
            for child in node:
                walk(child, semantics)
        walk(paragraph)
        return result

    def inventory(self):
        records = {}
        for part in sorted(n for n in self.files if STORY.fullmatch(n)):
            root = self.root(part)
            for node in root.iter():
                standalone_object = node.tag in {q("altChunk"), q("drawing"), q("pict"), q("object")} and not any(a.tag == q("p") for a in node.iterancestors())
                if node.tag not in {q("p"), q("tbl")} and not standalone_object:
                    continue
                kind = "opaque" if standalone_object else ("paragraph" if node.tag == q("p") else "table")
                if kind == "paragraph":
                    signature = {"kind": kind, "stream": self.stream(node, part)}
                elif kind == "opaque":
                    signature = {"kind": kind, "payload": self.canonical(node, part)}
                else:
                    rows = []
                    for row in node.findall(q("tr")):
                        cells = []
                        for cell in row.findall(q("tc")):
                            props = cell.find(q("tcPr"))
                            cells.append([self.canonical(p, part) for p in props if p.tag in {q("gridSpan"), q("vMerge"), q("hMerge")}] if props is not None else [])
                        rows.append(cells)
                    signature = {"kind": kind, "rows": rows}
                key = block_id(part, node)
                cells = []
                for cell in node.iterancestors(q("tc")):
                    row = cell.getparent()
                    table = row.getparent() if row is not None else None
                    if table is not None and table.tag == q("tbl"):
                        cells.append({"table": block_id(part, table), "row": list(table.findall(q("tr"))).index(row), "cell": list(row.findall(q("tc"))).index(cell)})
                records[key] = {"block_id": key, "part": part, "path": path_of(node), "kind": kind, "signature": signature, "fingerprint": digest(signature), "table_cells": cells, "preview": "".join(node.itertext())[:140]}
        return records

    def package_errors(self):
        errors = []
        for part in self.files:
            if part.endswith(".xml") or part.endswith(".rels"):
                try:
                    self.root(part)
                except E.XMLSyntaxError:
                    errors.append("invalid XML: " + part)
        if "[Content_Types].xml" not in self.files:
            errors.append("missing content types")
        else:
            root = self.root("[Content_Types].xml")
            overrides = {n.get("PartName", "").lstrip("/") for n in root if local(n.tag) == "Override"}
            defaults = {n.get("Extension") for n in root if local(n.tag) == "Default"}
            for part in self.files:
                if part != "[Content_Types].xml" and part not in overrides and part.rsplit(".", 1)[-1] not in defaults:
                    errors.append("missing content type: " + part)
        for name in self.files:
            if not name.endswith(".rels"):
                continue
            if name == "_rels/.rels":
                owner = ""
            elif "/_rels/" in name:
                folder, tail = name.rsplit("/_rels/", 1)
                owner = posixpath.join(folder, tail[:-5])
            else:
                errors.append("invalid relationship part: " + name)
                continue
            try:
                for rid in self.relationships(owner):
                    self.relationship(owner, rid)
            except IntegrityError as exc:
                errors.append(str(exc))
        return errors


def read_plan(path):
    path = Path(path).resolve()
    plan = json.loads(path.read_text(encoding="utf-8"))
    if type(plan.get("plan_version")) is not int or plan["plan_version"] != PLAN_VERSION:
        raise IntegrityError("plan_version must be 3; rebuild old plans from source inventory, do not only change the version number")
    for key in ("template", "research_docs", "canonical_sections", "content_register", "numbering_map", "review_comments"):
        if key not in plan:
            raise IntegrityError("missing required plan field: " + key)
    if not plan["research_docs"] or not plan["canonical_sections"] or not plan["content_register"]:
        raise IntegrityError("source list, section list and content register must be nonempty")
    if len(set(plan["canonical_sections"])) != len(plan["canonical_sections"]):
        raise IntegrityError("duplicate canonical sections")
    return plan, path.parent


def load_frozen(descriptor, base):
    if not isinstance(descriptor, dict) or not descriptor.get("file") or not descriptor.get("sha256"):
        raise IntegrityError("each source/template requires file and sha256")
    path = Path(descriptor["file"])
    package = Package(path if path.is_absolute() else base / path)
    if package.sha256 != descriptor["sha256"]:
        raise IntegrityError("frozen input changed: " + path.name)
    return package


def authorized(item):
    authority = item.get("user_authorization")
    return isinstance(authority, dict) and bool(authority.get("instruction")) and bool(authority.get("decision_ref"))


def changed_signature(record, entry, numbering):
    value = copy.deepcopy(record["signature"])
    for change in entry.get("text_changes", []):
        rule = numbering.get(change.get("mapping_id"))
        if not rule or rule.get("kind") not in TEXT_CHANGE_KINDS or rule.get("research_doc_id") != entry["research_doc_id"]:
            raise IntegrityError("unregistered/scopeless text change")
        if rule.get("kind") == "structural_label" and not rule.get("authority"):
            raise IntegrityError("structural text change requires recorded authority")
        if not rule.get("old") or not isinstance(rule.get("new"), str):
            raise IntegrityError("numbering mapping requires literal old/new strings")
        hits = []
        for token in value.get("stream", []):
            if token[0] == "text":
                hits.extend([token] * token[1].count(rule["old"]))
        if len(hits) != 1:
            raise IntegrityError("text change must match exactly one literal occurrence in a semantic text run")
        hits[0][1] = hits[0][1].replace(rule["old"], rule["new"], 1)
    return value


def combined_stream(signatures):
    result = []
    for signature in signatures:
        if signature.get("kind") != "paragraph":
            raise IntegrityError("structural split/join supports paragraph units only")
        for token in copy.deepcopy(signature["stream"]):
            if token[0] == "text" and result and result[-1][0] == "text" and token[2] == result[-1][2]:
                result[-1][1] += token[1]
            else:
                result.append(token)
    return result


def verify(plan_path, candidate=None):
    plan, base = read_plan(plan_path)
    template = load_frozen(plan["template"], base)
    sources, source_records = {}, {}
    for descriptor in plan["research_docs"]:
        sid = descriptor.get("id")
        if not sid or sid in sources:
            raise IntegrityError("missing/duplicate research document id")
        source = load_frozen(descriptor, base)
        sources[sid] = source
        source_records[sid] = source.inventory()
    output_path = candidate or plan.get("merged_docx")
    if not output_path:
        raise IntegrityError("missing merged_docx")
    output_path = Path(output_path)
    output = Package(output_path if output_path.is_absolute() else base / output_path)
    if output.path in {p.path for p in sources.values()} | {template.path}:
        raise IntegrityError("output must not overwrite inputs")
    records = output.inventory()
    errors = output.package_errors()
    for sid, package in sources.items():
        errors += [sid + ": " + e for e in package.package_errors()]
    entries = {}
    expected = {(sid, bid) for sid, blocks in source_records.items() for bid in blocks}
    identifiers = {}
    for item in plan["content_register"]:
        key = (item.get("research_doc_id"), item.get("source_block"))
        eid = item.get("id")
        if key in entries or not eid or eid in identifiers:
            errors.append("duplicate source registration or missing/duplicate entry id: " + str(key))
        if key not in expected:
            errors.append("unknown source block: " + str(key))
        entries[key] = item
        identifiers[eid] = item
    for missing in sorted(expected - set(entries)):
        errors.append("unregistered source block: " + str(missing))
    numbering = {}
    for rule in plan["numbering_map"]:
        if not rule.get("id") or rule["id"] in numbering:
            errors.append("duplicate/missing numbering map id")
        numbering[rule.get("id")] = rule
    used_targets = {}
    groups = {}
    comment_ids = {c.get("id") for c in plan["review_comments"].get("comments", [])}
    for group in plan.get("structural_changes", []):
        gid = group.get("id")
        if not gid or gid in groups:
            errors.append("missing/duplicate structural group id")
        groups[gid] = group
    reports = []
    for key, item in entries.items():
        if key not in expected:
            continue
        source = source_records[key[0]][key[1]]
        retention = item.get("retention")
        issues = []
        targets = item.get("targets", [])
        if retention not in RETENTIONS:
            issues.append("invalid retention; synthesis_evidence and not_applicable are not deletion states")
        if item.get("section") not in plan["canonical_sections"]:
            issues.append("missing/unknown destination section")
        if "uses" not in item or not isinstance(item["uses"], list):
            issues.append("uses must be an independent list")
        else:
            for use in item["uses"]:
                if use.get("kind") != "synthesis" or not use.get("evidence") or use.get("section") not in plan["canonical_sections"] or not use.get("targets") or any(t not in records for t in use.get("targets", [])):
                    issues.append("incomplete synthesis use")
        if any(t not in records for t in targets):
            issues.append("missing target block")
        if retention == "preserve":
            if item.get("structural_group"):
                group = groups.get(item["structural_group"], {})
                if item.get("id") not in group.get("entries", []) or targets != group.get("targets"):
                    issues.append("unregistered structural split/join")
            elif len(targets) != 1:
                issues.append("preserve requires one target or an explicit structural group")
            elif targets[0] in records:
                if targets[0] in used_targets:
                    issues.append("multiple preserve entries share one target; prove an exact_duplicate instead")
                used_targets[targets[0]] = item.get("id")
                try:
                    expected_signature = changed_signature(source, item, numbering)
                    if expected_signature != records[targets[0]]["signature"]:
                        issues.append("content/object signature differs at the mapped position")
                except IntegrityError as exc:
                    issues.append(str(exc))
        elif retention in {"exact_duplicate", "replace_shadow"}:
            canonical = identifiers.get(item.get("covered_by"))
            if not canonical or canonical.get("retention") != "preserve":
                issues.append("covered_by must identify a retained source entry, not a deletion chain")
            elif retention == "exact_duplicate":
                other = source_records.get(canonical.get("research_doc_id"), {}).get(canonical.get("source_block"))
                if not other or source["signature"] != other["signature"] or canonical.get("section") != item.get("section"):
                    issues.append("duplicate not proven by recomputed source signatures")
            if retention == "replace_shadow" and not authorized(item):
                issues.append("shadow replacement lacks user authorization")
            if targets:
                issues.append("deleted source uses covered_by, not direct targets")
        elif retention == "authorized_exclusion":
            if not authorized(item):
                issues.append("exclusion lacks user authorization")
            if targets:
                issues.append("excluded entry has direct targets")
        elif retention == "template_condensed":
            if not item.get("template_constraint") or not item.get("evidence") or not targets:
                issues.append("condensation requires scoped template constraint, evidence and actual targets")
            if item.get("detail_required") is not False:
                issues.append("condensation cannot excuse missing required detailed text")
        elif retention == "container_layout":
            harmless = source["kind"] == "table" or (source["kind"] == "paragraph" and not source["signature"]["stream"])
            if not harmless or not item.get("layout_evidence") or targets:
                issues.append("container_layout only accounts for proven source layout-table topology or empty layout paragraphs, never content")
        if issues:
            errors.extend(f"{item.get('id')}: {issue}" for issue in issues)
        reports.append({"id": item.get("id"), "retention": retention, "passed": not issues, "issues": issues})
    for gid, group in groups.items():
        try:
            members = [identifiers[eid] for eid in group["entries"]]
            targets = group["targets"]
            if len(set(group["entries"])) != len(members) or len(set(targets)) != len(targets):
                raise IntegrityError("repeated structural group member/target")
            if not group.get("authority") or group.get("comment_id") not in comment_ids:
                raise IntegrityError("structural changes require authority and a real review comment")
            if not ((group.get("operation") == "split" and len(members) == 1 and len(targets) >= 2) or (group.get("operation") == "join" and len(members) >= 2 and len(targets) == 1)):
                raise IntegrityError("invalid split/join shape")
            if any(e.get("retention") != "preserve" or e.get("structural_group") != gid or e.get("targets") != targets for e in members):
                raise IntegrityError("inconsistent structural group membership")
            if len({e["research_doc_id"] for e in members}) != 1 or len({e["section"] for e in members}) != 1:
                raise IntegrityError("structural groups cannot cross research documents or sections")
            originals = [source_records[e["research_doc_id"]][e["source_block"]] for e in members]
            for units in (originals, [records[t] for t in targets]):
                containers = {(u["part"], json.dumps(u["path"][:-1])) for u in units}
                ordinals = [u["path"][-1][1] for u in units]
                if len(containers) != 1 or ordinals != list(range(ordinals[0], ordinals[0] + len(ordinals))):
                    raise IntegrityError("split/join units must be adjacent and ordered within one container")
            source_signatures = [changed_signature(original, e, numbering) for original, e in zip(originals, members)]
            if combined_stream(source_signatures) != combined_stream([records[t]["signature"] for t in targets]):
                raise IntegrityError("split/join changed source content or objects")
            for target in targets:
                if target in used_targets:
                    raise IntegrityError("structural target already used by another source/group")
                used_targets[target] = gid
        except (KeyError, IntegrityError) as exc:
            errors.append(f"structural group {gid}: {exc}")
    # Declaring deletion does not delete a paragraph. Find whole-unit remnants
    # in the real output; legitimate retained copies must already be mapped.
    # Template-required mirrors can be exempted only at named actual targets.
    for key, item in entries.items():
        if key not in expected or item.get("retention") not in {"exact_duplicate", "replace_shadow", "authorized_exclusion"}:
            continue
        signature = source_records[key[0]][key[1]]["signature"]
        residual = {bid for bid, record in records.items() if record["signature"] == signature} - set(used_targets)
        allowed = set()
        for exception in item.get("allowed_residual_targets", []):
            if not exception.get("reason") or not exception.get("template_constraint") or exception.get("block_id") not in records:
                errors.append(f"{item['id']}: invalid template-residual exception")
            else:
                allowed.add(exception["block_id"])
        if residual - allowed:
            errors.append(f"{item['id']}: deleted source unit remains at unmapped output positions")
    # Tables are not bags of paragraphs: compare the cell of each mapped unit.
    for key, item in entries.items():
        if key not in expected or item.get("retention") != "preserve" or not item.get("targets") or any(t not in records for t in item["targets"]):
            continue
        for cell in source_records[key[0]][key[1]]["table_cells"]:
            table_entry = entries.get((key[0], cell["table"]), {})
            if table_entry.get("retention") == "preserve" and len(table_entry.get("targets", [])) == 1:
                expected_cell = {"table": table_entry["targets"][0], "row": cell["row"], "cell": cell["cell"]}
                if any(expected_cell not in records[t]["table_cells"] for t in item["targets"]):
                    errors.append(f"{item['id']}: target moved to a different technical-table cell")
    # Citation/cross-reference edits must identify actual source and destination
    # definitions. Content comparison of those definitions runs above as well.
    for rule in numbering.values():
        if rule.get("kind") in {"citation", "cross_reference"}:
            sid = rule.get("research_doc_id")
            definition = source_records.get(sid, {}).get(rule.get("source_definition"))
            target = records.get(rule.get("target_definition"))
            definition_entry = entries.get((sid, rule.get("source_definition")), {})
            if definition_entry.get("retention") == "exact_duplicate":
                definition_entry = identifiers.get(definition_entry.get("covered_by"), {})
            if not definition or not target or rule.get("target_definition") not in definition_entry.get("targets", []) or not rule.get("identity"):
                errors.append("unresolved numbering definition: " + str(rule.get("id")))
    if plan["review_comments"].get("mode") not in {"auto", "review", "clean"}:
        errors.append("invalid review_comments mode")
    return {"passed": not errors, "plan_version": PLAN_VERSION, "candidate_sha256": output.sha256, "source_units": len(expected), "registered_units": len(entries), "errors": errors, "entries": reports,
            "limits": ["semantic synthesis and authorization evidence require review", "effective style, numbering values, all-page visibility and field caches require separate gates"]}


def scaffold(template_path, source_paths):
    template = Package(template_path)
    descriptors, register = [], []
    for i, path in enumerate(source_paths, 1):
        source = Package(path)
        sid = f"research-{i:02d}"
        descriptors.append({"id": sid, "file": str(source.path), "sha256": source.sha256})
        for bid, record in source.inventory().items():
            register.append({"id": f"entry-{len(register)+1:05d}", "research_doc_id": sid, "source_block": bid, "source_preview": record["preview"], "section": None, "retention": "preserve", "targets": [], "uses": []})
    return {"plan_version": PLAN_VERSION, "template": {"file": str(template.path), "sha256": template.sha256}, "research_docs": descriptors,
            "canonical_sections": [], "content_register": register, "numbering_map": [], "review_comments": {"mode": "auto", "comments": []}, "merged_docx": None}
