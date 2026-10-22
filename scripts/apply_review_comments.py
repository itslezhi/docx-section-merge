#!/usr/bin/env python3
"""Add real Word comments, with optional frozen block IDs for repeated titles."""
import argparse
import datetime
import json
import shutil
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree as E
from merge_integrity import Package, IntegrityError, W, R, REL, CT, q, block_id


def apply(input_path, manifest, output):
    input_path, output = Path(input_path).resolve(), Path(output).resolve()
    if input_path == output:
        raise IntegrityError("comment writing requires a separate output")
    additions = manifest.get("comments", [])
    if manifest.get("mode", "auto") not in {"auto", "review", "clean"}:
        raise IntegrityError("invalid comment mode")
    if manifest.get("mode") == "clean" and additions:
        raise IntegrityError("clean mode cannot add review comments")
    package = Package(input_path)
    root = package.root("word/document.xml")
    paragraphs = list(root.iter(q("p")))
    targets = []
    decision_ids = set()
    for item in additions:
        if not item.get("id") or item["id"] in decision_ids or not item.get("comment_text"):
            raise IntegrityError("missing/duplicate decision ID or empty comment")
        decision_ids.add(item["id"])
        anchor = item.get("anchor_contains", "")
        bid = item.get("anchor_block_id")
        hits = [p for p in paragraphs if (not bid or block_id("word/document.xml", p) == bid) and (not anchor or anchor in "".join(p.xpath('.//w:t/text()', namespaces={"w": W})))]
        if (not anchor and not bid) or len(hits) != 1:
            raise IntegrityError("comment anchor must resolve to one main-story paragraph")
        targets.append((hits[0], item))
    output.parent.mkdir(parents=True, exist_ok=True)
    if not additions:
        shutil.copyfile(input_path, output)
        return
    comments = package.root("word/comments.xml") if "word/comments.xml" in package.files else E.Element(q("comments"), nsmap={"w": W})
    used = {int(n.get(q("id"))) for n in comments if (n.get(q("id")) or "").isdigit()}
    for p in paragraphs:
        for n in p.iter():
            if n.tag in {q("commentRangeStart"), q("commentRangeEnd"), q("commentReference")} and (n.get(q("id")) or "").isdigit():
                used.add(int(n.get(q("id"))))
    cid = max(used, default=-1) + 1
    for paragraph, item in targets:
        start = E.Element(q("commentRangeStart"), {q("id"): str(cid)})
        paragraph.insert(1 if len(paragraph) and paragraph[0].tag == q("pPr") else 0, start)
        E.SubElement(paragraph, q("commentRangeEnd"), {q("id"): str(cid)})
        ref = E.SubElement(paragraph, q("r"))
        E.SubElement(ref, q("commentReference"), {q("id"): str(cid)})
        comment = E.SubElement(comments, q("comment"), {q("id"): str(cid), q("author"): manifest.get("author", "合稿助手"), q("date"): datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds")})
        p = E.SubElement(comment, q("p"))
        run = E.SubElement(p, q("r"))
        text = E.SubElement(run, q("t"), {"{http://www.w3.org/XML/1998/namespace}space": "preserve"})
        text.text = item["comment_text"]
        cid += 1
    rel_name = "word/_rels/document.xml.rels"
    rels = package.root(rel_name) if rel_name in package.files else E.Element("{" + REL + "}Relationships", nsmap={None: REL})
    kind = R + "/comments"
    matching = [n for n in rels if n.get("Type") == kind]
    if matching and (len(matching) != 1 or matching[0].get("Target") not in {"comments.xml", "/word/comments.xml"}):
        raise IntegrityError("nonstandard comments relationship requires a Documents adapter")
    if not matching:
        ids = {n.get("Id") for n in rels}
        rid = "rIdReviewComments"
        while rid in ids:
            rid += "x"
        E.SubElement(rels, "{" + REL + "}Relationship", Id=rid, Type=kind, Target="comments.xml")
    types = package.root("[Content_Types].xml")
    if not any(n.get("PartName") == "/word/comments.xml" for n in types):
        E.SubElement(types, "{" + CT + "}Override", PartName="/word/comments.xml", ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")
    for name, tree in [("word/document.xml", root), ("word/comments.xml", comments), (rel_name, rels), ("[Content_Types].xml", types)]:
        package.files[name] = E.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)
    with ZipFile(output, "w", ZIP_DEFLATED) as archive:
        for name, data in package.files.items():
            archive.writestr(name, data)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("manifest", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    apply(args.input, json.loads(args.manifest.read_text(encoding="utf-8")), args.output)


if __name__ == "__main__":
    main()
