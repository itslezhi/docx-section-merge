#!/usr/bin/env python3
"""Audit Word review comments against an optional decision manifest."""

from __future__ import annotations

import argparse
import json
import sys
import zipfile
from collections import Counter
from pathlib import Path

from lxml import etree
from merge_integrity import block_id


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
COMMENTS_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
)
COMMENTS_CT = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
)
NS = {"w": W_NS, "pr": PKG_REL_NS}


def w(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def visible_text(node: etree._Element) -> str:
    return "".join(
        node.xpath(".//w:t/text() | .//w:delText/text()", namespaces=NS)
    )


def load_json(path: Path | None) -> dict:
    if path is None:
        return {"comments": []}
    return json.loads(path.read_text(encoding="utf-8"))


def audit(docx: Path, manifest_path: Path | None) -> dict:
    manifest = load_json(manifest_path)
    expected = manifest.get("comments", [])
    violations: list[dict] = []

    with zipfile.ZipFile(docx) as package:
        names = set(package.namelist())
        required = {
            "word/comments.xml",
            "word/document.xml",
            "word/_rels/document.xml.rels",
            "[Content_Types].xml",
        }
        missing = sorted(required - names)
        if missing:
            if expected:
                violations.append({"type": "missing_comment_parts", "parts": missing})
            return {
                "docx": str(docx),
                "manifest": str(manifest_path) if manifest_path else None,
                "comment_count": 0,
                "expected_count": len(expected),
                "violations": violations,
                "passed": not violations and not expected,
            }

        comments_root = etree.fromstring(package.read("word/comments.xml"))
        comment_text: dict[str, str] = {}
        for comment in comments_root.xpath("./w:comment", namespaces=NS):
            cid = comment.get(w("id"))
            if cid is None:
                violations.append({"type": "comment_without_id"})
                continue
            if cid in comment_text:
                violations.append({"type": "duplicate_comment_id", "id": cid})
            comment_text[cid] = visible_text(comment)
            if not comment_text[cid].strip():
                violations.append({"type": "empty_comment", "id": cid})

        story_names = [
            name
            for name in sorted(names)
            if name.startswith("word/")
            and name.endswith(".xml")
            and name != "word/comments.xml"
        ]
        starts: Counter[str] = Counter()
        ends: Counter[str] = Counter()
        refs: Counter[str] = Counter()
        paragraphs: list[dict] = []
        for name in story_names:
            try:
                root = etree.fromstring(package.read(name))
            except etree.XMLSyntaxError:
                continue
            for node in root.xpath(".//w:commentRangeStart", namespaces=NS):
                cid = node.get(w("id"))
                if cid is not None:
                    starts[cid] += 1
            for node in root.xpath(".//w:commentRangeEnd", namespaces=NS):
                cid = node.get(w("id"))
                if cid is not None:
                    ends[cid] += 1
            for node in root.xpath(".//w:commentReference", namespaces=NS):
                cid = node.get(w("id"))
                if cid is not None:
                    refs[cid] += 1
            for index, paragraph in enumerate(root.xpath(".//w:p", namespaces=NS)):
                ids = {
                    node.get(w("id"))
                    for node in paragraph.xpath(
                        ".//w:commentRangeStart | .//w:commentRangeEnd | .//w:commentReference",
                        namespaces=NS,
                    )
                    if node.get(w("id")) is not None
                }
                paragraphs.append(
                    {
                        "part": name,
                        "index": index,
                        "block_id": block_id(name, paragraph),
                        "text": visible_text(paragraph),
                        "comment_ids": ids,
                    }
                )
                ppr = paragraph.find(w("pPr"))
                if ppr is not None and paragraph.index(ppr) != 0:
                    violations.append(
                        {
                            "type": "paragraph_property_order",
                            "part": name,
                            "index": index,
                            "comment_ids": sorted(ids),
                        }
                    )

        all_ids = set(comment_text) | set(starts) | set(ends) | set(refs)
        for cid in sorted(all_ids, key=lambda value: (not value.isdigit(), value)):
            counts = {
                "body": 1 if cid in comment_text else 0,
                "start": starts[cid],
                "end": ends[cid],
                "reference": refs[cid],
            }
            if counts != {"body": 1, "start": 1, "end": 1, "reference": 1}:
                violations.append(
                    {"type": "comment_wiring_count", "id": cid, "counts": counts}
                )

        rels = etree.fromstring(package.read("word/_rels/document.xml.rels"))
        comment_rels = rels.xpath(
            "./pr:Relationship[@Type=$kind and (@Target='comments.xml' or @Target='/word/comments.xml')]",
            namespaces=NS,
            kind=COMMENTS_REL,
        )
        if len(comment_rels) != 1:
            violations.append(
                {"type": "comments_relationship_count", "count": len(comment_rels)}
            )

        content_types = etree.fromstring(package.read("[Content_Types].xml"))
        overrides = content_types.xpath(
            "/*[local-name()='Types']/*[local-name()='Override' and "
            "@PartName='/word/comments.xml' and @ContentType=$kind]",
            kind=COMMENTS_CT,
        )
        if len(overrides) != 1:
            violations.append(
                {"type": "comments_content_type_count", "count": len(overrides)}
            )

        for item in expected:
            anchor = item.get("anchor_contains", "")
            category = item.get("category", "")
            bid = item.get("anchor_block_id")
            matches = [p for p in paragraphs if (anchor or bid) and (not anchor or anchor in p["text"]) and (not bid or bid == p["block_id"])]
            if len(matches) != 1:
                violations.append(
                    {
                        "type": "manifest_anchor_count",
                        "decision_id": item.get("id"),
                        "anchor_contains": anchor,
                        "count": len(matches),
                    }
                )
                continue
            anchored_ids = matches[0]["comment_ids"]
            matching_ids = [
                cid
                for cid in anchored_ids
                if cid in comment_text
                and (not category or comment_text[cid].startswith(f"[{category}]"))
            ]
            if len(matching_ids) != 1:
                violations.append(
                    {
                        "type": "manifest_comment_match",
                        "decision_id": item.get("id"),
                        "anchor_contains": anchor,
                        "category": category,
                        "anchored_ids": sorted(anchored_ids),
                    }
                )
                continue
            expected_text = item.get("comment_text")
            actual_text = comment_text[matching_ids[0]]
            if expected_text is not None and actual_text != expected_text:
                violations.append(
                    {
                        "type": "manifest_comment_text",
                        "decision_id": item.get("id"),
                        "comment_id": matching_ids[0],
                    }
                )

    return {
        "docx": str(docx),
        "manifest": str(manifest_path) if manifest_path else None,
        "comment_count": len(comment_text),
        "expected_count": len(expected),
        "violations": violations,
        "passed": not violations,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--manifest", type=Path)
    parser.add_argument("--json", dest="json_path", type=Path)
    args = parser.parse_args()

    report = audit(args.docx, args.manifest)
    payload = json.dumps(report, ensure_ascii=False, indent=2)
    if args.json_path:
        args.json_path.parent.mkdir(parents=True, exist_ok=True)
        args.json_path.write_text(payload + "\n", encoding="utf-8")
    print(payload)
    sys.exit(0 if report["passed"] else 1)


if __name__ == "__main__":
    main()
