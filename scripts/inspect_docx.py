#!/usr/bin/env python3
"""Read-only DOCX inventory for planning a section transplant.

Uses only the standard library so it can inspect packages that python-docx
refuses to open because of a malformed, unrelated relationship.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import posixpath
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS = {"w": W, "r": R, "rel": REL, "wp": WP}

HEADING_RE = re.compile(
    r"^\s*(?:第[一二三四五六七八九十百0-9]+[章节部分]|[一二三四五六七八九十]+[、.]|"
    r"(?:\d+\.){0,4}\d+[．.、)]|[（(][一二三四五六七八九十0-9]+[）)])\s*"
)
CAPTION_RE = re.compile(r"^\s*(图|表|式|公式)\s*[（(]?[0-9一二三四五六七八九十]+(?:[-.．][0-9]+)?[）)]?")
CITATION_RE = re.compile(r"[\[［][0-9]+(?:\s*[-–—,，、]\s*[0-9]+)*[\]］]")


def qname(ns: str, name: str) -> str:
    return f"{{{ns}}}{name}"


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def paragraph_text(el: ET.Element) -> str:
    return "".join(t.text or "" for t in el.iter(qname(W, "t")))


def relationship_source(rels_name: str) -> str | None:
    if rels_name == "_rels/.rels":
        return ""
    if not rels_name.endswith(".rels") or "/_rels/" not in rels_name:
        return None
    parent, filename = rels_name.rsplit("/_rels/", 1)
    return posixpath.join(parent, filename[:-5])


def resolve_target(source: str, target: str) -> str:
    return posixpath.normpath(posixpath.join(posixpath.dirname(source), target)).lstrip("/")


def usable_page_height(root: ET.Element) -> int:
    """Return the last section's printable height in twips, with an A4 fallback."""
    sections = list(root.iter(qname(W, "sectPr")))
    if not sections:
        return 13958
    section = sections[-1]
    page = section.find("./w:pgSz", NS)
    margins = section.find("./w:pgMar", NS)
    if page is None or margins is None:
        return 13958
    try:
        height = int(page.attrib[qname(W, "h")])
        top = int(margins.attrib.get(qname(W, "top"), "1440"))
        bottom = int(margins.attrib.get(qname(W, "bottom"), "1440"))
    except (KeyError, ValueError):
        return 13958
    return max(1440, height - top - bottom)


def pagination_risks(root: ET.Element) -> list[dict]:
    """Conservatively flag row overflow and unsafe caption/table pagination."""
    page_height = usable_page_height(root)
    risks: list[dict] = []
    parents = {child: parent for parent in root.iter() for child in parent}
    for table_index, table in enumerate(root.iter(qname(W, "tbl")), start=1):
        table_props = table.find("./w:tblPr", NS)
        floating = table_props is not None and table_props.find("./w:tblpPr", NS) is not None
        parent = parents.get(table)
        previous = None
        if parent is not None:
            siblings = list(parent)
            position = siblings.index(table)
            previous = siblings[position - 1] if position else None
        caption = (
            previous
            if previous is not None
            and previous.tag == qname(W, "p")
            and re.match(r"^\s*表\s*\d+", paragraph_text(previous))
            else None
        )
        rows = table.findall("./w:tr", NS)
        table_reasons: list[str] = []
        if floating:
            table_reasons.append("floating table positioning requires target-template authorization")
        if caption is not None and len(rows) <= 10:
            caption_props = caption.find("./w:pPr", NS)
            if caption_props is None or caption_props.find("./w:keepNext", NS) is None:
                table_reasons.append("short-table caption is not kept with the table")
            missing_cant_split = [
                index
                for index, row in enumerate(rows, start=1)
                if row.find("./w:trPr/w:cantSplit", NS) is None
            ]
            if missing_cant_split:
                table_reasons.append("short-table rows may split across pages")
        if table_reasons:
            risks.append(
                {
                    "table": table_index,
                    "row": None,
                    "text_prefix": " ".join(paragraph_text(table).split())[:80],
                    "caption": paragraph_text(caption).strip() if caption is not None else None,
                    "row_count": len(rows),
                    "floating": floating,
                    "reasons": table_reasons,
                }
            )
        for row_index, row in enumerate(rows, start=1):
            text = paragraph_text(row)
            compact_chars = len(re.sub(r"\s+", "", text))
            paragraphs = list(row.iter(qname(W, "p")))
            nonempty_paragraphs = sum(bool(paragraph_text(p).strip()) for p in paragraphs)
            nested_tables = sum(1 for _ in row.iter(qname(W, "tbl")))
            drawings = sum(
                1
                for elem in row.iter()
                if elem.tag.rsplit("}", 1)[-1] in {"drawing", "pict", "object"}
            )

            # Chinese project templates average roughly 36 full-width
            # characters per 12 pt line.  A row's height follows its tallest
            # cell, not the sum of parallel cells, so estimate each cell
            # separately.  Inline drawing extents participate in text flow;
            # floating/VML objects remain listed for manual review but are not
            # blindly added because they may be overlays such as checkboxes.
            cell_heights = []
            for cell in row.findall("./w:tc", NS):
                cell_lines = sum(
                    max(1, math.ceil(len(re.sub(r"\s+", "", paragraph_text(p))) / 36))
                    for p in cell.iter(qname(W, "p"))
                    if paragraph_text(p).strip()
                )
                inline_height = 0
                for extent in cell.findall(".//wp:inline/wp:extent", NS):
                    try:
                        inline_height += math.ceil(int(extent.attrib.get("cy", "0")) / 635)
                    except ValueError:
                        pass
                text_height = cell_lines * 360
                cell_heights.append(text_height + max(0, inline_height - 360))
            estimated_height = max(cell_heights, default=0)

            row_props = row.find("./w:trPr", NS)
            cant_split = row_props is not None and row_props.find("./w:cantSplit", NS) is not None
            heights = [] if row_props is None else row_props.findall("./w:trHeight", NS)
            exact_heights = []
            for height in heights:
                if height.attrib.get(qname(W, "hRule")) == "exact":
                    try:
                        exact_heights.append(int(height.attrib.get(qname(W, "val"), "0")))
                    except ValueError:
                        pass

            reasons = []
            if cant_split and estimated_height > int(page_height * 0.9):
                reasons.append("cantSplit row is estimated to exceed one printable page")
            # Allow a small estimation margin for compact multi-column header
            # rows whose labels Word can fit more tightly than this heuristic.
            if exact_heights and estimated_height > int(max(exact_heights) * 1.25):
                reasons.append("content is estimated to exceed exact row height")

            if reasons:
                risks.append(
                    {
                        "table": table_index,
                        "row": row_index,
                        "text_prefix": " ".join(text.split())[:80],
                        "characters": compact_chars,
                        "nonempty_paragraphs": nonempty_paragraphs,
                        "nested_tables": nested_tables,
                        "drawings_or_objects": drawings,
                        "estimated_height_twips": estimated_height,
                        "printable_page_height_twips": page_height,
                        "cant_split": cant_split,
                        "exact_row_heights_twips": exact_heights,
                        "reasons": reasons,
                    }
                )
    return risks


def inventory(path: Path, include_blocks: bool = True) -> dict:
    raw = path.read_bytes()
    result: dict = {
        "file": str(path),
        "size_bytes": len(raw),
        "sha256": sha256(raw),
        "warnings": [],
    }
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        result["part_count"] = len(names)
        if "word/document.xml" not in names:
            raise ValueError("DOCX has no word/document.xml")
        root = ET.fromstring(zf.read("word/document.xml"))
        body = root.find("w:body", NS)
        if body is None:
            raise ValueError("DOCX has no w:body")

        blocks = []
        captions = []
        headings = []
        citations = []
        fields = []
        drawing_count = 0
        table_count = 0
        for idx, child in enumerate(list(body)):
            local = child.tag.rsplit("}", 1)[-1]
            if local == "sectPr":
                continue
            text = paragraph_text(child)
            style = ""
            if local == "p":
                style_el = child.find("./w:pPr/w:pStyle", NS)
                if style_el is not None:
                    style = style_el.attrib.get(qname(W, "val"), "")
            if local == "tbl":
                table_count += 1
            drawings = sum(
                1 for elem in child.iter()
                if elem.tag.rsplit("}", 1)[-1] in {"drawing", "pict"}
            )
            drawing_count += drawings
            for code_el in child.iter(qname(W, "instrText")):
                fields.append({"block": idx, "code": " ".join((code_el.text or "").split())})
            compact = " ".join(text.split())
            record = {
                "index": idx,
                "kind": local,
                "style": style,
                "text": compact,
                "drawings": drawings,
            }
            if include_blocks:
                blocks.append(record)
            if compact and (style.lower().startswith(("heading", "toc")) or HEADING_RE.match(compact)):
                headings.append(record)
            if compact and CAPTION_RE.match(compact):
                captions.append(record)
            for match in CITATION_RE.finditer(compact):
                citations.append({"block": idx, "text": match.group(0)})

        media = []
        for name in sorted(n for n in names if n.startswith("word/media/") and not n.endswith("/")):
            data = zf.read(name)
            media.append({"part": name, "size_bytes": len(data), "sha256": sha256(data)})

        broken = []
        external = []
        rel_count = 0
        for rels_name in sorted(n for n in names if n.endswith(".rels")):
            source = relationship_source(rels_name)
            if source is None:
                continue
            try:
                rel_root = ET.fromstring(zf.read(rels_name))
            except ET.ParseError as exc:
                result["warnings"].append(f"Malformed relationships XML: {rels_name}: {exc}")
                continue
            for rel in rel_root.findall("rel:Relationship", NS):
                rel_count += 1
                rid = rel.attrib.get("Id", "")
                target = rel.attrib.get("Target", "")
                if rel.attrib.get("TargetMode") == "External":
                    external.append({"source": source, "id": rid, "target": target})
                    continue
                resolved = resolve_target(source, target)
                if resolved not in names:
                    broken.append({"source": source, "id": rid, "target": target, "resolved": resolved})

        counts = Counter()
        for elem in root.iter():
            local = elem.tag.rsplit("}", 1)[-1]
            if local in {"bookmarkStart", "commentReference", "footnoteReference", "endnoteReference"}:
                counts[local] += 1

        result.update(
            {
                "block_count": len([c for c in list(body) if c.tag != qname(W, "sectPr")]),
                "table_count": table_count,
                "drawing_count": drawing_count,
                "media": media,
                "headings": headings,
                "captions": captions,
                "citations": citations,
                "fields": fields,
                "relationship_count": rel_count,
                "broken_relationships": broken,
                "external_relationships": external,
                "reference_counts": dict(counts),
                "pagination_risks": pagination_risks(root),
            }
        )
        if include_blocks:
            result["blocks"] = blocks
        if broken:
            result["warnings"].append(f"{len(broken)} internal relationship target(s) are missing")
        if external:
            result["warnings"].append(f"{len(external)} external relationship(s) require review")
        if result["pagination_risks"]:
            result["warnings"].append(
                f"{len(result['pagination_risks'])} table row(s) may hide content because of pagination constraints"
            )
    return result


def markdown_report(data: dict) -> str:
    lines = [
        f"# DOCX inventory: {Path(data['file']).name}",
        "",
        f"- SHA-256: `{data['sha256']}`",
        f"- Body blocks: {data['block_count']}",
        f"- Tables: {data['table_count']}",
        f"- Drawings/pictures: {data['drawing_count']}",
        f"- Media parts: {len(data['media'])}",
        f"- Broken relationships: {len(data['broken_relationships'])}",
        f"- External relationships: {len(data['external_relationships'])}",
        f"- Pagination risks: {len(data['pagination_risks'])}",
        "",
        "## Heading candidates",
        "",
    ]
    for item in data["headings"]:
        lines.append(f"- Block {item['index']} · style `{item['style']}` · {item['text']}")
    lines += ["", "## Caption candidates", ""]
    for item in data["captions"]:
        lines.append(f"- Block {item['index']} · {item['text']}")
    lines += ["", "## Table pagination risks", ""]
    if data["pagination_risks"]:
        for item in data["pagination_risks"]:
            lines.append(
                f"- Table {item['table']}, row {item['row']} · "
                f"{'; '.join(item['reasons'])} · {item['text_prefix']}"
            )
    else:
        lines.append("- None")
    lines += ["", "## Warnings", ""]
    if data["warnings"]:
        lines.extend(f"- {warning}" for warning in data["warnings"])
    else:
        lines.append("- None")
    return "\n".join(lines) + "\n"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx", type=Path)
    parser.add_argument("--json", dest="json_path", type=Path)
    parser.add_argument("--markdown", dest="markdown_path", type=Path)
    parser.add_argument("--summary-only", action="store_true", help="omit full body block list")
    parser.add_argument(
        "--fail-on-pagination-risk",
        action="store_true",
        help="exit with status 1 when a table row may hide content",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        data = inventory(args.docx, include_blocks=not args.summary_only)
    except (OSError, ValueError, zipfile.BadZipFile, ET.ParseError) as exc:
        print(f"inspect_docx: {exc}", file=sys.stderr)
        return 2
    rendered = json.dumps(data, ensure_ascii=False, indent=2)
    if args.json_path:
        args.json_path.write_text(rendered + "\n", encoding="utf-8")
    else:
        print(rendered)
    if args.markdown_path:
        args.markdown_path.write_text(markdown_report(data), encoding="utf-8")
    failed = bool(data["broken_relationships"])
    if args.fail_on_pagination_risk and data["pagination_risks"]:
        failed = True
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
