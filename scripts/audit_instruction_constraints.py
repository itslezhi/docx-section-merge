#!/usr/bin/env python3
"""Audit machine-checkable writing instructions for one DOCX table cell."""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
import unicodedata
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
NS = {"w": W_NS}
CJK_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")
LEADING_NUMBER_RE = re.compile(r"^\s*[（(]?\d+[）)]\s*[.．、]?\s*")


def parse_cell(value: str) -> tuple[int, int, int]:
    try:
        parts = tuple(int(item) for item in value.split(":"))
    except ValueError as exc:
        raise argparse.ArgumentTypeError("cell must be TABLE:ROW:CELL") from exc
    if len(parts) != 3 or any(item < 0 for item in parts):
        raise argparse.ArgumentTypeError("cell must be TABLE:ROW:CELL with non-negative indexes")
    return parts


def text(node: etree._Element) -> str:
    return "".join(node.xpath(".//w:t/text()", namespaces=NS)).strip()


def read_cell_paragraphs(path: Path, location: tuple[int, int, int]) -> list[str]:
    with ZipFile(path) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
    body = root.find(W + "body")
    if body is None:
        raise ValueError("DOCX has no w:body")
    tables = body.findall(W + "tbl")
    table_index, row_index, cell_index = location
    try:
        table = tables[table_index]
        row = table.findall(W + "tr")[row_index]
        cell = row.findall(W + "tc")[cell_index]
    except IndexError as exc:
        raise ValueError(f"cell {table_index}:{row_index}:{cell_index} does not exist") from exc
    return [text(paragraph) for paragraph in cell.findall(W + "p")]


def cjk_hanzi_count(value: str) -> int:
    return len(CJK_RE.findall(value))


def visible_character_count(value: str) -> int:
    return len(re.sub(r"\s+", "", value))


def normalized_for_match(value: str) -> str:
    value = unicodedata.normalize("NFKC", value)
    value = LEADING_NUMBER_RE.sub("", value)
    value = re.sub(r"\s+", "", value)
    return value.casefold()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("document", type=Path)
    parser.add_argument("--cell", required=True, type=parse_cell, metavar="TABLE:ROW:CELL")
    parser.add_argument(
        "--item-regex",
        required=True,
        help="select item paragraphs; compile this expression from the current template",
    )
    parser.add_argument(
        "--exclude-regex",
        action="append",
        default=[],
        help="exclude matching cell paragraphs before item selection; repeat as needed",
    )
    parser.add_argument(
        "--count-unit",
        choices=("cjk_hanzi", "visible"),
        default=None,
        help="length unit compiled from the current template",
    )
    parser.add_argument("--target", type=int)
    parser.add_argument("--min", dest="minimum", type=int)
    parser.add_argument("--max", dest="maximum", type=int)
    parser.add_argument(
        "--tolerance",
        type=float,
        default=0.10,
        help="fraction used with --target; default 0.10",
    )
    parser.add_argument("--expected-items", type=int)
    parser.add_argument(
        "--intermediate-punctuation",
        help="required ending of each item except the last; compile from the current template",
    )
    parser.add_argument("--final-punctuation")
    parser.add_argument(
        "--heading",
        action="append",
        default=[],
        help="expected final正文 heading in item order; repeat for each item",
    )
    parser.add_argument("--json", dest="json_path", type=Path)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.target is not None and (args.target < 0 or not 0 <= args.tolerance < 1):
        print("audit_instruction_constraints: invalid target or tolerance", file=sys.stderr)
        return 2
    if args.minimum is not None and args.maximum is not None and args.minimum > args.maximum:
        print("audit_instruction_constraints: --min cannot exceed --max", file=sys.stderr)
        return 2
    if any(value is not None for value in (args.target, args.minimum, args.maximum)) and not args.count_unit:
        print(
            "audit_instruction_constraints: --count-unit is required with a length constraint",
            file=sys.stderr,
        )
        return 2
    try:
        item_re = re.compile(args.item_regex)
        excludes = [re.compile(value) for value in args.exclude_regex]
        paragraphs = read_cell_paragraphs(args.document, args.cell)
    except (OSError, BadZipFile, KeyError, etree.XMLSyntaxError, ValueError, re.error) as exc:
        print(f"audit_instruction_constraints: {exc}", file=sys.stderr)
        return 2

    eligible = [
        value
        for value in paragraphs
        if value and not any(regex.search(value) for regex in excludes)
    ]
    items = [value for value in eligible if item_re.search(value)]
    content = "".join(items)
    hanzi = cjk_hanzi_count(content)
    visible = visible_character_count(content)

    actual_count = (
        hanzi if args.count_unit == "cjk_hanzi" else visible if args.count_unit == "visible" else None
    )
    minimum = args.minimum
    maximum = args.maximum
    if args.target is not None:
        if minimum is None:
            minimum = math.ceil(args.target * (1 - args.tolerance))
        if maximum is None:
            maximum = math.floor(args.target * (1 + args.tolerance))

    violations: list[dict[str, object]] = []
    if not items:
        violations.append({"kind": "no_items_found", "expected_regex": args.item_regex})
    if minimum is not None and actual_count is not None and actual_count < minimum:
        violations.append(
            {
                "kind": "length_too_short",
                "unit": args.count_unit,
                "expected_min": minimum,
                "actual": actual_count,
            }
        )
    if maximum is not None and actual_count is not None and actual_count > maximum:
        violations.append(
            {
                "kind": "length_too_long",
                "unit": args.count_unit,
                "expected_max": maximum,
                "actual": actual_count,
            }
        )
    if args.expected_items is not None and len(items) != args.expected_items:
        violations.append(
            {"kind": "item_count", "expected": args.expected_items, "actual": len(items)}
        )
    if args.intermediate_punctuation:
        for index, item in enumerate(items[:-1], start=1):
            if not item.rstrip().endswith(args.intermediate_punctuation):
                violations.append(
                    {
                        "kind": "item_separator",
                        "item": index,
                        "expected": args.intermediate_punctuation,
                        "actual_ending": item.rstrip()[-1:] or None,
                    }
                )
    if args.final_punctuation and items:
        actual = items[-1].rstrip()[-1:] or None
        if actual != args.final_punctuation:
            violations.append(
                {
                    "kind": "final_punctuation",
                    "expected": args.final_punctuation,
                    "actual": actual,
                }
            )

    if args.heading:
        if len(args.heading) != len(items):
            violations.append(
                {
                    "kind": "heading_item_cardinality",
                    "expected_headings": len(args.heading),
                    "actual_items": len(items),
                }
            )
        for index, (heading, item) in enumerate(zip(args.heading, items), start=1):
            expected = normalized_for_match(heading)
            actual = normalized_for_match(item)
            if expected not in actual:
                violations.append(
                    {
                        "kind": "heading_correspondence",
                        "item": index,
                        "expected_heading": heading,
                        "actual_item": item,
                    }
                )

    report = {
        "passed": not violations,
        "document": str(args.document),
        "cell": ":".join(str(value) for value in args.cell),
        "item_regex": args.item_regex,
        "item_count": len(items),
        "cjk_hanzi_count": hanzi,
        "visible_character_count": visible,
        "count_unit": args.count_unit,
        "actual_count": actual_count,
        "required_range": {"min": minimum, "max": maximum},
        "items": items,
        "violation_count": len(violations),
        "violations": violations,
    }
    rendered = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
    if args.json_path:
        args.json_path.write_text(rendered, encoding="utf-8")
    else:
        print(rendered, end="")
    return 0 if not violations else 1


if __name__ == "__main__":
    raise SystemExit(main())
