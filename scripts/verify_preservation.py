#!/usr/bin/env python3
"""Compare donor and merged DOCX ranges while masking allowed numbering changes."""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter
from pathlib import Path

from inspect_docx import CITATION_RE, HEADING_RE, inventory

CAPTION_PREFIX_RE = re.compile(
    r"^(\s*(?:图|表|式|公式)\s*)[（(]?[0-9一二三四五六七八九十]+"
    r"(?:[-.．][0-9]+)?[）)]?\s*"
)


def locate(blocks: list[dict], pattern: str | None, start: int) -> int:
    if pattern is None:
        return len(blocks)
    rx = re.compile(pattern)
    hits = [i for i, block in enumerate(blocks[start:], start) if rx.search(block["text"])]
    if len(hits) != 1:
        raise ValueError(f"selector {pattern!r} matched {len(hits)} blocks; expected exactly 1")
    return hits[0]


def select_range(data: dict, start_pattern: str | None, end_pattern: str | None) -> list[dict]:
    blocks = data["blocks"]
    start = locate(blocks, start_pattern, 0) if start_pattern else 0
    end = locate(blocks, end_pattern, start + 1) if end_pattern else len(blocks)
    if end <= start:
        raise ValueError("end selector must follow start selector")
    return blocks[start:end]


def normalize(text: str, mask_headings: bool, mask_captions: bool, mask_citations: bool) -> str:
    value = " ".join(text.split())
    if mask_captions:
        value = CAPTION_PREFIX_RE.sub(r"\1# ", value)
    if mask_headings:
        value = HEADING_RE.sub("<HNUM> ", value, count=1)
    if mask_citations:
        value = CITATION_RE.sub("[#]", value)
    return value.strip()


def canonical(blocks: list[dict], args: argparse.Namespace) -> list[dict]:
    return [
        {
            "kind": block["kind"],
            "text": normalize(block["text"], args.mask_headings, args.mask_captions, args.mask_citations),
            "drawings": block["drawings"],
        }
        for block in blocks
    ]


def first_differences(left: list[dict], right: list[dict], limit: int = 20) -> list[dict]:
    diffs = []
    total = max(len(left), len(right))
    for i in range(total):
        a = left[i] if i < len(left) else None
        b = right[i] if i < len(right) else None
        if a != b:
            diffs.append({"relative_block": i, "donor": a, "merged": b})
            if len(diffs) >= limit:
                break
    return diffs


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("donor", type=Path)
    parser.add_argument("merged", type=Path)
    parser.add_argument("--donor-start", help="regex selecting the first included donor block")
    parser.add_argument("--donor-end", help="regex selecting the first excluded donor block")
    parser.add_argument("--merged-start", help="regex selecting the first included merged block")
    parser.add_argument("--merged-end", help="regex selecting the first excluded merged block")
    parser.add_argument("--mask-headings", action="store_true")
    parser.add_argument("--mask-captions", action="store_true")
    parser.add_argument("--mask-citations", action="store_true")
    parser.add_argument("--skip-media", action="store_true")
    parser.add_argument("--json", dest="json_path", type=Path)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        donor = inventory(args.donor, include_blocks=True)
        merged = inventory(args.merged, include_blocks=True)
        donor_blocks = select_range(donor, args.donor_start, args.donor_end)
        merged_blocks = select_range(merged, args.merged_start, args.merged_end)
    except (OSError, ValueError) as exc:
        print(f"verify_preservation: {exc}", file=sys.stderr)
        return 2

    donor_canonical = canonical(donor_blocks, args)
    merged_canonical = canonical(merged_blocks, args)
    diffs = first_differences(donor_canonical, merged_canonical)
    media_missing = []
    if not args.skip_media:
        donor_hashes = Counter(item["sha256"] for item in donor["media"])
        merged_hashes = Counter(item["sha256"] for item in merged["media"])
        for digest, count in donor_hashes.items():
            missing = count - merged_hashes[digest]
            if missing > 0:
                media_missing.append({"sha256": digest, "missing_count": missing})

    report = {
        "passed": not diffs and not media_missing,
        "donor_range_blocks": len(donor_blocks),
        "merged_range_blocks": len(merged_blocks),
        "masked": {
            "headings": args.mask_headings,
            "captions": args.mask_captions,
            "citations": args.mask_citations,
        },
        "first_differences": diffs,
        "missing_donor_media": media_missing,
        "note": "Media comparison uses all donor media unless --skip-media is set; use only when the donor file contains the transplanted material.",
    }
    rendered = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
    if args.json_path:
        args.json_path.write_text(rendered, encoding="utf-8")
    else:
        print(rendered, end="")
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
