#!/usr/bin/env python3
"""Discover writing and formatting requirement candidates in a DOCX template.

The output is intentionally a candidate list: natural-language instructions
must still be reviewed and compiled into the current template's constraint
registry before merging content.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
W = f"{{{W_NS}}}"
NS = {"w": W_NS, "a": A_NS}
PART_RE = re.compile(
    r"^word/(?:document|header\d+|footer\d+|footnotes|endnotes|comments|"
    r"diagrams/data\d+)\.xml$"
)
INSTRUCTION_MARKERS = re.compile(
    r"(?:应当|应|必须|须|不得|禁止|严禁|不超过|不多于|不少于|不低于|"
    r"字左右|字以内|分条|逐条|列出|分号|句号|对应|一致|明确|重点阐述|"
    r"量化|考核|原始基数|附.+报告|要求|仅填|填写|说明|字体|字号|"
    r"行距|缩进|页边距|纸张|表格|页眉|页脚|附件|提交|提供|编号)"
)

LENGTH_PATTERNS = (
    ("length_around", re.compile(r"(?P<value>\d+)\s*字\s*(?:左右|上下|前后)")),
    ("length_max", re.compile(r"(?:不超过|不多于|最多)\s*(?P<value>\d+)\s*字")),
    ("length_max", re.compile(r"(?P<value>\d+)\s*字\s*以内")),
    ("length_min", re.compile(r"(?:不少于|不低于|至少)\s*(?P<value>\d+)\s*字")),
)

PHRASE_RULES = {
    "structure": (
        "分条列出",
        "逐条列出",
        "不得分述",
        "逐项填写",
        "按年度填写",
        "分层次",
        "有条理",
    ),
    "punctuation": ("分号隔开", "句号结束", "分号结束", "逗号隔开"),
    "correspondence": (
        "相对应",
        "与正文",
        "与封面一致",
        "前后一致",
        "保持一致",
    ),
    "content": (
        "量化可考核",
        "明确原始基数",
        "重点阐述",
        "附查新报告",
        "完整、清晰",
        "重点突出",
    ),
    "prohibition": ("不得", "禁止", "严禁", "仅填"),
    "format": (
        "字体",
        "字号",
        "加粗",
        "斜体",
        "首行缩进",
        "悬挂缩进",
        "单倍行距",
        "多倍行距",
        "固定值",
        "居中",
        "左对齐",
        "右对齐",
    ),
    "layout": (
        "页边距",
        "纸张",
        "A4",
        "页眉",
        "页脚",
        "不得修改表格",
        "表格格式不得修改",
        "表格不得调整",
    ),
    "numbering": ("连续编号", "编号连续", "按顺序编号", "统一编号"),
    "deliverable": ("附件", "提交", "提供", "附查新报告", "附报告"),
}

FORMAT_PATTERNS = (
    (
        "font_family",
        re.compile(
            r"(?P<value>宋体|黑体|仿宋(?:_GB2312)?|楷体(?:_GB2312)?|"
            r"方正小标宋简体|Times\s+New\s+Roman|Arial)",
            re.IGNORECASE,
        ),
    ),
    (
        "font_size_named",
        re.compile(
            r"(?P<value>初号|小初|一号|小一|二号|小二|三号|小三|四号|小四|"
            r"五号|小五|六号|小六|七号|八号)(?:字)?"
        ),
    ),
    ("font_size_points", re.compile(r"(?P<value>\d+(?:\.\d+)?)\s*磅(?:字)?")),
    ("line_spacing_multiple", re.compile(r"(?P<value>\d+(?:\.\d+)?)\s*倍行距")),
    ("line_spacing_fixed", re.compile(r"固定值\s*(?P<value>\d+(?:\.\d+)?)\s*磅")),
    (
        "indent_characters",
        re.compile(r"(?P<value>\d+(?:\.\d+)?)\s*(?:个)?(?:汉字|字符)(?:的)?(?:首行)?缩进"),
    ),
    (
        "indent_characters",
        re.compile(r"(?:首行)?缩进\s*(?P<value>\d+(?:\.\d+)?)\s*(?:个)?(?:汉字|字符)"),
    ),
)


def text(node: etree._Element) -> str:
    return "".join(node.xpath(".//w:t/text() | .//a:t/text()", namespaces=NS)).strip()


def local_name(node: etree._Element) -> str:
    return etree.QName(node).localname


def element_path(node: etree._Element, root: etree._Element) -> str:
    parts: list[str] = []
    current = node
    while current is not None:
        parent = current.getparent()
        name = local_name(current)
        if parent is None:
            parts.append(name)
            break
        siblings = [child for child in parent if child.tag == current.tag]
        parts.append(f"{name}[{siblings.index(current)}]")
        if current is root:
            break
        current = parent
    return "/".join(reversed(parts))


def direct_run_signal(paragraph: etree._Element, tag: str) -> bool:
    return paragraph.find(f".//{W}r/{W}rPr/{W}{tag}") is not None


def candidate_id(part: str, location: str, value: str) -> str:
    digest = hashlib.sha256(f"{part}\0{location}\0{value}".encode("utf-8")).hexdigest()
    return "template-rule-candidate-" + digest[:16]


def parse_rules(value: str) -> list[dict[str, object]]:
    rules: list[dict[str, object]] = []
    seen: set[tuple[object, ...]] = set()
    for kind, regex in LENGTH_PATTERNS:
        for match in regex.finditer(value):
            item = (kind, int(match.group("value")), match.group(0))
            if item in seen:
                continue
            seen.add(item)
            rules.append({"type": kind, "value": item[1], "source": item[2]})
    for kind, phrases in PHRASE_RULES.items():
        found = [phrase for phrase in phrases if phrase in value]
        if found:
            rules.append({"type": kind, "signals": found})
    for kind, regex in FORMAT_PATTERNS:
        for match in regex.finditer(value):
            source = match.group(0)
            item = (kind, match.group("value"), source)
            if item in seen:
                continue
            seen.add(item)
            rules.append({"type": kind, "value": item[1], "source": source})
    return rules


def control_candidates(root: etree._Element, part: str) -> list[dict[str, object]]:
    candidates: list[dict[str, object]] = []
    for sdt in root.iter(W + "sdt"):
        props = sdt.find(W + "sdtPr")
        if props is None:
            continue
        values: list[dict[str, str]] = []
        for tag in ("alias", "tag", "placeholder"):
            for node in props.findall(".//" + W + tag):
                value = node.get(W + "val")
                if value:
                    values.append({"field": tag, "value": value})
        if values:
            location = element_path(sdt, root)
            value = text(sdt)
            candidates.append(
                {
                    "candidate_id": candidate_id(part, location, value),
                    "part": part,
                    "location": location,
                    "text": value,
                    "source_signals": ["content_control_metadata"],
                    "metadata": values,
                    "parsed_rules": parse_rules(" ".join(item["value"] for item in values)),
                    "requires_semantic_review": True,
                }
            )
    return candidates


def discover(path: Path) -> dict[str, object]:
    requirements: list[dict[str, object]] = []
    with ZipFile(path) as archive:
        parts = sorted(name for name in archive.namelist() if PART_RE.match(name))
        for part in parts:
            root = etree.fromstring(archive.read(part))
            paragraphs = list(root.iter(W + "p")) + list(root.iter(f"{{{A_NS}}}p"))
            paragraph_texts = [text(paragraph) for paragraph in paragraphs]
            for index, paragraph in enumerate(paragraphs):
                value = text(paragraph)
                if not value:
                    continue
                italic = direct_run_signal(paragraph, "i") or direct_run_signal(paragraph, "iCs")
                colored = False
                for color in paragraph.findall(f".//{W}r/{W}rPr/{W}color"):
                    setting = (color.get(W + "val") or "").casefold()
                    if setting and setting not in {"auto", "000000"}:
                        colored = True
                        break
                parsed = parse_rules(value)
                markers = sorted(set(INSTRUCTION_MARKERS.findall(value)))
                bracketed = (value.startswith(("（", "(")) and value.endswith(("）", ")")))
                if not parsed and not markers and not italic and not colored:
                    continue
                signals: list[str] = []
                if markers:
                    signals.append("instruction_language")
                if italic:
                    signals.append("italic_prompt")
                if colored:
                    signals.append("colored_prompt")
                if bracketed:
                    signals.append("bracketed_prompt")
                if paragraph.xpath("ancestor-or-self::w:tc", namespaces=NS):
                    signals.append("table_cell")
                location = element_path(paragraph, root)
                requirements.append(
                    {
                        "candidate_id": candidate_id(part, location, value),
                        "part": part,
                        "location": location,
                        "text": value,
                        "context_before": next(
                            (item for item in reversed(paragraph_texts[:index]) if item), None
                        ),
                        "context_after": next(
                            (item for item in paragraph_texts[index + 1 :] if item), None
                        ),
                        "source_signals": signals,
                        "instruction_markers": markers,
                        "parsed_rules": parsed,
                        "requires_semantic_review": True,
                    }
                )
            requirements.extend(control_candidates(root, part))

    return {
        "template": str(path),
        "candidate_count": len(requirements),
        "requirements": requirements,
        "status": "candidates_only",
        "coverage_rule": (
            "Every candidate must be classified as accepted_requirement, duplicate_requirement, "
            "template_label_or_example, or not_a_requirement; no candidate may remain unreviewed. "
            "Rendered pages must also be reviewed for requirements embedded in images or unsupported objects."
        ),
        "next_step": (
            "Review every candidate and the template's complete visible text, determine scope, "
            "then compile the current template's template_constraints.json."
        ),
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("template", type=Path)
    parser.add_argument("--json", dest="json_path", type=Path)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        report = discover(args.template)
    except (OSError, BadZipFile, KeyError, etree.XMLSyntaxError, ValueError) as exc:
        print(f"extract_template_requirements: {exc}", file=sys.stderr)
        return 2
    rendered = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
    if args.json_path:
        args.json_path.write_text(rendered, encoding="utf-8")
    else:
        print(rendered, end="")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
