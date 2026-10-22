#!/usr/bin/env python3
"""Audit template-owned DOCX tables while allowing text-slot replacements.

The audit compares the non-text OOXML skeleton of selected top-level tables:
table/grid properties, row and cell topology, paragraph properties, and run
format patterns.  A shared actual-value typography prototype can be declared
for fixed-form value cells so that target-cell layout is preserved while
instruction-prompt fonts cannot leak into inserted values.  It also detects
extra paragraphs and newly introduced leading whitespace.  Optional role checks flag long narrative paragraphs that
lack a required character-based first-line indent, reference-entry Latin runs
whose effective font differs from the required font, and reference numbers
whose effective weight differs after resolving styles.  The audit also
protects the template's default-style set.  Use one
``--strict-table`` per table declared template-owned in the merge plan.
"""

from __future__ import annotations

import argparse
import copy
import json
import re
import sys
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ET.register_namespace("w", W)

BODY_EXCLUDE_RE = re.compile(
    r"^\s*(?:第[一二三四五六七八九十百0-9]+[章节部分]|"
    r"[一二三四五六七八九十]+[、.]|\d+[．.、]|"
    r"[（(][一二三四五六七八九十0-9]+[）)]|\d+）|"
    r"[①②③④⑤⑥⑦⑧⑨⑩]|[A-Za-z][）)]|"
    r"图\s*\d+|表\s*\d+|式\s*[（(]?\d+|参考文献|"
    r"[\[［]\d+[\]］])"
)
DEFAULT_REFERENCE_RE = r"^\s*[\[［]\d+[\]］]"
NUMBERED_BODY_RE = re.compile(
    r"^\s*(?:[（(]\d+[）)]|\d+）|[①②③④⑤⑥⑦⑧⑨⑩]|[A-Za-z][）)])"
)


def qname(namespace: str, name: str) -> str:
    return f"{{{namespace}}}{name}"


def local_name(name: str) -> str:
    return name.rsplit("}", 1)[-1]


def read_package(path: Path) -> tuple[ET.Element, ET.Element | None]:
    with ZipFile(path) as archive:
        if "word/document.xml" not in archive.namelist():
            raise ValueError(f"{path} has no word/document.xml")
        document = ET.fromstring(archive.read("word/document.xml"))
        styles = (
            ET.fromstring(archive.read("word/styles.xml"))
            if "word/styles.xml" in archive.namelist()
            else None
        )
        return document, styles


def body_tables(root: ET.Element) -> list[ET.Element]:
    body = root.find(qname(W, "body"))
    if body is None:
        raise ValueError("DOCX has no w:body")
    return [child for child in body if child.tag == qname(W, "tbl")]


def strip_volatile(element: ET.Element | None) -> bytes:
    if element is None:
        return b""
    clone = copy.deepcopy(element)
    volatile = {
        "rsidDel",
        "rsidP",
        "rsidR",
        "rsidRDefault",
        "rsidRPr",
        "rsidSect",
        "paraId",
        "textId",
    }
    for node in clone.iter():
        for attr in list(node.attrib):
            if local_name(attr) in volatile:
                del node.attrib[attr]
    xml_text = ET.tostring(clone, encoding="unicode")
    return ET.canonicalize(xml_data=xml_text).encode("utf-8")


def child(element: ET.Element, name: str) -> ET.Element | None:
    return element.find(qname(W, name))


def leading_controls(paragraph: ET.Element) -> list[str]:
    controls: list[str] = []
    for node in paragraph.iter():
        lname = local_name(node.tag)
        if lname == "t":
            value = node.text or ""
            if value:
                if value[0] in {" ", "\t", "\u00a0", "\u3000"}:
                    controls.append(f"leading_{ord(value[0]):04x}")
                break
        elif lname in {"tab", "br"}:
            controls.append(f"leading_{lname}")
            break
    return controls


def normalized_run_properties(
    run: ET.Element,
    *,
    strip_instruction_emphasis: bool = False,
) -> bytes:
    rpr = child(run, "rPr")
    if not strip_instruction_emphasis or rpr is None:
        return strip_volatile(rpr)
    clone = copy.deepcopy(rpr)
    for name in ("i", "iCs", "color", "highlight", "shd", "u"):
        node = child(clone, name)
        while node is not None:
            clone.remove(node)
            node = child(clone, name)
    return strip_volatile(clone)


def normalized_paragraph_layout(paragraph: ET.Element) -> bytes:
    """Return paragraph properties without direct character typography."""
    ppr = child(paragraph, "pPr")
    if ppr is None:
        return b""
    clone = copy.deepcopy(ppr)
    direct_rpr = child(clone, "rPr")
    if direct_rpr is not None:
        clone.remove(direct_rpr)
    return strip_volatile(clone)


def paragraph_direct_typography(paragraph: ET.Element) -> bytes:
    ppr = child(paragraph, "pPr")
    return strip_volatile(child(ppr, "rPr")) if ppr is not None else b""


def value_typography_prototype(
    table: ET.Element,
    row_index: int,
    cell_index: int,
) -> dict[str, object]:
    rows = table.findall(qname(W, "tr"))
    if row_index >= len(rows):
        raise ValueError("value-run prototype row is outside the template table")
    cells = rows[row_index].findall(qname(W, "tc"))
    if cell_index >= len(cells):
        raise ValueError("value-run prototype cell is outside the template row")
    paragraphs = [
        paragraph
        for paragraph in cells[cell_index].findall(qname(W, "p"))
        if paragraph_text(paragraph)
    ]
    if not paragraphs:
        raise ValueError("value-run prototype cell has no visible template paragraph")
    paragraph_typography = {
        paragraph_direct_typography(paragraph) for paragraph in paragraphs
    }
    run_signatures = {
        signature
        for paragraph in paragraphs
        for signature in run_format_pattern(paragraph)
    }
    if not run_signatures:
        run_signatures.add("")
    return {
        "paragraph_typography": paragraph_typography,
        "run_signatures": run_signatures,
    }


def matches_value_typography(
    paragraph: ET.Element,
    prototype: dict[str, object],
) -> bool:
    paragraph_typography = prototype["paragraph_typography"]
    run_signatures = prototype["run_signatures"]
    actual_runs = run_format_pattern(paragraph)
    if not actual_runs:
        actual_runs = [""]
    return (
        paragraph_direct_typography(paragraph) in paragraph_typography
        and all(signature in run_signatures for signature in actual_runs)
    )


def run_format_pattern(
    paragraph: ET.Element,
    *,
    strip_instruction_emphasis: bool = False,
) -> list[str]:
    pattern: list[str] = []
    for run in paragraph.iter(qname(W, "r")):
        has_text = any((node.text or "") for node in run.iter(qname(W, "t")))
        has_control_or_object = any(
            local_name(node.tag) in {"tab", "br", "drawing", "pict", "object", "fldChar"}
            for node in run.iter()
            if node is not run
        )
        has_payload = has_text or has_control_or_object
        if not has_payload:
            continue
        signature = normalized_run_properties(
            run,
            strip_instruction_emphasis=strip_instruction_emphasis,
        ).decode("utf-8")
        if not pattern or pattern[-1] != signature:
            pattern.append(signature)
    return pattern


def special_counts(paragraph: ET.Element) -> dict[str, int]:
    names = {"tab", "br", "drawing", "pict", "object", "fldChar"}
    counts = {name: 0 for name in sorted(names)}
    for node in paragraph.iter():
        lname = local_name(node.tag)
        if lname in names:
            counts[lname] += 1
    return counts


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.iter(qname(W, "t"))).strip()


def on_off_value(element: ET.Element | None) -> bool | None:
    if element is None:
        return None
    value = element.attrib.get(qname(W, "val"))
    return (value or "1").casefold() not in {"0", "false", "off", "no"}


def on_off_attribute(value: str | None) -> bool:
    return value is not None and value.casefold() not in {"0", "false", "off", "no"}


def paragraph_all_italic(paragraph: ET.Element) -> bool:
    text_runs = [
        run
        for run in paragraph.iter(qname(W, "r"))
        if any((node.text or "").strip() for node in run.iter(qname(W, "t")))
    ]
    return bool(text_runs) and all(
        on_off_value(child(child(run, "rPr"), "i")) is True
        if child(run, "rPr") is not None
        else False
        for run in text_runs
    )


def paragraph_all_effectively_bold(paragraph: ET.Element, resolver: "FontResolver") -> bool:
    text_runs = [
        run
        for run in paragraph.iter(qname(W, "r"))
        if any((node.text or "").strip() for node in run.iter(qname(W, "t")))
    ]
    return bool(text_runs) and all(resolver.effective_bold(paragraph, run) for run in text_runs)


def paragraph_table_depth(
    paragraph: ET.Element, table: ET.Element, parents: dict[ET.Element, ET.Element]
) -> int:
    depth = 0
    node = paragraph
    while node in parents:
        node = parents[node]
        if node.tag == qname(W, "tbl"):
            depth += 1
            if node is table:
                return depth
    return depth


def audit_body_role(
    table_index: int,
    table: ET.Element,
    expected_indent: dict[str, str],
    min_chars: int,
    extra_exclusions: list[re.Pattern[str]],
    include_numbered_items: bool,
    styles: ET.Element | None,
    protected_template_texts: set[str],
) -> list[dict[str, object]]:
    violations: list[dict[str, object]] = []
    resolver = FontResolver(styles)
    parents = {child_node: parent for parent in table.iter() for child_node in parent}
    ordinal = 0
    for paragraph in table.iter(qname(W, "p")):
        ordinal += 1
        text = paragraph_text(paragraph)
        if text in protected_template_texts:
            continue
        compact_length = len(re.sub(r"\s+", "", text))
        if compact_length < min_chars:
            continue
        if paragraph_table_depth(paragraph, table, parents) > 1:
            continue
        excluded = BODY_EXCLUDE_RE.match(text) is not None or any(
            regex.search(text) for regex in extra_exclusions
        )
        if excluded:
            numbered_body = (
                include_numbered_items
                and NUMBERED_BODY_RE.match(text) is not None
                and not paragraph_all_effectively_bold(paragraph, resolver)
            )
            if not numbered_body:
                continue
        if paragraph_all_italic(paragraph):
            continue
        ind = paragraph.find(f"./{qname(W, 'pPr')}/{qname(W, 'ind')}")
        attrs = {local_name(key): value for key, value in ind.attrib.items()} if ind is not None else {}
        if attrs != expected_indent:
            violations.append(
                {
                    "kind": "body_first_line_chars",
                    "path": f"table[{table_index}]/paragraph[{ordinal}]",
                    "text_prefix": text[:100],
                    "expected": {"template_indent": expected_indent},
                    "actual": {"candidate_indent": attrs},
                }
            )
    return violations


def template_indent_prototype(table: ET.Element, pattern: re.Pattern[str]) -> dict[str, str]:
    matches = [
        paragraph
        for paragraph in table.iter(qname(W, "p"))
        if pattern.search(paragraph_text(paragraph))
    ]
    if len(matches) != 1:
        raise ValueError(
            f"body prototype regex must match exactly one template paragraph; found {len(matches)}"
        )
    ind = matches[0].find(f"./{qname(W, 'pPr')}/{qname(W, 'ind')}")
    return (
        {local_name(key): value for key, value in ind.attrib.items()}
        if ind is not None
        else {}
    )


class FontResolver:
    def __init__(self, styles: ET.Element | None):
        self.styles = styles
        self.by_id: dict[str, ET.Element] = {}
        self.default_font: str | None = None
        self.default_rpr: ET.Element | None = None
        if styles is None:
            return
        for style in styles.findall(qname(W, "style")):
            style_id = style.attrib.get(qname(W, "styleId"))
            if style_id:
                self.by_id[style_id] = style
        self.default_rpr = styles.find(
            f"./{qname(W, 'docDefaults')}/{qname(W, 'rPrDefault')}/{qname(W, 'rPr')}"
        )
        default_fonts = child(self.default_rpr, "rFonts") if self.default_rpr is not None else None
        self.default_font = self._font_from_element(default_fonts)

    @staticmethod
    def _font_from_element(fonts: ET.Element | None) -> str | None:
        if fonts is None:
            return None
        return fonts.attrib.get(qname(W, "ascii")) or fonts.attrib.get(qname(W, "hAnsi"))

    def _font_from_rpr(self, rpr: ET.Element | None) -> str | None:
        return self._font_from_element(child(rpr, "rFonts")) if rpr is not None else None

    def _style_font(self, style_id: str | None) -> str | None:
        visited: set[str] = set()
        while style_id and style_id not in visited:
            visited.add(style_id)
            style = self.by_id.get(style_id)
            if style is None:
                break
            font = self._font_from_rpr(child(style, "rPr"))
            if font:
                return font
            based = child(style, "basedOn")
            style_id = based.attrib.get(qname(W, "val")) if based is not None else None
        return None

    @staticmethod
    def _bold_from_rpr(rpr: ET.Element | None) -> bool | None:
        if rpr is None:
            return None
        value = on_off_value(child(rpr, "b"))
        return value if value is not None else on_off_value(child(rpr, "bCs"))

    def _style_bold(self, style_id: str | None) -> bool | None:
        visited: set[str] = set()
        while style_id and style_id not in visited:
            visited.add(style_id)
            style = self.by_id.get(style_id)
            if style is None:
                break
            value = self._bold_from_rpr(child(style, "rPr"))
            if value is not None:
                return value
            based = child(style, "basedOn")
            style_id = based.attrib.get(qname(W, "val")) if based is not None else None
        return None

    def effective_latin_font(self, paragraph: ET.Element, run: ET.Element) -> str | None:
        run_rpr = child(run, "rPr")
        direct = self._font_from_rpr(run_rpr)
        if direct:
            return direct
        run_style = child(run_rpr, "rStyle") if run_rpr is not None else None
        if run_style is not None:
            value = self._style_font(run_style.attrib.get(qname(W, "val")))
            if value:
                return value
        ppr = child(paragraph, "pPr")
        paragraph_direct = self._font_from_rpr(child(ppr, "rPr")) if ppr is not None else None
        if paragraph_direct:
            return paragraph_direct
        paragraph_style = child(ppr, "pStyle") if ppr is not None else None
        if paragraph_style is not None:
            value = self._style_font(paragraph_style.attrib.get(qname(W, "val")))
            if value:
                return value
        return self.default_font

    def effective_bold(self, paragraph: ET.Element, run: ET.Element) -> bool:
        run_rpr = child(run, "rPr")
        direct = self._bold_from_rpr(run_rpr)
        if direct is not None:
            return direct
        run_style = child(run_rpr, "rStyle") if run_rpr is not None else None
        if run_style is not None:
            value = self._style_bold(run_style.attrib.get(qname(W, "val")))
            if value is not None:
                return value
        ppr = child(paragraph, "pPr")
        paragraph_direct = self._bold_from_rpr(child(ppr, "rPr")) if ppr is not None else None
        if paragraph_direct is not None:
            return paragraph_direct
        paragraph_style = child(ppr, "pStyle") if ppr is not None else None
        if paragraph_style is not None:
            value = self._style_bold(paragraph_style.attrib.get(qname(W, "val")))
            if value is not None:
                return value
        return bool(self._bold_from_rpr(self.default_rpr))


def audit_reference_fonts(
    root: ET.Element,
    styles: ET.Element | None,
    reference_regex: re.Pattern[str],
    required_font: str,
) -> list[dict[str, object]]:
    violations: list[dict[str, object]] = []
    resolver = FontResolver(styles)
    ordinal = 0
    for paragraph in root.iter(qname(W, "p")):
        text = paragraph_text(paragraph)
        if not reference_regex.search(text):
            continue
        ordinal += 1
        for run_index, run in enumerate(paragraph.iter(qname(W, "r"))):
            run_text = "".join(node.text or "" for node in run.iter(qname(W, "t")))
            if not re.search(r"[A-Za-z]", run_text):
                continue
            actual = resolver.effective_latin_font(paragraph, run)
            if (actual or "").casefold() != required_font.casefold():
                violations.append(
                    {
                        "kind": "reference_latin_font",
                        "path": f"reference[{ordinal}]/run[{run_index}]",
                        "text_prefix": run_text[:100],
                        "expected": required_font,
                        "actual": actual,
                    }
                )
    return violations


def audit_reference_number_weight(
    root: ET.Element,
    styles: ET.Element | None,
    reference_regex: re.Pattern[str],
    required_weight: str,
) -> list[dict[str, object]]:
    violations: list[dict[str, object]] = []
    resolver = FontResolver(styles)
    expected_bold = required_weight == "bold"
    ordinal = 0
    for paragraph in root.iter(qname(W, "p")):
        if not reference_regex.search(paragraph_text(paragraph)):
            continue
        ordinal += 1
        first_run = None
        first_text = ""
        for run in paragraph.iter(qname(W, "r")):
            run_text = "".join(node.text or "" for node in run.iter(qname(W, "t")))
            if run_text.strip():
                first_run = run
                first_text = run_text
                break
        if first_run is None:
            continue
        actual_bold = resolver.effective_bold(paragraph, first_run)
        if actual_bold != expected_bold:
            violations.append(
                {
                    "kind": "reference_number_weight",
                    "path": f"reference[{ordinal}]/first_text_run",
                    "text_prefix": first_text[:40],
                    "expected": required_weight,
                    "actual": "bold" if actual_bold else "normal",
                }
            )
    return violations


def default_style_keys(styles: ET.Element | None) -> list[tuple[str | None, str | None]]:
    if styles is None:
        return []
    return [
        (style.attrib.get(qname(W, "type")), style.attrib.get(qname(W, "styleId")))
        for style in styles.findall(qname(W, "style"))
        if on_off_attribute(style.attrib.get(qname(W, "default")))
    ]


def styles_by_id(styles: ET.Element | None) -> dict[str, ET.Element]:
    if styles is None:
        return {}
    return {
        style.attrib[qname(W, "styleId")]: style
        for style in styles.findall(qname(W, "style"))
        if qname(W, "styleId") in style.attrib
    }


def audit_template_style_authority(
    template_styles: ET.Element | None,
    candidate_styles: ET.Element | None,
) -> list[dict[str, object]]:
    violations: list[dict[str, object]] = []
    expected_defaults = default_style_keys(template_styles)
    actual_defaults = default_style_keys(candidate_styles)
    if actual_defaults != expected_defaults:
        add_diff(
            violations,
            "default_style_set",
            "styles.xml/defaults",
            expected_defaults,
            actual_defaults,
        )
    template_defaults = child(template_styles, "docDefaults") if template_styles is not None else None
    candidate_defaults = child(candidate_styles, "docDefaults") if candidate_styles is not None else None
    compare_property(
        violations,
        "doc_defaults",
        "styles.xml/docDefaults",
        template_defaults,
        candidate_defaults,
    )
    expected_styles = styles_by_id(template_styles)
    actual_styles = styles_by_id(candidate_styles)
    for style_id, expected in expected_styles.items():
        actual = actual_styles.get(style_id)
        if actual is None:
            add_diff(
                violations,
                "missing_template_style",
                f"styles.xml/style[{style_id}]",
                "present",
                "missing",
            )
        elif strip_volatile(expected) != strip_volatile(actual):
            add_diff(
                violations,
                "template_style_modified",
                f"styles.xml/style[{style_id}]",
                "template definition",
                "candidate differs",
            )
    return violations


def add_diff(
    violations: list[dict[str, object]],
    kind: str,
    path: str,
    expected: object,
    actual: object,
) -> None:
    violations.append({"kind": kind, "path": path, "expected": expected, "actual": actual})


def compare_property(
    violations: list[dict[str, object]],
    kind: str,
    path: str,
    left: ET.Element | None,
    right: ET.Element | None,
) -> None:
    if strip_volatile(left) != strip_volatile(right):
        add_diff(violations, kind, path, "template property skeleton", "candidate differs")


def parse_cell_path(value: str) -> tuple[int, int, int]:
    try:
        parts = tuple(int(item) for item in value.split(":"))
    except ValueError as exc:
        raise argparse.ArgumentTypeError("cell path must be TABLE:ROW:CELL") from exc
    if len(parts) != 3 or any(item < 0 for item in parts):
        raise argparse.ArgumentTypeError("cell path must be TABLE:ROW:CELL with non-negative indexes")
    return parts


def compare_table(
    index: int,
    template: ET.Element,
    candidate: ET.Element,
    allow_extra_cells: set[tuple[int, int, int]],
    value_cells: set[tuple[int, int, int]],
    allow_pagination_tables: set[int],
    exact_unchanged_text: bool,
    value_prototype: dict[str, object] | None,
) -> list[dict[str, object]]:
    violations: list[dict[str, object]] = []
    prefix = f"table[{index}]"
    compare_property(violations, "table_properties", prefix, child(template, "tblPr"), child(candidate, "tblPr"))
    compare_property(violations, "table_grid", prefix, child(template, "tblGrid"), child(candidate, "tblGrid"))

    template_rows = template.findall(qname(W, "tr"))
    candidate_rows = candidate.findall(qname(W, "tr"))
    if len(template_rows) != len(candidate_rows):
        add_diff(violations, "row_count", prefix, len(template_rows), len(candidate_rows))

    for row_index, (template_row, candidate_row) in enumerate(zip(template_rows, candidate_rows)):
        row_path = f"{prefix}/row[{row_index}]"
        if index not in allow_pagination_tables:
            compare_property(
                violations,
                "row_properties",
                row_path,
                child(template_row, "trPr"),
                child(candidate_row, "trPr"),
            )
        template_cells = template_row.findall(qname(W, "tc"))
        candidate_cells = candidate_row.findall(qname(W, "tc"))
        if len(template_cells) != len(candidate_cells):
            add_diff(violations, "cell_count", row_path, len(template_cells), len(candidate_cells))

        for cell_index, (template_cell, candidate_cell) in enumerate(zip(template_cells, candidate_cells)):
            cell_path = f"{row_path}/cell[{cell_index}]"
            compare_property(
                violations,
                "cell_properties",
                cell_path,
                child(template_cell, "tcPr"),
                child(candidate_cell, "tcPr"),
            )
            template_paragraphs = template_cell.findall(qname(W, "p"))
            candidate_paragraphs = candidate_cell.findall(qname(W, "p"))
            extra_allowed = (index, row_index, cell_index) in allow_extra_cells
            value_cell = (index, row_index, cell_index) in value_cells
            if (
                len(template_paragraphs) != len(candidate_paragraphs)
                and not extra_allowed
                and not value_cell
            ):
                add_diff(
                    violations,
                    "paragraph_count",
                    cell_path,
                    len(template_paragraphs),
                    len(candidate_paragraphs),
                )

            for paragraph_index, (template_p, candidate_p) in enumerate(
                zip(template_paragraphs, candidate_paragraphs)
            ):
                paragraph_path = f"{cell_path}/paragraph[{paragraph_index}]"
                unchanged_template_paragraph = (
                    paragraph_text(template_p) == paragraph_text(candidate_p)
                )
                actual_value = value_cell and not unchanged_template_paragraph
                if actual_value:
                    if normalized_paragraph_layout(template_p) != normalized_paragraph_layout(candidate_p):
                        add_diff(
                            violations,
                            "value_paragraph_layout",
                            paragraph_path,
                            "target-cell paragraph layout",
                            "candidate layout differs",
                        )
                else:
                    compare_property(
                        violations,
                        "paragraph_properties",
                        paragraph_path,
                        child(template_p, "pPr"),
                        child(candidate_p, "pPr"),
                    )
                expected_pattern = run_format_pattern(template_p)
                actual_pattern = run_format_pattern(candidate_p)
                if actual_value and value_prototype is not None:
                    pattern_valid = matches_value_typography(candidate_p, value_prototype)
                else:
                    pattern_valid = expected_pattern == actual_pattern
                if not pattern_valid:
                    add_diff(
                        violations,
                        "value_typography_role" if actual_value else "run_format_pattern",
                        paragraph_path,
                        "shared actual-value typography" if actual_value else len(expected_pattern),
                        "candidate differs" if actual_value else len(actual_pattern),
                    )
                expected_special = special_counts(template_p)
                actual_special = special_counts(candidate_p)
                if expected_special != actual_special:
                    add_diff(
                        violations,
                        "special_content_pattern",
                        paragraph_path,
                        expected_special,
                        actual_special,
                    )
                template_leading = leading_controls(template_p)
                candidate_leading = leading_controls(candidate_p)
                if candidate_leading and not template_leading:
                    add_diff(
                        violations,
                        "leading_whitespace_or_control",
                        paragraph_path,
                        template_leading,
                        candidate_leading,
                    )
                if (
                    exact_unchanged_text
                    and not actual_value
                    and paragraph_text(template_p)
                    and paragraph_text(template_p) == paragraph_text(candidate_p)
                    and strip_volatile(template_p) != strip_volatile(candidate_p)
                ):
                    add_diff(
                        violations,
                        "unchanged_text_paragraph_xml",
                        paragraph_path,
                        "exact template paragraph XML",
                        "candidate paragraph was rebuilt or reformatted",
                    )

            if len(candidate_paragraphs) > len(template_paragraphs):
                extra_paragraphs = candidate_paragraphs[len(template_paragraphs) :]
                if not (extra_allowed or value_cell):
                    continue
                template_layouts = {
                    normalized_paragraph_layout(paragraph)
                    for paragraph in template_paragraphs
                }
                template_run_signatures = {
                    signature
                    for paragraph in template_paragraphs
                    for signature in run_format_pattern(
                        paragraph,
                        strip_instruction_emphasis=True,
                    )
                }
                if not template_run_signatures:
                    # A blank template paragraph may carry the entire local
                    # prototype in pPr/rPr and legitimately have no run.  A
                    # value run with no direct rPr preserves that inheritance.
                    template_run_signatures.add("")
                for extra_index, paragraph in enumerate(extra_paragraphs):
                    extra_path = (
                        f"{cell_path}/paragraph[{len(template_paragraphs) + extra_index}]"
                    )
                    if normalized_paragraph_layout(paragraph) not in template_layouts:
                        add_diff(
                            violations,
                            "extra_paragraph_template_prototype",
                            extra_path,
                            "one local template paragraph property prototype",
                            "candidate paragraph properties differ",
                        )
                    actual_signatures = run_format_pattern(paragraph)
                    typography_valid = (
                        matches_value_typography(paragraph, value_prototype)
                        if value_cell and value_prototype is not None
                        else all(
                            signature in template_run_signatures
                            for signature in actual_signatures
                        )
                    )
                    if not typography_valid:
                        add_diff(
                            violations,
                            "value_typography_role" if value_cell else "extra_run_template_prototype",
                            extra_path,
                            "shared actual-value typography" if value_cell else "local template run properties",
                            "candidate run properties differ",
                        )

    return violations


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("template", type=Path)
    parser.add_argument("candidate", type=Path)
    parser.add_argument(
        "--strict-table",
        action="append",
        type=int,
        default=[],
        help="zero-based top-level table index; repeat for multiple tables",
    )
    parser.add_argument(
        "--all-template-tables",
        action="store_true",
        help="audit every top-level table present in the template",
    )
    parser.add_argument(
        "--allow-extra-cell",
        action="append",
        type=parse_cell_path,
        default=[],
        metavar="TABLE:ROW:CELL",
        help="allow extra paragraphs in one explicitly authorized cell",
    )
    parser.add_argument(
        "--allow-value-cell",
        action="append",
        type=parse_cell_path,
        default=[],
        metavar="TABLE:ROW:CELL",
        help=(
            "authorize a template slot to receive actual values; target-cell paragraph layout "
            "is preserved and value typography is checked against --value-run-prototype-cell"
        ),
    )
    parser.add_argument(
        "--value-run-prototype-cell",
        type=parse_cell_path,
        metavar="TABLE:ROW:CELL",
        help=(
            "template cell containing the neutral/shared typography prototype for actual "
            "fixed-form values; required whenever --allow-value-cell is used"
        ),
    )
    parser.add_argument(
        "--allow-pagination-table",
        action="append",
        type=int,
        default=[],
        help="ignore row-property differences for an authorized pagination-repair table",
    )
    parser.add_argument(
        "--body-table",
        action="append",
        type=int,
        default=[],
        help="zero-based top-level table containing narrative body paragraphs; repeat as needed",
    )
    parser.add_argument(
        "--body-first-line-chars",
        type=int,
        help="required first-line indentation in hundredths of a character, e.g. 200",
    )
    parser.add_argument(
        "--body-prototype-regex",
        help=(
            "regex selecting exactly one paragraph in each template body table whose full "
            "indentation is the authority for candidate body paragraphs"
        ),
    )
    parser.add_argument(
        "--body-min-chars",
        type=int,
        default=40,
        help="minimum non-whitespace characters for automatic body classification (default: 40)",
    )
    parser.add_argument(
        "--body-exclude-regex",
        action="append",
        default=[],
        help="additional regex excluded from automatic body classification; repeat as needed",
    )
    parser.add_argument(
        "--body-numbered-items",
        action="store_true",
        help="treat non-bold numbered content items as body paragraphs instead of headings",
    )
    parser.add_argument(
        "--reference-latin-font",
        help="audit Latin-bearing runs in reference entries against this effective font",
    )
    parser.add_argument(
        "--reference-number-weight",
        choices=("normal", "bold"),
        help="audit each reference-entry number after resolving character and paragraph styles",
    )
    parser.add_argument(
        "--reference-regex",
        default=DEFAULT_REFERENCE_RE,
        help="regex selecting reference-entry paragraphs",
    )
    parser.add_argument(
        "--exact-unchanged-text",
        action="store_true",
        help="within strict tables, require unchanged visible text to retain exact template paragraph XML",
    )
    parser.add_argument("--json", dest="json_path", type=Path)
    return parser.parse_args()


def unique_indexes(values: Iterable[int]) -> list[int]:
    return list(dict.fromkeys(values))


def main() -> int:
    args = parse_args()
    if (
        not args.strict_table
        and not args.all_template_tables
        and not args.body_table
        and not args.reference_latin_font
        and not args.reference_number_weight
    ):
        print(
            "audit_template_format: select a strict table, body rule, or reference-font rule",
            file=sys.stderr,
        )
        return 2
    if (
        args.body_table
        and args.body_first_line_chars is None
        and args.body_prototype_regex is None
    ):
        print(
            "audit_template_format: --body-table requires --body-prototype-regex "
            "(preferred) or --body-first-line-chars",
            file=sys.stderr,
        )
        return 2
    if args.allow_value_cell and args.value_run_prototype_cell is None:
        print(
            "audit_template_format: --allow-value-cell requires "
            "--value-run-prototype-cell",
            file=sys.stderr,
        )
        return 2
    try:
        template_root, template_styles = read_package(args.template)
        candidate_root, candidate_styles = read_package(args.candidate)
        template_tables = body_tables(template_root)
        candidate_tables = body_tables(candidate_root)
        body_exclusions = [re.compile(value) for value in args.body_exclude_regex]
        body_prototype_regex = (
            re.compile(args.body_prototype_regex)
            if args.body_prototype_regex
            else None
        )
        reference_regex = re.compile(args.reference_regex)
        value_prototype = None
        if args.value_run_prototype_cell is not None:
            prototype_table, prototype_row, prototype_cell = args.value_run_prototype_cell
            if prototype_table >= len(template_tables):
                raise ValueError("value-run prototype table is outside the template")
            value_prototype = value_typography_prototype(
                template_tables[prototype_table],
                prototype_row,
                prototype_cell,
            )
    except (OSError, BadZipFile, ET.ParseError, ValueError, re.error) as exc:
        print(f"audit_template_format: {exc}", file=sys.stderr)
        return 2

    indexes = list(range(len(template_tables))) if args.all_template_tables else unique_indexes(args.strict_table)
    violations: list[dict[str, object]] = []
    violations.extend(audit_template_style_authority(template_styles, candidate_styles))
    for index in indexes:
        if index < 0 or index >= len(template_tables):
            add_diff(violations, "missing_template_table", f"table[{index}]", "present", "missing")
            continue
        if index >= len(candidate_tables):
            add_diff(violations, "missing_candidate_table", f"table[{index}]", "present", "missing")
            continue
        violations.extend(
            compare_table(
                index,
                template_tables[index],
                candidate_tables[index],
                set(args.allow_extra_cell),
                set(args.allow_value_cell),
                set(args.allow_pagination_table),
                args.exact_unchanged_text,
                value_prototype,
            )
        )

    for index in unique_indexes(args.body_table):
        if index < 0 or index >= len(candidate_tables):
            add_diff(violations, "missing_body_table", f"table[{index}]", "present", "missing")
            continue
        if body_prototype_regex is not None:
            if index >= len(template_tables):
                add_diff(
                    violations,
                    "missing_template_body_table",
                    f"table[{index}]",
                    "present",
                    "missing",
                )
                continue
            expected_indent = template_indent_prototype(
                template_tables[index],
                body_prototype_regex,
            )
        else:
            expected_indent = {"firstLineChars": str(args.body_first_line_chars)}
        violations.extend(
            audit_body_role(
                index,
                candidate_tables[index],
                expected_indent,
                args.body_min_chars,
                body_exclusions,
                args.body_numbered_items,
                candidate_styles,
                {
                    paragraph_text(paragraph)
                    for paragraph in template_tables[index].iter(qname(W, "p"))
                    if paragraph_text(paragraph)
                },
            )
        )

    if args.reference_latin_font:
        violations.extend(
            audit_reference_fonts(
                candidate_root,
                candidate_styles,
                reference_regex,
                args.reference_latin_font,
            )
        )

    if args.reference_number_weight:
        violations.extend(
            audit_reference_number_weight(
                candidate_root,
                candidate_styles,
                reference_regex,
                args.reference_number_weight,
            )
        )

    report = {
        "passed": not violations,
        "template": str(args.template),
        "candidate": str(args.candidate),
        "strict_tables": indexes,
        "body_tables": unique_indexes(args.body_table),
        "body_first_line_chars": args.body_first_line_chars,
        "body_prototype_regex": args.body_prototype_regex,
        "body_numbered_items": args.body_numbered_items,
        "reference_latin_font": args.reference_latin_font,
        "reference_number_weight": args.reference_number_weight,
        "exact_unchanged_text": args.exact_unchanged_text,
        "value_run_prototype_cell": args.value_run_prototype_cell,
        "violation_count": len(violations),
        "violations": violations,
        "note": "Text values may change; template-owned skeletons and selected role rules must pass unless explicitly excepted and reviewed.",
    }
    rendered = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
    if args.json_path:
        args.json_path.write_text(rendered, encoding="utf-8")
    else:
        print(rendered, end="")
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
