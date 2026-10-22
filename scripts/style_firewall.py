#!/usr/bin/env python3
"""Prevent donor DOCX styles from changing template-owned formatting.

Use this module from a merge implementation *before* importing a donor
fragment.  ``detach_normalized_fragment`` removes donor style references and
direct prose formatting before any dependency walker can copy those styles
into the target package.  ``localize_imported_style`` is the only supported
way to copy a style for a strict-preserve fragment: it always removes the
global ``w:default`` flag.

The CLI is a transaction gate.  It verifies that a candidate keeps the
template's docDefaults, default-style identities, existing style definitions,
theme and font table.  Added non-default styles are allowed.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import sys
from collections import Counter
from pathlib import Path
from typing import Iterable
from zipfile import BadZipFile, ZipFile

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
NS = {"w": W_NS}
TRUE_VALUES = {None, "1", "true", "on", "yes"}
STYLE_REF_TAGS = ("pStyle", "rStyle", "tblStyle")
DEFAULT_DIRECT_RUN_TAGS = ("vertAlign", "lang", "noProof")
GLOBAL_FORMAT_PARTS = (
    "word/theme/theme1.xml",
    "word/fontTable.xml",
)


class StyleFirewallError(RuntimeError):
    """Raised when a fragment or output violates the template style boundary."""


def qname(name: str) -> str:
    return W + name


def canonical(node: etree._Element | None) -> bytes:
    if node is None:
        return b""
    return etree.tostring(node, method="c14n", with_comments=False)


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def style_map(styles: etree._Element | None) -> dict[str, etree._Element]:
    if styles is None:
        return {}
    return {
        node.get(qname("styleId")): node
        for node in styles.findall(qname("style"))
        if node.get(qname("styleId"))
    }


def default_style_keys(styles: etree._Element | None) -> list[tuple[str | None, str | None]]:
    if styles is None:
        return []
    result: list[tuple[str | None, str | None]] = []
    for node in styles.findall(qname("style")):
        if node.get(qname("default")) in TRUE_VALUES and qname("default") in node.attrib:
            result.append((node.get(qname("type")), node.get(qname("styleId"))))
    return result


def _on_off(node: etree._Element | None) -> bool | None:
    if node is None:
        return None
    value = node.get(qname("val"))
    return (value or "1").casefold() not in {"0", "false", "off", "no"}


def _style_effect(styles: dict[str, etree._Element], style_id: str, prop: str) -> bool | None:
    visited: set[str] = set()
    while style_id and style_id not in visited:
        visited.add(style_id)
        style = styles.get(style_id)
        if style is None:
            return None
        rpr = style.find(qname("rPr"))
        value = _on_off(rpr.find(qname(prop))) if rpr is not None else None
        if value is not None:
            return value
        based = style.find(qname("basedOn"))
        style_id = based.get(qname("val")) if based is not None else ""
    return None


def _iter_including_root(root: etree._Element, tag: str) -> Iterable[etree._Element]:
    # lxml's ``iter(tag)`` already includes the root when it matches.
    yield from root.iter(qname(tag))


def style_references(root: etree._Element) -> Counter[str]:
    refs: Counter[str] = Counter()
    for tag in STYLE_REF_TAGS:
        for node in _iter_including_root(root, tag):
            value = node.get(qname("val"))
            if value:
                refs[value] += 1
    return refs


def assert_no_style_references(root: etree._Element) -> None:
    refs = style_references(root)
    if refs:
        rendered = ", ".join(f"{key}={count}" for key, count in sorted(refs.items()))
        raise StyleFirewallError(f"normalized fragment still contains donor style references: {rendered}")


def detach_normalized_fragment(
    element: etree._Element,
    donor_styles: etree._Element | None = None,
    *,
    semantic_character_styles: Iterable[str] = (),
    preserve_direct_run_tags: Iterable[str] = DEFAULT_DIRECT_RUN_TAGS,
    reject_tables: bool = True,
) -> tuple[etree._Element, dict[str, object]]:
    """Return a donor-style-free staging fragment for template normalization.

    Call this before relationship/numbering import.  The returned prose has no
    paragraph, character or table style references and no donor paragraph
    layout.  The caller must then apply a frozen template role prototype.

    Character styles are never copied.  For explicitly whitelisted semantic
    character styles, only their effective bold/italic meaning is converted to
    direct properties.  This prevents names such as ``Strong`` from silently
    changing structural tokens such as bibliography numbers.
    """

    staged = copy.deepcopy(element)
    if reject_tables and (staged.tag == qname("tbl") or staged.find(".//" + qname("tbl")) is not None):
        raise StyleFirewallError(
            "template-normalize fragment contains a table; map it to a template table role "
            "and transfer only its content payload"
        )

    styles = style_map(donor_styles)
    semantic = set(semantic_character_styles)
    allowed_direct = set(preserve_direct_run_tags)
    removed_refs: Counter[str] = Counter()
    semantic_overlays: Counter[str] = Counter()

    paragraphs = list(_iter_including_root(staged, "p"))
    for paragraph in paragraphs:
        ppr = paragraph.find(qname("pPr"))
        if ppr is not None:
            for style_node in ppr.findall(qname("pStyle")):
                sid = style_node.get(qname("val"))
                if sid:
                    removed_refs[sid] += 1
            paragraph.remove(ppr)

    runs = list(_iter_including_root(staged, "r"))
    for run in runs:
        # Drawing/OLE anchor runs are treated as technical objects, not prose.
        if any(run.find(".//" + qname(tag)) is not None for tag in ("drawing", "object", "pict")):
            continue
        old_rpr = run.find(qname("rPr"))
        style_id = None
        if old_rpr is not None:
            rstyle = old_rpr.find(qname("rStyle"))
            if rstyle is not None:
                style_id = rstyle.get(qname("val"))
                if style_id:
                    removed_refs[style_id] += 1

        replacements: list[etree._Element] = []
        if old_rpr is not None:
            for child in old_rpr:
                if etree.QName(child).localname in allowed_direct:
                    replacements.append(copy.deepcopy(child))
        if style_id in semantic:
            for prop in ("b", "i"):
                value = _style_effect(styles, style_id, prop)
                if value is not None:
                    node = etree.Element(qname(prop))
                    node.set(qname("val"), "1" if value else "0")
                    replacements.append(node)
                    semantic_overlays[f"{style_id}:{prop}={int(value)}"] += 1

        if old_rpr is not None:
            run.remove(old_rpr)
        if replacements:
            new_rpr = etree.Element(qname("rPr"))
            for child in replacements:
                new_rpr.append(child)
            run.insert(0, new_rpr)

    # Defensive pass: style refs can also occur in nested or unusual OOXML.
    for tag in STYLE_REF_TAGS:
        for node in list(_iter_including_root(staged, tag)):
            parent = node.getparent()
            if parent is None:
                continue
            sid = node.get(qname("val"))
            if sid:
                removed_refs[sid] += 1
            parent.remove(node)

    assert_no_style_references(staged)
    return staged, {
        "policy": "template_normalize",
        "removed_style_references": dict(sorted(removed_refs.items())),
        "semantic_overlays": dict(sorted(semantic_overlays.items())),
        "requires_template_role_prototype": True,
    }


def localize_imported_style(
    style: etree._Element,
    new_style_id: str,
    dependency_map: dict[str, str] | None = None,
) -> etree._Element:
    """Copy one strict-preserve style as a non-default, locally named style."""

    result = copy.deepcopy(style)
    result.set(qname("styleId"), new_style_id)
    result.attrib.pop(qname("default"), None)
    mapping = dependency_map or {}
    for tag in ("basedOn", "next", "link"):
        node = result.find(qname(tag))
        if node is None:
            continue
        old = node.get(qname("val"))
        if old in mapping:
            node.set(qname("val"), mapping[old])
    return result


def apply_paragraph_prototype(paragraph: etree._Element, prototype_ppr: etree._Element | None) -> None:
    """Replace paragraph layout with the frozen template role prototype."""

    current = paragraph.find(qname("pPr"))
    if current is not None:
        paragraph.remove(current)
    if prototype_ppr is not None:
        paragraph.insert(0, copy.deepcopy(prototype_ppr))


def read_package(path: Path) -> dict[str, bytes]:
    with ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def audit_template_globals(template: Path, candidate: Path) -> list[dict[str, object]]:
    """Return format-global violations that must abort a merge transaction."""

    template_files = read_package(template)
    candidate_files = read_package(candidate)
    violations: list[dict[str, object]] = []
    if "word/styles.xml" not in template_files or "word/styles.xml" not in candidate_files:
        return [{"kind": "missing_styles_part", "path": "word/styles.xml"}]

    left = etree.fromstring(template_files["word/styles.xml"])
    right = etree.fromstring(candidate_files["word/styles.xml"])
    expected_defaults = default_style_keys(left)
    actual_defaults = default_style_keys(right)
    if actual_defaults != expected_defaults:
        violations.append(
            {
                "kind": "default_style_set",
                "path": "word/styles.xml",
                "expected": expected_defaults,
                "actual": actual_defaults,
            }
        )

    left_defaults = left.find(qname("docDefaults"))
    right_defaults = right.find(qname("docDefaults"))
    if canonical(left_defaults) != canonical(right_defaults):
        violations.append({"kind": "doc_defaults_changed", "path": "word/styles.xml/docDefaults"})

    right_styles = style_map(right)
    for style_id, expected in style_map(left).items():
        actual = right_styles.get(style_id)
        if actual is None:
            violations.append(
                {"kind": "template_style_missing", "path": f"word/styles.xml/style[{style_id}]"}
            )
        elif canonical(expected) != canonical(actual):
            violations.append(
                {"kind": "template_style_changed", "path": f"word/styles.xml/style[{style_id}]"}
            )

    for part in GLOBAL_FORMAT_PARTS:
        expected = template_files.get(part)
        actual = candidate_files.get(part)
        if expected is None and actual is None:
            continue
        if expected is None or actual is None or sha256(expected) != sha256(actual):
            violations.append({"kind": "template_global_part_changed", "path": part})
    return violations


def assert_template_globals(template: Path, candidate: Path) -> None:
    violations = audit_template_globals(template, candidate)
    if violations:
        kinds = ", ".join(sorted({str(item["kind"]) for item in violations}))
        raise StyleFirewallError(f"template format transaction rejected: {kinds}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("template", type=Path)
    parser.add_argument("candidate", type=Path)
    parser.add_argument("--json", dest="json_path", type=Path)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        violations = audit_template_globals(args.template, args.candidate)
    except (OSError, BadZipFile, etree.XMLSyntaxError, ValueError) as exc:
        print(f"style_firewall: {exc}", file=sys.stderr)
        return 2
    report = {
        "passed": not violations,
        "template": str(args.template),
        "candidate": str(args.candidate),
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
