#!/usr/bin/env python3
"""Inventory -> reviewed builder -> numbering -> comments -> checks -> release.

The builder is a task-authored Python adapter supplied explicitly on the CLI,
never an executable path read from a donor document or merge plan. Semantic
mapping stays with the model; this runner owns order, clean staging, input
immutability and release gates. It does not claim a universal DOCX renderer.
"""
from __future__ import annotations
import argparse
import copy
import importlib.util
import json
import shutil
import subprocess
import sys
from pathlib import Path

from merge_integrity import Package, IntegrityError, digest, load_frozen, read_plan, scaffold, verify

HERE = Path(__file__).resolve().parent
GATES = {"template_constraints", "template_format", "numbering_references", "semantic_review", "visual_pages"}


def write_json(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def run_tool(script, *args):
    completed = subprocess.run([sys.executable, str(HERE / script), *map(str, args)], capture_output=True, text=True)
    if completed.returncode:
        raise IntegrityError(script + " failed: " + (completed.stderr or completed.stdout)[-5000:])


def check(plan_path, candidate):
    content = verify(plan_path, candidate)
    if not content["passed"]:
        return content
    plan, base = read_plan(plan_path)
    template = load_frozen(plan["template"], base)
    run_tool("style_firewall.py", template.path, candidate)
    # The legacy firewall allows non-default additions. v3 also requires
    # every added style to have a scoped, recorded exception.
    source = template.root("word/styles.xml")
    output = Package(candidate)
    styles = output.root("word/styles.xml")
    attr = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId"
    original_ids = {n.get(attr) for n in source if n.get(attr)}
    added = {n.get(attr) for n in styles if n.get(attr)} - original_ids
    approved = {x.get("style_id") for x in plan.get("style_exceptions", []) if x.get("authority") and x.get("scope")}
    if added - approved:
        raise IntegrityError("unregistered added styles: " + ", ".join(sorted(added - approved)))
    for index, flags in enumerate(plan.get("format_checks", [])):
        if not isinstance(flags, list) or not all(isinstance(v, str) for v in flags):
            raise IntegrityError("format_checks entries must be literal argument arrays")
        if "--json" in flags:
            raise IntegrityError("format_checks cannot redirect audit output")
        run_tool("audit_template_format.py", template.path, candidate, *flags)
    manifest = Path(plan_path).with_suffix(".comments.json")
    write_json(manifest, plan["review_comments"])
    run_tool("audit_review_comments.py", candidate, "--manifest", manifest)
    return {**content, "machine_gates": ["complete_source_coverage", "package_relationships", "template_style_isolation", "registered_format_checks", "comments"]}


def build(plan_path, builder_path, work_dir):
    plan, base = read_plan(plan_path)
    if work_dir.exists() and any(work_dir.iterdir()):
        raise IntegrityError("work directory must be new or empty; keep passed checkpoints separately")
    work_dir.mkdir(parents=True, exist_ok=True)
    for descriptor in [plan["template"], *plan["research_docs"]]:
        frozen = load_frozen(descriptor, base)
        descriptor["file"] = str(frozen.path)
    candidate = work_dir / "candidate.docx"
    if candidate.resolve() in {Path(d["file"]).resolve() for d in [plan["template"], *plan["research_docs"]]}:
        raise IntegrityError("candidate aliases an input")
    shutil.copyfile(plan["template"]["file"], candidate)
    spec = importlib.util.spec_from_file_location("task_merge_adapter", builder_path.resolve())
    if spec is None or spec.loader is None:
        raise IntegrityError("cannot load reviewed builder")
    adapter = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(adapter)
    if not callable(getattr(adapter, "build", None)):
        raise IntegrityError("builder must expose build(plan, candidate_path)")
    result = adapter.build(copy.deepcopy(plan), candidate)
    if not isinstance(result, dict):
        raise IntegrityError("build must return the resolved plan with actual target block IDs")
    if result.get("template") != plan["template"] or result.get("research_docs") != plan["research_docs"]:
        raise IntegrityError("builder changed the frozen input manifest")
    if result.get("numbering_map"):
        if not callable(getattr(adapter, "renumber", None)):
            raise IntegrityError("numbering_map requires an explicit renumber(plan, candidate_path) stage")
        result = adapter.renumber(result, candidate)
        if not isinstance(result, dict):
            raise IntegrityError("renumber must return a resolved plan")
    if result.get("template") != plan["template"] or result.get("research_docs") != plan["research_docs"]:
        raise IntegrityError("numbering stage changed the frozen input manifest")
    # Mapping IDs describe final placement. Renumbering changes text, not unit
    # order; an adapter that changes structure must refresh targets itself.
    result["merged_docx"] = str(candidate.resolve())
    resolved = work_dir / "resolved-plan.json"
    write_json(resolved, result)
    manifest = result["review_comments"]
    if manifest.get("comments"):
        comment_path = work_dir / "comments.json"
        write_json(comment_path, manifest)
        commented = work_dir / "commented.docx"
        run_tool("apply_review_comments.py", candidate, comment_path, commented)
        commented.replace(candidate)
    for descriptor in [plan["template"], *plan["research_docs"]]:
        load_frozen(descriptor, base)
    report = check(resolved, candidate)
    write_json(work_dir / "machine-checks.json", report)
    if not report["passed"]:
        raise IntegrityError("candidate failed content checks; inspect machine-checks.json")
    write_json(work_dir / "review-gates.json", {"candidate_sha256": report["candidate_sha256"], "plan_sha256": digest(resolved.read_bytes()), "gates": {name: {"status": "pending", "evidence": []} for name in sorted(GATES)}})
    return {"status": "candidate_ready_for_review", "candidate": str(candidate), "plan": str(resolved), "candidate_sha256": report["candidate_sha256"]}


def verify_evidence(evidence, base):
    if not isinstance(evidence, list) or not evidence:
        raise IntegrityError("review evidence must name actual files and their SHA-256 values")
    for item in evidence:
        path = Path(item["file"])
        path = path if path.is_absolute() else base / path
        if not item.get("sha256") or digest(path.read_bytes()) != item["sha256"]:
            raise IntegrityError("missing/stale review evidence: " + str(path))


def release(plan_path, candidate, gates_path, output, mode):
    plan, base = read_plan(plan_path)
    if output.resolve() in {candidate.resolve(), *[load_frozen(d, base).path for d in [plan["template"], *plan["research_docs"]]]}:
        raise IntegrityError("release output must be distinct from source and candidate paths")
    if output.exists():
        raise IntegrityError("release never silently overwrites an existing file")
    report = check(plan_path, candidate)
    if not report["passed"]:
        raise IntegrityError("release content verification failed")
    reviews = json.loads(gates_path.read_text(encoding="utf-8"))
    sha = report["candidate_sha256"]
    if reviews.get("candidate_sha256") != sha:
        raise IntegrityError("review applies to an earlier candidate")
    if reviews.get("plan_sha256") != digest(Path(plan_path).read_bytes()):
        raise IntegrityError("review applies to an earlier merge plan")
    gates = reviews.get("gates", {})
    required = GATES | ({"strict_preserve"} if any(e.get("policy") == "strict_preserve" for e in plan["content_register"]) else set())
    for name in required:
        gate = gates.get(name, {})
        if gate.get("status") != "passed" or not gate.get("reviewed_by") or not gate.get("finding"):
            raise IntegrityError("uncompleted review gate: " + name)
        verify_evidence(gate.get("evidence"), gates_path.parent)
    visual = gates["visual_pages"]
    pages = visual.get("pages", [])
    if type(visual.get("page_count")) is not int or visual["page_count"] < 1 or [p.get("page") for p in pages] != list(range(1, visual["page_count"] + 1)):
        raise IntegrityError("visual review must cover every rendered page")
    for page in pages:
        if page.get("status") != "passed":
            raise IntegrityError("unreviewed/failed page")
        verify_evidence([page], gates_path.parent)
    comments = {c.get("id") for c in plan["review_comments"].get("comments", [])}
    for decision in plan.get("decisions", []):
        if decision.get("status") == "resolved":
            continue
        if mode == "final":
            raise IntegrityError("unresolved decision prevents final release")
        if decision.get("comment_id") not in comments:
            raise IntegrityError("open review decision lacks a real comment")
        if decision.get("risk") == "technical" and (decision.get("material_change_applied") is not False or decision.get("original_retained") is not True):
            raise IntegrityError("unresolved technical decision must retain source content")
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(candidate, output)
    return {"status": "released", "mode": mode, "output": str(output), "sha256": sha}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    init = sub.add_parser("inventory")
    init.add_argument("template", type=Path)
    init.add_argument("sources", nargs="+", type=Path)
    init.add_argument("--out", type=Path, required=True)
    catalog = sub.add_parser("catalog")
    catalog.add_argument("document", type=Path)
    catalog.add_argument("--out", type=Path, required=True)
    run = sub.add_parser("run")
    run.add_argument("plan", type=Path)
    run.add_argument("--builder", type=Path, required=True)
    run.add_argument("--work-dir", type=Path, required=True)
    audit = sub.add_parser("check")
    audit.add_argument("plan", type=Path)
    audit.add_argument("--candidate", type=Path, required=True)
    audit.add_argument("--out", type=Path, required=True)
    final = sub.add_parser("release")
    final.add_argument("plan", type=Path)
    final.add_argument("--candidate", type=Path, required=True)
    final.add_argument("--gates", type=Path, required=True)
    final.add_argument("--output", type=Path, required=True)
    final.add_argument("--mode", choices=["review", "final"], default="review")
    args = parser.parse_args()
    try:
        if args.command == "inventory":
            write_json(args.out, scaffold(args.template, args.sources))
            result = {"status": "inventory_created", "plan": str(args.out)}
        elif args.command == "catalog":
            package = Package(args.document)
            write_json(args.out, {"file": str(package.path), "sha256": package.sha256, "units": list(package.inventory().values())})
            result = {"status": "catalog_created", "catalog": str(args.out)}
        elif args.command == "run":
            result = build(args.plan.resolve(), args.builder, args.work_dir.resolve())
        elif args.command == "check":
            result = check(args.plan, args.candidate)
            write_json(args.out, result)
        else:
            result = release(args.plan, args.candidate, args.gates, args.output, args.mode)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result.get("passed", True) else 1
    except Exception as exc:
        print(json.dumps({"passed": False, "error_type": type(exc).__name__, "error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
