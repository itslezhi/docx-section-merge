#!/usr/bin/env python3
"""Verify complete frozen-source coverage with merge plan v3."""
import argparse
import json
import sys
from pathlib import Path
from merge_integrity import verify


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("plan", type=Path)
    parser.add_argument("--candidate", type=Path)
    parser.add_argument("--json", dest="json_path", type=Path)
    args = parser.parse_args()
    try:
        report = verify(args.plan, args.candidate)
    except Exception as exc:
        report = {"passed": False, "errors": [str(exc)], "error_type": type(exc).__name__}
    payload = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
    if args.json_path:
        args.json_path.parent.mkdir(parents=True, exist_ok=True)
        args.json_path.write_text(payload, encoding="utf-8")
    print(payload, end="")
    return 0 if report["passed"] else 1


if __name__ == "__main__":
    sys.exit(main())
