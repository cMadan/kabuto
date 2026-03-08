#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import load_workbook

from kabuto_common import ensure_sheet_headers, excel_row_to_question, validate_questions


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Validate curated Excel MCQ bank")
    p.add_argument("input_xlsx", help="Path to Excel workbook")
    p.add_argument("--sheet", default=None, help="Worksheet name (default: active sheet)")
    p.add_argument("--warnings-as-errors", action="store_true", help="Return nonzero exit code if warnings exist")
    return p


def main() -> int:
    args = build_parser().parse_args()
    wb = load_workbook(args.input_xlsx)
    ws = wb[args.sheet] if (args.sheet and args.sheet in wb.sheetnames) else wb.active
    ensure_sheet_headers(ws)

    questions = []
    structural_errors = []
    for r in range(2, ws.max_row + 1):
        try:
            q = excel_row_to_question(ws, r)
        except Exception as exc:
            structural_errors.append(f"Row {r}: {exc}")
            continue
        if q is not None:
            questions.append(q)

    errors, warnings = validate_questions(questions, require_ids=True)
    errors = structural_errors + errors

    print(f"File: {Path(args.input_xlsx)}")
    print(f"Sheet: {ws.title}")
    print(f"Rows parsed: {len(questions)}")
    print(f"Errors: {len(errors)}")
    print(f"Warnings: {len(warnings)}")
    if errors:
        print("\nERRORS")
        for e in errors:
            print(f" - {e}")
    if warnings:
        print("\nWARNINGS")
        for w in warnings:
            print(f" - {w}")

    if errors:
        return 1
    if warnings and args.warnings_as_errors:
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
