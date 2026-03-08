#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

from kabuto_common import (
    DEFAULT_SHEET_NAME,
    append_questions_to_sheet,
    assign_ids,
    fill_import_batch,
    load_intake_yaml_questions,
    load_or_create_workbook,
    now_batch_label,
    scan_existing_ids,
)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Import intake YAML into Excel question bank and assign stable IDs")
    p.add_argument("input_yaml", help="Path to intake YAML")
    p.add_argument("output_xlsx", help="Path to output Excel workbook (.xlsx)")
    p.add_argument("--mode", choices=["new", "append"], default="append", help="Workbook mode (default: append)")
    p.add_argument("--sheet", default=DEFAULT_SHEET_NAME, help=f"Worksheet name (default: {DEFAULT_SHEET_NAME})")
    p.add_argument(
        "--import-batch",
        default=None,
        help="Value for import_batch column (default: generated timestamp label)",
    )
    return p


def main() -> int:
    args = build_parser().parse_args()
    _, questions = load_intake_yaml_questions(args.input_yaml)
    if not questions:
        raise SystemExit("No questions found in input YAML")

    wb, ws = load_or_create_workbook(args.output_xlsx, mode=args.mode, sheet_name=args.sheet)
    existing_ids = scan_existing_ids(ws)
    assign_ids(questions, existing_ids)
    fill_import_batch(questions, args.import_batch or now_batch_label("import"))
    append_questions_to_sheet(ws, questions)

    out_path = Path(args.output_xlsx)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)

    print(f"Imported {len(questions)} questions -> {out_path}")
    print(f"Sheet: {ws.title}")
    print(f"First ID: {questions[0].qid}")
    print(f"Last ID:  {questions[-1].qid}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
