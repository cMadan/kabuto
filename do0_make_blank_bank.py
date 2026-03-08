#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

from kabuto_common import DEFAULT_SHEET_NAME, create_workbook_with_headers


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description='Create a blank Kabuto Excel question bank with canonical headers')
    p.add_argument('output_xlsx', help='Path to output workbook (.xlsx)')
    p.add_argument('--sheet', default=DEFAULT_SHEET_NAME, help=f'Worksheet name (default: {DEFAULT_SHEET_NAME})')
    p.add_argument('--force', action='store_true', help='Overwrite existing file')
    return p


def main() -> int:
    args = build_parser().parse_args()
    out = Path(args.output_xlsx)
    if out.exists() and not args.force:
        print(f"ERROR: {out} already exists (use --force to overwrite)")
        return 1
    out.parent.mkdir(parents=True, exist_ok=True)
    wb = create_workbook_with_headers(args.sheet)
    wb.save(out)
    print(f'Created blank bank: {out}')
    print(f'Sheet: {args.sheet}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
