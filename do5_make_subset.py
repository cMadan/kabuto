#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path
from typing import Dict, Iterable, List, Sequence, Tuple

from kabuto_common import (
    DEFAULT_SHEET_NAME,
    append_questions_to_sheet,
    chapter_section_in_range,
    create_workbook_with_headers,
    load_excel_questions,
    matches_section_selector,
    validate_questions,
)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description='Create a subset workbook from a curated Kabuto bank (active rows only by default)')
    p.add_argument('input_xlsx', help='Path to curated bank workbook')
    p.add_argument('output_xlsx', help='Path to subset workbook to create')
    p.add_argument('--sheet', default=None, help='Worksheet name for input_xlsx (default: active sheet)')
    p.add_argument('--output-sheet', default=DEFAULT_SHEET_NAME, help=f'Worksheet name for output workbook (default: {DEFAULT_SHEET_NAME})')
    p.add_argument('--include-inactive', action='store_true', help='Include rows with active!=1 (default excludes them)')
    p.add_argument('--force', action='store_true', help='Overwrite output workbook if it exists')
    p.add_argument('--allow-warnings', action='store_true', help='Continue if validator warnings are present')

    p.add_argument(
        '--chapter-section',
        action='append',
        default=[],
        help='Include only rows matching this chapter_section exactly (repeatable; comma-separated also allowed)',
    )
    p.add_argument(
        '--section-prefix',
        action='append',
        default=[],
        help='Prefix/wildcard/range selector (repeatable; comma-separated also allowed), e.g. MTM3.* or MTM1.*-MTM5.*',
    )
    p.add_argument(
        '--section-range',
        nargs=2,
        action='append',
        metavar=('START', 'END'),
        default=[],
        help='Inclusive chapter_section range, e.g. --section-range MTM1.* MTM5.*',
    )
    p.add_argument('--ids', default=None, help='Comma-separated question IDs to include (preserves this order)')
    p.add_argument('--id-file', default=None, help='Text file with question IDs to include (newline or comma separated)')
    return p


def _split_csvish(values: Iterable[str]) -> List[str]:
    out: List[str] = []
    for v in values:
        for part in str(v).split(','):
            s = part.strip()
            if s:
                out.append(s)
    return out


def _read_ids_from_file(path: str | Path) -> List[str]:
    txt = Path(path).read_text(encoding='utf-8')
    parts: List[str] = []
    for line in txt.splitlines():
        parts.extend(line.split(','))
    return [p.strip() for p in parts if p.strip()]


def _order_and_filter_by_ids(questions, ordered_ids: Sequence[str]) -> Tuple[List, List[str]]:
    by_id: Dict[str, object] = {}
    for q in questions:
        if q.qid:
            by_id[q.qid] = q
    selected = []
    missing = []
    seen = set()
    for qid in ordered_ids:
        if qid in seen:
            continue
        seen.add(qid)
        q = by_id.get(qid)
        if q is None:
            missing.append(qid)
        else:
            selected.append(q)
    return selected, missing


def _filter_questions(questions, args: argparse.Namespace):
    notes: Dict[str, object] = {}
    filtered = list(questions)
    notes['include_inactive'] = bool(args.include_inactive)
    notes['rows_loaded'] = len(filtered)

    chapter_sections = set(_split_csvish(args.chapter_section))
    if chapter_sections:
        before = len(filtered)
        filtered = [q for q in filtered if (q.chapter_section or '') in chapter_sections]
        notes['chapter_sections'] = sorted(chapter_sections)
        notes['chapter_section_filter_before'] = before
        notes['chapter_section_filter_after'] = len(filtered)

    section_selectors = _split_csvish(args.section_prefix)
    if section_selectors:
        before = len(filtered)
        filtered = [q for q in filtered if any(matches_section_selector(q.chapter_section, s) for s in section_selectors)]
        notes['section_prefix'] = section_selectors
        notes['section_prefix_filter_before'] = before
        notes['section_prefix_filter_after'] = len(filtered)

    if args.section_range:
        before = len(filtered)
        ranges = [(a, b) for a, b in args.section_range]
        filtered = [
            q for q in filtered
            if any(chapter_section_in_range(q.chapter_section, start, end) for start, end in ranges)
        ]
        notes['section_range'] = ranges
        notes['section_range_filter_before'] = before
        notes['section_range_filter_after'] = len(filtered)

    ordered_ids: List[str] = []
    if args.ids:
        ordered_ids.extend(_split_csvish([args.ids]))
    if args.id_file:
        ordered_ids.extend(_read_ids_from_file(args.id_file))
    if ordered_ids:
        selected, missing = _order_and_filter_by_ids(filtered, ordered_ids)
        notes['id_filter_count'] = len(ordered_ids)
        notes['id_filter_missing'] = missing
        filtered = selected

    return filtered, notes


def main() -> int:
    args = build_parser().parse_args()
    out = Path(args.output_xlsx)
    if out.exists() and not args.force:
        print(f'ERROR: {out} already exists (use --force to overwrite)')
        return 1

    questions = load_excel_questions(args.input_xlsx, sheet_name=args.sheet, active_only=not args.include_inactive)
    errors, warnings = validate_questions(questions, require_ids=True)
    if errors:
        for e in errors:
            print(f'ERROR: {e}')
        return 2
    if warnings and not args.allow_warnings:
        for w in warnings:
            print(f'WARNING: {w}')
        print('Refusing subset creation due to warnings (use --allow-warnings to override)')
        return 3

    try:
        subset, notes = _filter_questions(questions, args)
    except Exception as exc:
        print(f'ERROR: {exc}')
        return 4

    if not subset:
        print('ERROR: no questions selected for subset workbook')
        return 5

    wb = create_workbook_with_headers(args.output_sheet)
    ws = wb[args.output_sheet]
    append_questions_to_sheet(ws, subset)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)

    print(f'Subset workbook created: {out}')
    print(f'Sheet: {ws.title}')
    print(f'Questions: {len(subset)}')
    if notes:
        print(f'Filters: {notes}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
