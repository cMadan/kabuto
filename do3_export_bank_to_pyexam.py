#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import time
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

from kabuto_common import (
    dump_json,
    dump_yaml,
    load_excel_questions,
    matches_section_selector,
    chapter_section_in_range,
    questions_to_pyexam_yaml,
    randomize_questions_for_export,
    validate_questions,
)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description='Export curated Excel bank/subset to pyexam YAML')
    p.add_argument('input_xlsx', help='Path to curated bank workbook')
    p.add_argument('output_yaml', help='Output pyexam YAML path')
    p.add_argument('--sheet', default=None, help='Worksheet name for input_xlsx (default: active sheet)')
    p.add_argument('--exam-name', default='Exam', help="pyexam document 'name' field")
    p.add_argument('--header-file', default=None, help='Path to text file inserted as pyexam header (optional)')
    p.add_argument('--header-text', default=None, help='Header text literal (optional; ignored if --header-file is used)')
    p.add_argument('--seed', type=int, default=None, help='Seed for question/option randomisation')
    p.add_argument('--auto-seed', action='store_true', help='Generate a seed automatically when --seed is not provided')
    p.add_argument('--no-shuffle-questions', action='store_true', help='Preserve row order from Excel (default shuffles questions)')
    p.add_argument('--no-shuffle-options', action='store_true', help='Do not shuffle any options (ignores shuffle_mode=random)')
    p.add_argument('--show-ids', action='store_true', help='Prefix each question stem with [id] in exported text')
    p.add_argument('--manifest', default=None, help='Path for export manifest JSON (default: <output_yaml>.manifest.json)')
    p.add_argument('--allow-warnings', action='store_true', help='Continue export if validator warnings are present')

    # Subsetting / filtering
    p.add_argument(
        '--chapter-section',
        action='append',
        default=[],
        help='Include only rows matching this chapter_section (repeatable; comma-separated values also allowed)',
    )
    p.add_argument(
        '--section-prefix',
        action='append',
        default=[],
        help=(
            'Include rows whose chapter_section matches a prefix/wildcard/range (repeatable; comma-separated also allowed). '
            'Examples: MTM3.*, MTM3., MTM1.*-MTM5.*'
        ),
    )
    p.add_argument(
        '--section-range',
        nargs=2,
        action='append',
        metavar=('START', 'END'),
        default=[],
        help='Include rows with chapter_section in inclusive range (e.g., MTM1.* MTM5.* or MTM3.2 MTM4.4)',
    )
    p.add_argument('--ids', default=None, help='Comma-separated question IDs to include (preserves this order)')
    p.add_argument('--id-file', default=None, help='Text file with question IDs to include (newline or comma separated)')
    p.add_argument('--subset-xlsx', default=None, help='Excel subset workbook whose IDs define inclusion and order')
    p.add_argument('--subset-sheet', default=None, help='Worksheet name for --subset-xlsx (default: active sheet)')

    # Answer key outputs
    p.add_argument('--answer-key-csv', default=None, help='Write answer key CSV (default: alongside YAML when --answer-key is used)')
    p.add_argument('--answer-key-md', default=None, help='Write answer key Markdown table (PDF-friendly intermediate)')
    p.add_argument('--answer-key', action='store_true', help='Write default answer-key CSV and Markdown alongside output YAML')
    return p


def _read_header(args: argparse.Namespace) -> Optional[str]:
    if args.header_file:
        return Path(args.header_file).read_text(encoding='utf-8')
    if args.header_text:
        return args.header_text
    return None


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
    # allow comma and/or newline separated
    parts = []
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
    notes['active_only'] = True
    notes['active_rows_loaded'] = len(filtered)

    # chapter_section exact filter (can be combined)
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
    if args.subset_xlsx:
        subset_questions = load_excel_questions(args.subset_xlsx, sheet_name=args.subset_sheet, active_only=True)
        subset_errors, subset_warnings = validate_questions(subset_questions, require_ids=True)
        if subset_errors:
            raise ValueError('subset workbook has validation errors: ' + '; '.join(subset_errors[:10]))
        ordered_ids.extend([q.qid for q in subset_questions if q.qid])
        notes['subset_xlsx'] = str(Path(args.subset_xlsx))
        notes['subset_sheet'] = args.subset_sheet
        notes['subset_warning_count'] = len(subset_warnings)

    if ordered_ids:
        selected, missing = _order_and_filter_by_ids(filtered, ordered_ids)
        notes['id_filter_count'] = len(ordered_ids)
        notes['id_filter_unique_count'] = len({x for x in ordered_ids})
        notes['id_filter_missing'] = missing
        filtered = selected

    return filtered, notes


def _answer_key_rows(randomized_questions, rand_manifest: dict) -> List[dict]:
    rows = []
    perms = rand_manifest.get('option_permutations', {}) if rand_manifest else {}
    for i, q in enumerate(randomized_questions, start=1):
        perm_info = perms.get(q.qid or q.stem[:20], {})
        rows.append(
            {
                'order': i,
                'id': q.qid or '',
                'chapter_section': q.chapter_section or '',
                'answer': q.answer,
                'points': q.points,
                'shuffle_mode': q.shuffle_mode,
                'original_answer': perm_info.get('original_answer', q.answer),
                'exported_answer': perm_info.get('new_answer', q.answer),
                'stem': q.stem,
                'explanation': q.explanation or '',
            }
        )
    return rows


def _write_answer_key_csv(path: str | Path, rows: List[dict]) -> None:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = ['order', 'id', 'chapter_section', 'answer', 'points', 'shuffle_mode', 'original_answer', 'exported_answer', 'stem', 'explanation']
    with open(p, 'w', encoding='utf-8', newline='') as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)


def _write_answer_key_md(path: str | Path, rows: List[dict], exam_name: str, seed: Optional[int]) -> None:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    lines = [f'# Answer Key: {exam_name}', '']
    if seed is not None:
        lines.append(f'- Seed: `{seed}`')
        lines.append('')
    lines += [
        '| # | ID | Section | Answer | Points |',
        '|---:|---|---|:---:|---:|',
    ]
    for r in rows:
        lines.append(
            f"| {r['order']} | {r['id']} | {r['chapter_section']} | {r['answer']} | {r['points']} |"
        )
    lines.append('')
    lines.append('## Detailed explanations')
    lines.append('')
    for r in rows:
        lines.append(f"### {r['order']}. {r['id']} ({r['chapter_section']})")
        lines.append(f"- Correct answer: **{r['answer']}**")
        lines.append(f"- Stem: {r['stem']}")
        if r.get('explanation'):
            lines.append(f"- Explanation: {r['explanation']}")
        lines.append('')
    p.write_text('\n'.join(lines), encoding='utf-8')


def main() -> int:
    args = build_parser().parse_args()
    questions = load_excel_questions(args.input_xlsx, sheet_name=args.sheet, active_only=True)
    errors, warnings = validate_questions(questions, require_ids=True)
    if errors:
        for e in errors:
            print(f'ERROR: {e}')
        return 1
    if warnings and not args.allow_warnings:
        for w in warnings:
            print(f'WARNING: {w}')
        print('Refusing export due to warnings (use --allow-warnings to override)')
        return 2

    try:
        questions, subset_notes = _filter_questions(questions, args)
    except Exception as exc:
        print(f'ERROR: {exc}')
        return 3

    if not questions:
        print('ERROR: no questions selected for export after filters')
        return 4

    seed = args.seed
    if seed is None and args.auto_seed:
        seed = int(time.time())

    randomized_questions, rand_manifest = randomize_questions_for_export(
        questions,
        seed=seed,
        shuffle_questions=not args.no_shuffle_questions,
        shuffle_options=not args.no_shuffle_options,
    )

    payload = questions_to_pyexam_yaml(
        randomized_questions,
        exam_name=args.exam_name,
        header=_read_header(args),
        include_points=True,
        include_ids_in_text=args.show_ids,
    )

    out_yaml = Path(args.output_yaml)
    out_yaml.parent.mkdir(parents=True, exist_ok=True)
    dump_yaml(payload, out_yaml)

    manifest_path = Path(args.manifest) if args.manifest else out_yaml.with_suffix(out_yaml.suffix + '.manifest.json')
    manifest = {
        'input_xlsx': str(Path(args.input_xlsx)),
        'sheet': args.sheet,
        'output_yaml': str(out_yaml),
        'exam_name': args.exam_name,
        'seed': seed,
        'warnings': warnings,
        'subset_filter': subset_notes,
        'randomisation': rand_manifest,
        'question_ids': [q.qid for q in randomized_questions],
    }
    dump_json(manifest, manifest_path)

    # Answer key outputs
    if args.answer_key or args.answer_key_csv or args.answer_key_md:
        rows = _answer_key_rows(randomized_questions, rand_manifest)
        ak_csv = Path(args.answer_key_csv) if args.answer_key_csv else out_yaml.with_suffix('.answer_key.csv')
        ak_md = Path(args.answer_key_md) if args.answer_key_md else out_yaml.with_suffix('.answer_key.md')
        if args.answer_key or args.answer_key_csv:
            _write_answer_key_csv(ak_csv, rows)
            print(f'Answer key CSV -> {ak_csv}')
        if args.answer_key or args.answer_key_md:
            _write_answer_key_md(ak_md, rows, args.exam_name, seed)
            print(f'Answer key Markdown -> {ak_md}')

    print(f'Exported {len(randomized_questions)} questions -> {out_yaml}')
    print(f'Manifest -> {manifest_path}')
    if seed is not None:
        print(f'Seed: {seed}')
    if subset_notes:
        print(f'Subset/filter applied: {subset_notes}')
    if warnings:
        print(f'Warnings present: {len(warnings)}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
