# Kabuto MCQ Exam Pipeline

A small pipeline for managing MCQ questions in Excel and generating student-facing exam PDFs.

It is designed to keep question management simple and under your control:

- intake in minimal structured YAML (schema shown below)
- curation in Excel (canonical source)
- export to `pyexam` YAML
- render final PDF via `pyexam` (via LaTeX)

## Pipeline overview

0. `do0_make_blank_bank.py`  
   Create a blank Excel bank with the canonical header row.

1. `do1_import_to_bank.py`  
   Import minimal/draft YAML into an Excel question bank and assign stable IDs.

2. `do2_validate_bank.py`  
   Validate the curated Excel bank (schema + content checks).

3. `do3_export_bank_to_pyexam.py`  
   Export curated questions from Excel to `pyexam` YAML (with optional seeded randomisation).

4. `do4_build_pyexam.py`  
   Run `pyexam` to generate LaTeX/PDF outputs.

## Canonical Excel schema

Columns:

- `import_batch` *(optional; provenance label, first column by design)*
- `id` *(required after import; generated on first YAML→Excel import)*
- `chapter_section` *(optional; e.g., `MTM3.3`)*
- `stem` *(required)*
- `opt_a` *(required if used; at least 2 total options required across A–E)*
- `opt_b`
- `opt_c`
- `opt_d`
- `opt_e` *(rarely used)*
- `answer` *(required; one of `A`–`E`, must point to a non-blank option)*
- `shuffle_mode` *(optional; `random` or `none`; blank defaults to `random`)*
- `points` *(optional; blank defaults to `1`)*
- `explanation` *(optional; instructor-facing)*

Notes:

- MCQ only (single correct answer)
- Maximum 5 options
- Excel is the curated source of truth

## ID generation

IDs are generated from `chapter_section` using zero-padded chapter/section numbers.

Examples:

- `MTM3.3` → `mtm03_03_001`
- `MTM12.10` → `mtm12_10_001`

Rules:

- chapter padded to 2 digits
- section padded to 2 digits
- item index padded to 3 digits
- if `chapter_section` is missing/invalid, fallback prefix is `unassigned` (e.g., `unassigned_001`)

## Intake YAML format (minimal structured input)

Supported per-question fields:

- `chapter_section` (optional)
- `stem` (required)
- `options` (required, 2–5 items; typically 4)
- `answer` (required, `A`–`E`)
- `shuffle_mode` (optional: `random` / `none`)
- `points` (optional)
- `explanation` (optional)

## Validation (what `do2_...` checks)

Hard errors:

- missing `id` (in curated Excel)
- missing `stem`
- fewer than 2 or more than 5 options
- gapped options (e.g., A/B/D with blank C)
- invalid `answer`
- answer points to blank option
- invalid `shuffle_mode`
- duplicate `id`

Warnings:

- duplicate option text within a question
- likely letter-referential wording with `shuffle_mode=random` (e.g., “Both A and B”)
- missing `chapter_section`

## Randomisation behaviour

- `shuffle_mode=random` → options may be shuffled
- `shuffle_mode=none` → option order is preserved

Question-order randomisation is handled during export/build (seeded if enabled).

## Repository layout

```text
kabuto/
├─ README.md
├─ requirements.txt
├─ kabuto_common.py
├─ do0_make_blank_bank.py
├─ do1_import_to_bank.py
├─ do2_validate_bank.py
├─ do3_export_bank_to_pyexam.py
├─ do4_build_pyexam.py
├─ examples/
│  ├─ intake.yaml
│  ├─ curated_bank.xlsx
│  └─ exam_for_pyexam.yaml
└─ builds/
   └─ ...
```

## Workflow

1. Import minimal YAML into a new or existing Excel bank.
2. Curate questions in Excel (edit wording, set `shuffle_mode`, assign `chapter_section`, trim to an exam subset).
3. Validate the bank/subset.
4. Export to `pyexam` YAML (optionally filter by `chapter_section`, ID list, or a subset workbook; optionally emit answer-key CSV/Markdown).
5. Build final PDF with `pyexam`.