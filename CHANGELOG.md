# Changelog

All notable changes to the Kabuto MCQ Exam Pipeline are documented in this file.

## v1.3

- Added build-record folder support to `do3_export_bank_to_pyexam.py`
  - `--build-dir` writes versioned exam YAMLs, per-version manifests, answer keys, and a top-level build-set manifest
  - writes `selected_question_ids.txt` for traceability
- Added exam set generation to `do3_export_bank_to_pyexam.py`
  - `--num-versions N`
  - same selected questions, different randomisation seeds
  - auto-generates a base seed for sets if none is provided
- Updated `do4_build_pyexam.py` to support batch builds from a directory of exported exam YAML files
  - writes per-version build manifests
  - writes batch build manifest

## v1.2

- Added `active` column to canonical Excel schema (after `id`)
  - defaults to `1` on import
  - rows with `active == 1` are exported by default
- Added `do5_make_subset.py` (subset workbook creator)
- Added chapter-section filtering enhancements to `do3` and `do5`
  - `--section-prefix MTM3.*`
  - single-argument range syntax: `--section-prefix MTM1.*-MTM5.*`
  - two-argument inclusive range syntax: `--section-range MTM1.* MTM5.*`
  - supports chapter wildcards and exact section endpoints (e.g., `MTM3.2`)

## v1.1

- Added `do0_make_blank_bank.py` to create a blank Excel bank with canonical headers
- Expanded `do3_export_bank_to_pyexam.py` subsetting/filtering options:
  - `--chapter-section` (repeatable / comma-separated)
  - `--ids` (comma-separated ID list)
  - `--id-file` (newline/comma-separated IDs)
  - `--subset-xlsx` (subset workbook IDs, preserving subset workbook order)
- Added answer-key export in `do3`

## v1.0

- `do1_import_to_bank.py` (YAML intake → Excel bank with stable IDs + `import_batch`)
- `do2_validate_bank.py` (canonical Excel validation)
- `do3_export_bank_to_pyexam.py` (Excel → pyexam YAML with seeded randomisation + manifest)
- `do4_build_pyexam.py` (pyexam build wrapper with dry-run support)
- `kabuto_common.py` (shared schema, parsing, validation, export helpers)
