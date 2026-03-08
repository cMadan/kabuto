# Command examples

This file collects example command lines for the `do*` scripts (`do0`–`do5`).

## `do0_make_blank_bank.py`

Create a blank bank workbook with canonical headers:

```bash
python do0_make_blank_bank.py examples/blank_bank.xlsx
```

Overwrite an existing workbook:

```bash
python do0_make_blank_bank.py examples/blank_bank.xlsx --force
```

## `do1_import_to_bank.py`

Create a new bank from intake YAML:

```bash
python do1_import_to_bank.py examples/intake.yaml examples/curated_bank.xlsx --mode new --import-batch demo_batch
```

Append another intake file:

```bash
python do1_import_to_bank.py more_intake.yaml examples/curated_bank.xlsx --mode append --import-batch batch_02
```

## `do2_validate_bank.py`

Validation only (default):

```bash
python do2_validate_bank.py examples/curated_bank.xlsx
```

Validation + summary in one run:

```bash
python do2_validate_bank.py examples/curated_bank.xlsx --summary --show-sections 20 --json-out examples/bank_summary.json
```

Summary only (does not fail on validation errors):

```bash
python do2_validate_bank.py examples/curated_bank.xlsx --summary-only --show-sections 20
```

Treat warnings as errors (validation mode):

```bash
python do2_validate_bank.py examples/curated_bank.xlsx --warnings-as-errors
```

## `do3_export_bank_to_pyexam.py`

Basic export with seed and answer key:

```bash
python do3_export_bank_to_pyexam.py examples/curated_bank.xlsx examples/exam_for_pyexam.yaml --exam-name "MTM Chapter 3" --seed 123 --answer-key
```

Export filtered by exact chapter section:

```bash
python do3_export_bank_to_pyexam.py examples/curated_bank.xlsx examples/exam_mtm3_3.yaml --chapter-section MTM3.3 --seed 123 --allow-warnings
```

Export filtered by chapter wildcard/range:

```bash
python do3_export_bank_to_pyexam.py examples/curated_bank.xlsx examples/exam_mtm3_to5.yaml --section-prefix MTM3.*-MTM5.* --auto-seed --allow-warnings
```

Export a set of versions (same selected questions, different randomisation):

```bash
python do3_export_bank_to_pyexam.py examples/curated_bank.xlsx exam.yaml --build-dir builds/midterm_set --num-versions 4 --answer-key
```

Export using a subset workbook (IDs and order taken from subset workbook):

```bash
python do3_export_bank_to_pyexam.py examples/curated_bank.xlsx examples/exam_from_subset.yaml --subset-xlsx examples/subset_exam.xlsx --seed 123 --allow-warnings
```

## `do4_build_pyexam.py`

Build one exported exam to PDF:

```bash
python do4_build_pyexam.py examples/exam_for_pyexam.yaml builds/ch3_v1 --format pdf
```

Dry-run a pyexam command (debugging CLI invocation):

```bash
python do4_build_pyexam.py examples/exam_for_pyexam.yaml builds/ch3_v1 --dry-run --python-module-fallback
```

Build a whole set from a directory of exported YAMLs:

```bash
python do4_build_pyexam.py builds/midterm_set builds/midterm_pdf_set --python-module-fallback
```

## `do5_make_subset.py`

Create a subset workbook by chapter-section range (active rows only by default):

```bash
python do5_make_subset.py examples/curated_bank.xlsx examples/subset_mtm3_to5.xlsx --section-prefix MTM3.*-MTM5.* --allow-warnings
```

Create a subset workbook from an explicit ID list:

```bash
python do5_make_subset.py examples/curated_bank.xlsx examples/subset_selected.xlsx --ids mtm03_03_001,mtm03_03_005
```

Include inactive rows in subset selection:

```bash
python do5_make_subset.py examples/curated_bank.xlsx examples/subset_all.xlsx --section-prefix MTM3.* --include-inactive --allow-warnings
```
