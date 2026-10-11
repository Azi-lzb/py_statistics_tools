# Validation Rule Generator Implementation Plan

**Goal:** Generate formal validation rules from report sheets and human-readable validation relationships, with D/F/J caliber matching and explicit failure output.

**Architecture:** Add a dedicated core module for workbook discovery, relation parsing and output generation. Keep runtime options in a new config Sheet, expose the feature under menu 6, and use focused openpyxl tests for the supplied same-table example and a cross-table missing-caliber case.

**Tech Stack:** Python 3.7+, openpyxl, pytest

## Task 1: Regression tests

- Create `pytools/tests/test_validation_rule_generator.py`.
- Verify the supplied `1=1.1+1.2` example expands to A/B/C using `S0801_J_26`.
- Verify variable header rows are detected.
- Verify cross-table D/F/J matching generates available calibers and reports a missing caliber.
- Verify config initialization creates the new Sheet and comments.

## Task 2: Core generator

- Create `pytools/core/validation_rule_generator.py`.
- Discover report metadata and mappings.
- Parse local and cross-table references.
- Generate valid rows and failure rows.
- Write `校验规则` and `未生成明细` output sheets.

## Task 3: Config and menu integration

- Add config constants, columns, validation and loader.
- Add initialization comments and one example row.
- Add menu `6-3` and interactive workbook selection.

## Task 4: Verification and sample output

- Run focused tests and compile checks.
- Run the generator against `xlsx/表格.xlsx` and `xlsx/表格校验.xlsx`.
- Inspect the resulting rules and failure sheet.
