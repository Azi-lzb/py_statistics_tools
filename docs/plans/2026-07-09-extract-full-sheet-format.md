# Full Sheet Format Preservation Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Preserve the complete Excel worksheet layout when 7-6 extracts a whole sheet.

**Architecture:** The COM path first attempts a native worksheet copy and verifies that the destination workbook gained a populated sheet. If Excel cannot perform that cross-workbook operation, the existing UsedRange copy remains as a fallback, augmented with row heights and column widths. Formula cells are converted to their displayed results after copying.

**Tech Stack:** Python, pywin32 Excel COM, openpyxl, pytest, PyInstaller.

---

### Task 1: Add a format regression test

**Files:**
- Modify: `pytools/tests/test_extract_sheets_76.py`

1. Create a workbook containing custom column widths, row heights, merged cells, styles, and a formula.
2. Extract it as a whole sheet.
3. Assert that layout and styles survive and that the formula is stored as a value.
4. Run the focused test and verify that it fails for the missing COM layout-copy behavior.

### Task 2: Implement reliable whole-sheet copying

**Files:**
- Modify: `pytools/core/feature_1_6_extract_sheets_com.py`

1. Attempt Excel's native `Worksheet.Copy` into the output workbook.
2. Verify that the output workbook sheet count increased and the copied sheet contains data.
3. Rename the copied sheet and convert formulas to values.
4. If native copy fails, copy UsedRange and explicitly transfer column widths and row heights.
5. Preserve the existing value-only fallback when both formatted copy paths fail.

### Task 3: Verify and package

**Files:**
- No source changes expected.

1. Run the focused extraction tests.
2. Compile the normal and Python 3.7 code paths.
3. Validate against the real `.xls` sample, checking sheet count, data, widths, and heights.
4. Build the normal and Win7 executables.
