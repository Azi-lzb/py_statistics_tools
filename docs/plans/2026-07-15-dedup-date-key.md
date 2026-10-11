# Date-only Deduplication Key Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Make 2-5 treat date strings and Excel date-time values on the same calendar day as the same deduplication key.

**Architecture:** Centralize the behavior in `dedup_311._norm`, which is already used by duplicate checking, deletion, existing-target key loading, and append filtering. Normalize only values that are dates or unambiguous full-date strings; preserve the existing string behavior for all other values.

**Tech Stack:** Python, datetime, openpyxl, pytest

---

### Task 1: Add date-key regression coverage

**Files:**
- Create: `pytools/tests/test_dedup_311.py`

**Step 1: Write the failing tests**

Add tests proving that `2026-01-31`, a midnight `datetime`, and a non-midnight `datetime` normalize to the same key, and that 2-5 filters a source row when its target key contains the same date in another representation.

**Step 2: Run test to verify it fails**

Run: `C:\Users\Azi\miniconda3\envs\HZPBC\python.exe -m pytest pytools\tests\test_dedup_311.py -v`

Expected: FAIL because `_norm` currently returns the time component for `datetime` values.

### Task 2: Normalize dedup dates by calendar day

**Files:**
- Modify: `pytools/core/dedup_311.py`
- Test: `pytools/tests/test_dedup_311.py`

**Step 1: Implement the minimal normalization**

For `date`/`datetime`, return `YYYY-MM-DD`. For strings, normalize only when the whole trimmed value is a supported full-date or date-time representation. Otherwise retain `str(value).strip()`.

**Step 2: Run focused tests**

Run: `C:\Users\Azi\miniconda3\envs\HZPBC\python.exe -m pytest pytools\tests\test_dedup_311.py -v`

Expected: PASS.

**Step 3: Run regression checks**

Run: `C:\Users\Azi\miniconda3\envs\HZPBC\python.exe -m pytest pytools\tests\test_e2e.py pytools\tests\test_date_excel_format.py -v`

Expected: PASS.

Run: `C:\Users\Azi\miniconda3\envs\HZPBC\python.exe -m py_compile pytools\core\dedup_311.py`

Expected: exit code 0.
