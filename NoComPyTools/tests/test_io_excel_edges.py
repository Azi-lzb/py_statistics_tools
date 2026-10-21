from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from types import SimpleNamespace

import pandas as pd
from openpyxl import Workbook, load_workbook

from nocompytools.core import io_excel


def _book(path: Path, sheets: dict[str, list[list[object]]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for row in rows:
            ws.append(row)
    wb.save(path)
    wb.close()


def test_date_coercion_key_normalization_and_excel_engine_selection() -> None:
    original = pd.DataFrame({"value": ["x"]})
    assert io_excel._coerce_data_date_columns(original) is original
    converted = io_excel._coerce_data_date_columns(pd.DataFrame({"数据日期": ["2026.04", "unknown"]}))
    assert converted.loc[0, "数据日期"] == date(2026, 4, 30)
    assert converted.loc[1, "数据日期"] == "unknown"

    values = pd.DataFrame({"key": [
        None, float("nan"), pd.Timestamp("2026-01-02"), pd.Timestamp("2026-01-02 03:04:05"),
        datetime(2026, 1, 3), date(2026, 1, 4), True, 1, 1.0, 1.25, "  text  ",
    ]})
    assert io_excel._to_text_key_df(values)["key"].tolist() == [
        "", "", "2026-01-02", "2026-01-02 03:04:05", "2026-01-03", "2026-01-04",
        "True", "1", "1", "1.25", "text",
    ]
    list_value = pd.DataFrame({"key": pd.Series([["x"]], dtype=object)})
    assert io_excel._to_text_key_df(list_value).iloc[0, 0] == "['x']"
    assert io_excel._excel_engine_for_path("a.xlsx") == "calamine"
    assert io_excel._excel_engine_for_path("a.xlsm") == "calamine"
    assert io_excel._excel_engine_for_path("a.xls") == "calamine"
    assert io_excel._excel_engine_for_path("a.csv") is None


def test_print_area_formatting_and_sheet_names_handle_duplicates() -> None:
    class FakeSheet:
        max_row = 3

        def __init__(self):
            self.formats = {}

        def cell(self, row, col):
            return SimpleNamespace(number_format=self.formats.get((row, col)))

    ws = FakeSheet()
    writer = SimpleNamespace(sheets={"Data": ws})
    io_excel._apply_data_date_format(writer, "Data", pd.DataFrame({"id": [1], "数据日期": [date(2026, 1, 1)]}))
    assert ws.formats == {}
    # Fake cells need stable objects so number formats written by the helper can be inspected.
    cells = {}
    ws.cell = lambda row, col: cells.setdefault((row, col), SimpleNamespace(number_format="General"))
    io_excel._apply_data_date_format(writer, "Data", pd.DataFrame({"id": [1], "数据日期": [date(2026, 1, 1)]}))
    assert cells[(2, 2)].number_format == "yyyy/m/d" and cells[(3, 2)].number_format == "yyyy/m/d"
    io_excel._apply_data_date_format(writer, "missing", pd.DataFrame({"数据日期": [date(2026, 1, 1)]}))
    io_excel._apply_data_date_format(writer, "Data", pd.DataFrame({"other": [1]}))

    used: set[str] = set()
    first = io_excel.safe_sheet_name("bad/name?", used)
    second = io_excel.safe_sheet_name("bad:name?", used)
    third = io_excel.safe_sheet_name("bad_name_2", used)
    repeated = io_excel.safe_sheet_name("bad/name?", used)
    long_name = io_excel.safe_sheet_name("x" * 40, used)
    assert first == "bad_name_"
    assert second == "bad_name__2"
    assert third == "bad_name_2"
    assert repeated == "bad_name__3"
    assert len(long_name) == 31
    assert io_excel.safe_sheet_name("", used) == "Sheet"


def test_workbook_io_reads_caches_clears_and_lists_supported_sources(tmp_path: Path) -> None:
    path = tmp_path / "nested" / "source.xlsx"
    _book(path, {"Data": [["x", 1]], "Other": [["y", 2]]})
    assert io_excel.list_sheet_names(path) == ["Data", "Other"]
    frame = io_excel.read_sheet_2d(path, "Data")
    assert frame.iloc[0].tolist() == ["x", 1]
    assert io_excel.read_sheet_2d(tmp_path / "missing.xlsx", "Data").empty
    assert io_excel.list_sheet_names(tmp_path / "missing.xlsx") == []

    listed = [path.parent / name for name in ("a.xlsx", "B.XLSX", "c.txt", "~$locked.xlsx")]
    for item in listed:
        item.write_text("x", encoding="utf-8")
    assert io_excel.list_source_files(path.parent, [".xlsx"]) == sorted([path, listed[0], listed[1]])
    assert io_excel.list_source_files(tmp_path / "absent", [".xlsx"]) == []

    io_excel.clear_sheet_cache()
    refreshed = io_excel.read_sheet_2d(path, "Other")
    assert refreshed.iloc[0].tolist() == ["y", 2]


def test_write_workbook_sanitizes_duplicate_sheet_names(tmp_path: Path) -> None:
    path = tmp_path / "nested" / "result.xlsx"
    io_excel.write_workbook(path, {
        "bad/name": pd.DataFrame({"x": [1]}),
        "bad:name": pd.DataFrame({"x": [2]}),
        "bad_name_2": pd.DataFrame({"x": [3]}),
    })
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        assert wb.sheetnames == ["bad_name", "bad_name_2", "bad_name_2_2"]
        assert wb["bad_name_2_2"]["A2"].value == 3
    finally:
        wb.close()


def test_append_to_target_handles_empty_batch_and_batch_dedup(tmp_path: Path) -> None:
    target = tmp_path / "target.xlsx"
    empty = io_excel.append_to_target(target, "Rows", ["id"], [], [0])
    assert empty["input_rows"] == 0 and not target.exists()

    result = io_excel.append_to_target(
        target, "Rows", ["id", "value"], [[1, "a"], [1.0, "duplicate"], [2, "b"]], [0]
    )
    assert result["input_rows"] == 3 and result["batch_dedup_rows"] == 2
    assert result["written"] == 1
    frame = pd.read_excel(target, sheet_name="Rows")
    assert frame.to_dict("records") == [{"id": 1, "value": "a"}, {"id": 2, "value": "b"}]

    all_duplicates = io_excel.append_to_target(target, "Rows", ["id"], [[1], [1.0]], [0])
    assert all_duplicates["input_rows"] == 2 and all_duplicates["batch_dedup_rows"] == 1
    assert all_duplicates["written"] == 0


def test_append_to_target_filters_existing_rows_and_checks_header_prefix(tmp_path: Path) -> None:
    target = tmp_path / "target.xlsx"
    _book(target, {"Rows": [["id", "value"], [1, "old"]], "Keep": [["other"], ["kept"]]})
    result = io_excel.append_to_target(
        target, "Rows", ["id", "value"], [[1, "duplicate"], [2, "new"]], [0],
        required_prefix=["id"],
    )
    assert result["written"] == 1 and result["existing_filtered_rows"] == 1
    wb = load_workbook(target, data_only=True)
    try:
        assert wb.sheetnames == ["Rows", "Keep"]
        assert wb["Rows"].max_row == 3 and wb["Keep"]["A2"].value == "kept"
    finally:
        wb.close()

    wrong_header = tmp_path / "wrong.xlsx"
    _book(wrong_header, {"Rows": [["other"], ["safe"]]})
    skipped = io_excel.append_to_target(
        wrong_header, "Rows", ["id", "value"], [[2, "new"]], [0], required_prefix=["id"]
    )
    assert skipped["written"] == 0 and skipped["skipped_reason"] == "header_mismatch"


def test_append_to_target_supports_column_extension_and_sequence_dedup(tmp_path: Path) -> None:
    target = tmp_path / "target.xlsx"
    _book(target, {"Rows": [["group", "id", "value"], ["A", "x", 1], ["A", "x", 2]]})
    result = io_excel.append_to_target(
        target, "Rows", ["group", "id", "value", "extra"],
        [["A", "x", 1, "dup"], ["A", "x", 2, "second"], ["A", "x", 3, "third"]],
        [0, 1], required_prefix=["group", "id"], allow_extend=True, dedup_seq_prefix_idx=[0],
    )
    assert result["written"] == 1 and result["existing_filtered_rows"] == 1
    frame = pd.read_excel(target, sheet_name="Rows", dtype=object)
    assert list(frame.columns) == ["group", "id", "value", "extra"]
    assert frame["value"].tolist() == [1, 2, 3]
    assert frame["extra"].tolist()[-1] == "third"


def test_append_to_target_creates_missing_sheet_and_accepts_empty_key_columns(tmp_path: Path) -> None:
    target = tmp_path / "target.xlsx"
    _book(target, {"Keep": [["value"], ["keep"]]})
    result = io_excel.append_to_target(target, "Added", ["a", "b"], [[1, 2], [1, 2]], [])
    assert result["written"] == 1 and result["batch_dedup_rows"] == 2
    wb = load_workbook(target, data_only=True)
    try:
        assert wb.sheetnames == ["Keep", "Added"]
        assert wb["Added"].max_row == 3
    finally:
        wb.close()
