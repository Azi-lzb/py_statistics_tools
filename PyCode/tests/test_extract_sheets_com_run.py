from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

from PyCode.core import feature_1_6_extract_sheets_com as extract_com


class _Sheets:
    def __init__(self, sheets):
        self.items = list(sheets)

    @property
    def Count(self):
        return len(self.items)

    def __iter__(self):
        return iter(self.items)

    def __call__(self, index):
        return self.items[index - 1]

    def Add(self, **_kwargs):
        sheet = _OutputSheet(f"Sheet{len(self.items) + 1}")
        self.items.append(sheet)
        return sheet


class _OutputSheet:
    def __init__(self, name):
        self.Name = name
        self.values = None
        self.deleted = False

    def Delete(self):
        self.deleted = True
        self.book.Worksheets.items.remove(self)


class _OutputBook:
    def __init__(self, name, saves):
        self.Worksheets = _Sheets([])
        default = _OutputSheet(name)
        default.book = self
        self.Worksheets.items.append(default)
        self.saves = saves

    def SaveAs(self, path, **kwargs):
        self.saves.append(("SaveAs", Path(path), kwargs))
        if Path(path).name == "save-fail.xlsx":
            raise OSError("disk full")
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        Path(path).touch()

    def Save(self):
        self.saves.append(("Save",))


class _SourceSheet:
    def __init__(self, name, values):
        self.Name = name
        self.values = values


class _SourceBook:
    def __init__(self, sheets):
        self.Worksheets = _Sheets(sheets)


def test_run_extract_sheets_com_routes_tasks_and_saves_outputs(tmp_path, monkeypatch):
    source = tmp_path / "annual_source.xlsx"
    append_source = tmp_path / "append_source.xlsx"
    broken = tmp_path / "broken_source.xlsx"
    for path in (source, append_source, broken):
        path.touch()

    annual_book = _SourceBook([
        _SourceSheet("Data", [["a", 1], ["b", 2]]),
        _SourceSheet("Skip Data", [["ignored"]]),
        _SourceSheet("Copy", [["copy", 3]]),
        _SourceSheet("Fallback", [["fallback", None], [None, 4]]),
        _SourceSheet("Empty", [[None, ""]]),
        _SourceSheet("SaveFail", [["not saved"]]),
    ])
    append_book = _SourceBook([_SourceSheet("Second", [["append"]])])
    existing_path = tmp_path / "out" / "existing.xlsx"
    existing_path.parent.mkdir()
    existing_path.touch()

    saves = []
    existing_output = _OutputBook("Existing", saves)
    existing_output.Worksheets(1).Name = "Existing"
    new_outputs = []

    class _Workbooks:
        def Add(self):
            book = _OutputBook("Default", saves)
            new_outputs.append(book)
            return book

        def Open(self, *_args, **_kwargs):
            return existing_output

    app = SimpleNamespace(SheetsInNewWorkbook=3, Workbooks=_Workbooks())
    closed = []
    quit_apps = []
    monkeypatch.setattr(extract_com, "start_excel_app", lambda: app)

    def open_source(_app, path):
        if path.name == "broken_source.xlsx":
            raise OSError("cannot open")
        return append_book if path.name == "append_source.xlsx" else annual_book

    monkeypatch.setattr(extract_com, "open_readonly", open_source)
    monkeypatch.setattr(extract_com, "safe_close", lambda book: closed.append(book))
    monkeypatch.setattr(extract_com, "safe_quit", lambda current: quit_apps.append(current))
    monkeypatch.setattr(extract_com, "_read_full_2d", lambda sheet: sheet.values)

    def write_values(sheet, values):
        sheet.values = values

    monkeypatch.setattr(extract_com, "_write_values_com", write_values)

    def copy_sheet(book, sheet, title):
        if sheet.Name == "Copy":
            new_sheet = _OutputSheet(title)
            new_sheet.book = book
            book.Worksheets.items.append(new_sheet)
            return True, ""
        return False, "copy unavailable"

    monkeypatch.setattr(extract_com, "_copy_full_sheet_into_output_com", copy_sheet)
    tasks = [
        {"workbook_keywords": "annual", "sheet_keywords": "Data", "exclude_keywords": "Skip",
         "area_spec": "A1:B2", "output_name": "partial"},
        {"workbook_keywords": "annual", "sheet_keywords": "Data", "area_spec": "bad!",
         "output_name": "invalid"},
        {"workbook_keywords": "annual", "sheet_keywords": "Copy", "area_spec": "整表",
         "output_name": "copied"},
        {"workbook_keywords": "annual", "sheet_keywords": "Fallback", "area_spec": "整表",
         "output_name": "fallback"},
        {"workbook_keywords": "annual", "sheet_keywords": "Empty", "area_spec": "整表",
         "output_name": "empty"},
        {"workbook_keywords": "annual", "sheet_keywords": "SaveFail", "area_spec": "整表",
         "output_name": "save-fail"},
        {"workbook_keywords": "append_source", "sheet_keywords": "Second", "area_spec": "A1",
         "output_name": "existing"},
    ]

    result = extract_com.run_extract_sheets_com(
        [source, append_source, broken], tasks, tmp_path / "out", tmp_path / "logs",
        append_existing=True,
    )

    assert result["files"] == 2 and result["failed_files"] == 1
    assert result["extracted_sheets"] == 5
    assert result["output_files"] == 6
    assert {Path(path).name for path in result["paths"]} == {
        "partial.xlsx", "invalid.xlsx", "copied.xlsx", "fallback.xlsx", "empty.xlsx", "existing.xlsx",
    }
    assert new_outputs[0].Worksheets(1).values == [["a", 1], ["b", 2]]
    assert any(entry[0] == "Save" for entry in saves)
    assert len(closed) == 9 and quit_apps == [app]


def test_run_extract_sheets_com_empty_tasks_does_not_start_excel(tmp_path, monkeypatch):
    monkeypatch.setattr(
        extract_com, "start_excel_app", lambda: (_ for _ in ()).throw(AssertionError("not needed"))
    )

    result = extract_com.run_extract_sheets_com([], [], tmp_path / "out", tmp_path / "logs")

    assert result == {"files": 0, "extracted_sheets": 0, "output_files": 0,
                      "failed_files": 0, "paths": []}


@pytest.mark.parametrize("outcome", ["copied", "empty", "copy-error"])
def test_copy_full_sheet_com_falls_back_and_cleans_failed_sheets(outcome):
    class _UsedRange:
        def __init__(self, values):
            self.Value2 = values
            self.Value = values
            self.Row = 2
            self.Column = 3
            self.Rows = SimpleNamespace(Count=2)
            self.Columns = SimpleNamespace(Count=2)

    class _Destination:
        def __init__(self, sheet):
            self.sheet = sheet

    class _Dimension:
        def __init__(self, values, index, default):
            self.values = values
            self.index = index
            self.default = default

        @property
        def ColumnWidth(self):
            return self.values.get(self.index, self.default)

        @ColumnWidth.setter
        def ColumnWidth(self, value):
            self.values[self.index] = value

        @property
        def RowHeight(self):
            return self.values.get(self.index, self.default)

        @RowHeight.setter
        def RowHeight(self, value):
            self.values[self.index] = value

    class _Collection:
        def __init__(self):
            self.items = []

        @property
        def Count(self):
            return len(self.items)

        def __iter__(self):
            return iter(self.items)

        def __call__(self, index):
            return self.items[index - 1]

        def Add(self, **_kwargs):
            sheet = _Sheet("Sheet2", self)
            self.items.append(sheet)
            return sheet

    class _Sheet:
        def __init__(self, name, collection=None, values=(("value",),)):
            self.Name = name
            self.collection = collection
            self.UsedRange = _UsedRange(values)
            self.widths = {}
            self.heights = {}

        def Copy(self, **_kwargs):
            raise RuntimeError("Excel sheet copy unavailable")

        def Range(self, _address):
            return _Destination(self)

        def Columns(self, index):
            return _Dimension(self.widths, index, index + 10)

        def Rows(self, index):
            return _Dimension(self.heights, index, index + 10)

        def Delete(self):
            self.collection.items.remove(self)

    class _SourceRange(_UsedRange):
        def Copy(self, destination):
            if outcome == "copy-error":
                raise OSError("range copy failed")
            destination.sheet.UsedRange.Value2 = (
                ((None, ""),) if outcome == "empty" else (("kept", 4), ("row", 5))
            )
            destination.sheet.UsedRange.Value = destination.sheet.UsedRange.Value2

    class _SourceSheet(_Sheet):
        def __init__(self):
            super().__init__("source")
            self.UsedRange = _SourceRange((("kept", 4), ("row", 5)))

        def Range(self, _address):
            return _Destination(self)

    source = _SourceSheet()
    output = SimpleNamespace(Worksheets=_Collection())
    output.Worksheets.items.append(_Sheet("Default", output.Worksheets))

    result = extract_com._copy_full_sheet_into_output_com(output, source, "Copied")

    if outcome == "copied":
        assert result == (True, "")
        copied = output.Worksheets(2)
        assert copied.Name == "Copied"
        assert copied.widths == {3: 13, 4: 14}
        assert copied.heights == {2: 12, 3: 13}
    elif outcome == "empty":
        assert result == (False, "UsedRange复制后为空")
        assert output.Worksheets.Count == 1
    else:
        assert result == (False, "range copy failed")
        assert output.Worksheets.Count == 1
