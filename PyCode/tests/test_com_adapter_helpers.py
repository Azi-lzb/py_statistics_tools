from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

from PyCode.core import feature_1_6_extract_sheets_com as extract_com
from PyCode.core import print_310_com as print_com
from PyCode.core import summary_321_322_com as summary_com
from PyCode.core.com_engine import ArrayLike


class _Collection:
    def __init__(self, items):
        self.items = list(items)

    @property
    def Count(self):
        return len(self.items)

    def __iter__(self):
        return iter(self.items)

    def __call__(self, index):
        return self.items[index - 1]

    def Add(self, **_kwargs):
        sheet = _Sheet(f"Sheet{len(self.items) + 1}")
        self.items.append(sheet)
        return sheet


class _Range:
    def __init__(self, value=None, row=1, col=1, rows=1, cols=1):
        self.Value = value
        self.Value2 = value
        self.Row = row
        self.Column = col
        self.Rows = SimpleNamespace(Count=rows)
        self.Columns = SimpleNamespace(Count=cols)
        self.Copied = False
        self.pastes = []
        self.cleared_comments = False

    def Copy(self, *_args):
        self.Copied = True

    def PasteSpecial(self, **kwargs):
        self.pastes.append(kwargs["Paste"])

    def ClearComments(self):
        self.cleared_comments = True


class _Sheet:
    def __init__(self, name, value=((),)):
        self.Name = name
        self.UsedRange = _Range(value)
        self.Cells = lambda *_args: SimpleNamespace(ColumnWidth=9, RowHeight=15, HasFormula=False, Value=None)
        self.Columns = lambda _index: SimpleNamespace(ColumnWidth=9)
        self.Rows = lambda _index: SimpleNamespace(RowHeight=15)
        self.Cells.Clear = lambda: None


class _WorksheetBook:
    def __init__(self, sheets):
        self.Worksheets = _Collection(sheets)


def test_extract_com_helper_matching_unique_names_and_workbook_creation() -> None:
    sheets = [_Sheet("Data A"), _Sheet("Data A"), _Sheet("Skip Data"), _Sheet("Other")]
    wb = SimpleNamespace(Worksheets=sheets)
    assert extract_com._list_sheet_names(wb) == ["Data A", "Data A", "Skip Data", "Other"]
    matched = extract_com._find_matched_sheets_com(
        wb, {"sheet_keywords": "Data", "exclude_keywords": "Skip"}
    )
    assert [ws.Name for ws in matched] == ["Data A"]
    out = _WorksheetBook([_Sheet("bad_name")])
    assert extract_com._unique_sheet_name_com(out, "bad/name") == "bad_name_2"
    assert extract_com._unique_sheet_name_com(out, "clean") == "clean"

    book = _WorksheetBook([_Sheet("Sheet1")])
    app = SimpleNamespace(SheetsInNewWorkbook=3, Workbooks=SimpleNamespace(Add=lambda: book))
    assert extract_com._add_clean_workbook(app) is book
    assert app.SheetsInNewWorkbook == 3


def test_extract_com_read_write_content_and_copy_helpers() -> None:
    class Source:
        UsedRange = _Range(None, row=1, col=1, rows=2, cols=2)

        def __init__(self, value):
            self.value = value

        def Cells(self, row, col):
            return (row, col)

        def Range(self, *_args):
            return _Range(self.value)

    assert extract_com._read_full_2d(Source(((1, 2), (3, 4)))) == [[1, 2], [3, 4]]
    assert extract_com._read_full_2d(Source(5)) == [[5]]
    assert extract_com._read_full_2d(Source(None)) == []

    class Destination:
        def Cells(self, row, col):
            return (row, col)

        def Range(self, *_args):
            self.range = _Range()
            return self.range

    dst = Destination()
    extract_com._write_values_com(dst, [])
    extract_com._write_values_com(dst, [[1, 2], [3]])
    assert dst.range.Value2 == ((1, 2), (3, None))
    assert extract_com._sheet_has_content_com(_Sheet("empty", ((None, ""),))) is False
    assert extract_com._sheet_has_content_com(_Sheet("data", ((None,), ("x",)))) is True
    assert extract_com._sheet_has_content_com(_Sheet("scalar", "x")) is True
    class BrokenRange:
        @property
        def Value2(self):
            raise RuntimeError("COM read failed")

    assert extract_com._sheet_has_content_com(SimpleNamespace(UsedRange=BrokenRange())) is False

    class CopySource(_Sheet):
        def __init__(self, book):
            super().__init__("source", (("value",),))
            self.book = book

        def Copy(self, **_kwargs):
            self.book.Worksheets.items.append(_Sheet("Copy", (("value",),)))

    output = _WorksheetBook([_Sheet("Sheet1")])
    success, error = extract_com._copy_full_sheet_into_output_com(output, CopySource(output), "Copied")
    assert success and not error and output.Worksheets(2).Name == "Copied"


def test_summary_com_array_and_xls_compatibility_helpers(tmp_path: Path) -> None:
    assert summary_com._supported_com(Path("book.xls"))
    assert not summary_com._supported_com(Path("book.csv"))
    array = ArrayLike([["title", None], ["x", 3]])
    assert summary_com._gcv_from_array(array, [(0, 1, 0, 1)], 2, 2) == "title"
    assert summary_com._gcv_from_array(array, [], 2, 2) == 3
    assert summary_com._gcv_from_array(array, [], 0, 2) is None

    xlsx = tmp_path / "normal.xlsx"
    xls = tmp_path / "legacy.xls"
    xlsx.touch()
    xls.touch()
    cache = {}
    temp_dir = tmp_path / "temp"
    assert summary_com._ensure_openpyxl_compatible(xlsx, None, cache, temp_dir) == xlsx
    calls = []

    class LegacyBook:
        def SaveAs(self, out, FileFormat):
            calls.append((Path(out), FileFormat))
            Path(out).write_text("xlsx", encoding="utf-8")

        def Close(self, **kwargs):
            calls.append(kwargs)

    app = SimpleNamespace(Workbooks=SimpleNamespace(Open=lambda *_args, **_kwargs: LegacyBook()))
    converted = summary_com._ensure_openpyxl_compatible(xls, app, cache, temp_dir)
    assert converted.is_file() and calls[0][1] == summary_com._XLSX_FILE_FORMAT
    assert summary_com._ensure_openpyxl_compatible(xls, app, cache, temp_dir) == converted
    assert len(calls) == 2

    class ReadSheet:
        UsedRange = _Range(None, row=1, col=1, rows=2, cols=2)

        def Cells(self, row, col):
            return (row, col)

        def Range(self, *_args):
            return _Range((("a", "b"), (1, 2)))

    data, merges = summary_com._read_src_sheet_via_com(ReadSheet(), {2}, {2})
    assert data.shape == (2, 2) and data.iat[1, 1] == 2 and merges == ()
    empty, _ = summary_com._read_src_sheet_via_com(ReadSheet(), set(), {1})
    assert empty.shape == (0, 0)


def test_print_com_range_orientation_zero_and_print_setup_helpers() -> None:
    assert print_com._is_supported_com(Path("book.xls"))
    assert not print_com._is_supported_com(Path("book.csv"))
    assert print_com._col_letter(0) == "" and print_com._col_letter(28) == "AB"
    ur = _Range(None, row=2, col=3, rows=4, cols=5)
    assert print_com._used_bounds_com(SimpleNamespace(UsedRange=ur)) == (2, 3, 5, 7)
    assert print_com._a1_comment_text(SimpleNamespace(Cells=lambda *_: SimpleNamespace(Comment=None))) == ""
    assert print_com._a1_comment_text(SimpleNamespace(
        Cells=lambda *_: SimpleNamespace(Comment=SimpleNamespace(Text=lambda: "  landscape "))
    )) == "landscape"
    assert print_com._auto_orientation_by_dim(4, 5) == "landscape"
    assert print_com._auto_orientation_by_dim(5, 4) == "portrait"

    class Cell:
        RowHeight = 10
        ColumnWidth = 20

    ws = SimpleNamespace(Cells=lambda *_: Cell())
    assert print_com._auto_orientation_by_range_com(ws, (1, 1, 2, 2)) == "landscape"
    assert print_com._auto_orientation_by_range_com(SimpleNamespace(Cells=lambda *_: (_ for _ in ()).throw(RuntimeError())), (1, 1, 2, 1)) == "portrait"
    assert print_com._is_zero_like(0) and print_com._is_zero_like("0.0")
    assert not print_com._is_zero_like(None) and not print_com._is_zero_like(True)
    assert not print_com._is_zero_like("bad")

    setup = SimpleNamespace()
    print_com._apply_print_setup_com(
        SimpleNamespace(PageSetup=setup), 0, 28, 0, 0, "landscape", True, False
    )
    assert setup.Orientation == print_com.XL_ORIENT_LANDSCAPE
    assert setup.PrintArea == "A1:AB1" and setup.FitToPagesWide == 1
    assert setup.CenterHorizontally is True and setup.CenterVertically is False


def test_print_com_comment_ranges_and_zero_block_cleanup() -> None:
    def comment(text, row, col):
        return SimpleNamespace(Text=lambda: text, Parent=SimpleNamespace(Row=row, Column=col))

    comments = [comment("打印区域1", 3, 4), comment("打印区域 #1", 1, 2), comment("ignored", 5, 5)]
    collection = lambda index: comments[index - 1]
    collection.Count = len(comments)
    assert print_com._find_comment_ranges_com(SimpleNamespace(Comments=collection)) == [(1, 2, 3, 4)]
    assert print_com._find_comment_ranges_com(SimpleNamespace(Comments=SimpleNamespace(Count="bad"))) == []

    class ValuesRange:
        def __init__(self):
            self.Value = ((0, "0", 2), (False, None, -0.0))

    class Dest:
        def Cells(self, *_args):
            return None

        def Range(self, *_args):
            self.rng = ValuesRange()
            return self.rng

    dst = Dest()
    print_com._clear_zero_in_block(dst, 1, 2, 3)
    assert dst.rng.Value == ((None, None, 2), (False, None, None))
    print_com._clear_zero_in_block(dst, 1, 0, 3)

    class CommentRange:
        def ClearComments(self):
            self.cleared = True

    class CommentDest:
        def Cells(self, *_args):
            return None

        def Range(self, *_args):
            self.rng = CommentRange()
            return self.rng

    c_dst = CommentDest()
    print_com._clear_comments_in_block(c_dst, 1, 1, 1)
    assert c_dst.rng.cleared is True


def test_print_com_formula_copy_footer_and_sheet_helpers(tmp_path: Path) -> None:
    src_range = _Range(((1, 2),))
    src_range.Value = ((1, 2),)

    class Src:
        def Range(self, *_args):
            return src_range

        def Cells(self, *_args):
            return SimpleNamespace(ColumnWidth=12, Value="cached")

    class Anchor:
        def __init__(self):
            self.pastes = []
            self.ColumnWidth = 9

        def PasteSpecial(self, **kwargs):
            self.pastes.append(kwargs["Paste"])

    class Dst:
        def __init__(self):
            self.anchor = Anchor()

        def Cells(self, *_args):
            return self.anchor

        def Range(self, *_args):
            return SimpleNamespace(Value=None, EntireColumn=SimpleNamespace(AutoFit=lambda: None))

    dst = Dst()
    assert print_com._copy_block_com(Src(), dst, (1, 1, 1, 2), 1, 3) == (1, 2)
    assert dst.anchor.pastes == [print_com.XL_PASTE_VALUES, print_com.XL_PASTE_FORMATS]
    assert dst.anchor.ColumnWidth == 12

    ws1 = _Sheet("one")
    ws2 = _Sheet("two")
    ws1.PageSetup = SimpleNamespace()
    ws2.PageSetup = SimpleNamespace()
    book = SimpleNamespace(Worksheets=_Collection([ws1, ws2]))
    print_com._apply_sheet_order_footer_com(book)
    assert ws1.PageSetup.CenterFooter == "第 1 页 / 共 2 页"
    assert ws2.PageSetup.CenterFooter == "第 2 页 / 共 2 页"

    class AppBook:
        def SaveAs(self, *_args, **_kwargs):
            raise OSError("first save failure")

    # A second SaveAs succeeds; avoid creating a real COM object.
    calls = []
    def save_as(_path, **_kwargs):
        calls.append(True)
        if len(calls) > 1:
            return
        raise OSError("locked")
    app_book = SimpleNamespace(SaveAs=save_as)
    saved = print_com._save_xlsx_with_fallback(None, app_book, tmp_path / "out.xlsx")
    assert saved.name == "out_1.xlsx" and len(calls) == 2
