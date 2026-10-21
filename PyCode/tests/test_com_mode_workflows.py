from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

from PyCode.core import feature_1_4_fx_header_com as fx_com
from PyCode.core import print_310_com as print_com


def _header(text: str):
    return SimpleNamespace(Text=text)


def _page_setup(changed: bool):
    return SimpleNamespace(
        LeftHeader="单位：万元" if changed else "单位：元",
        CenterHeader="", RightHeader="",
        OddAndEvenPagesHeaderFooter=True,
        EvenPage=SimpleNamespace(LeftHeader=_header("万元" if changed else "元"), CenterHeader=None,
                                 RightHeader=_header("其他")),
        DifferentFirstPageHeaderFooter=True,
        FirstPage=SimpleNamespace(LeftHeader=None, CenterHeader=_header("万元" if changed else "元"),
                                  RightHeader=_header("")),
    )


class _Book:
    def __init__(self, sheets):
        self.Worksheets = sheets
        self.saved = False

    def Save(self):
        self.saved = True


def test_fx_com_header_replacement_handles_optional_sections():
    setup = _page_setup(True)

    assert fx_com._replace_unit_in_headers(setup) is True
    assert setup.LeftHeader == "单位：万美元"
    assert setup.EvenPage.LeftHeader.Text == "万美元"
    assert setup.FirstPage.CenterHeader.Text == "万美元"
    assert fx_com._replace_unit_in_headers(_page_setup(False)) is False
    assert fx_com._replace_in_subpage(None) is False

    no_optional_pages = SimpleNamespace(
        LeftHeader="", CenterHeader="", RightHeader="",
        OddAndEvenPagesHeaderFooter=False, DifferentFirstPageHeaderFooter=False,
    )
    assert fx_com._replace_unit_in_headers(no_optional_pages) is False

    class _BrokenText:
        @property
        def Text(self):
            raise RuntimeError("COM header unavailable")

    class _BrokenSetup:
        LeftHeader = ""
        RightHeader = ""
        OddAndEvenPagesHeaderFooter = True
        EvenPage = SimpleNamespace(LeftHeader=_BrokenText(), CenterHeader=None, RightHeader=None)
        DifferentFirstPageHeaderFooter = True
        FirstPage = SimpleNamespace(LeftHeader=None, CenterHeader=_BrokenText(), RightHeader=None)

        @property
        def CenterHeader(self):
            raise RuntimeError("COM property unavailable")

    assert fx_com._replace_unit_in_headers(_BrokenSetup()) is False


def test_run_fx_header_fix_com_filters_opens_modifies_and_closes(tmp_path, monkeypatch):
    paths = [tmp_path / name for name in (
        "fx.xlsx", "no-change.xlsm", "open-fail.xls", "process-fail.xlsx", "skip.csv",
    )]
    for path in paths:
        path.touch()
    changed = _Book([
        SimpleNamespace(Name="2026外汇资产", PageSetup=_page_setup(True)),
        SimpleNamespace(Name="普通表", PageSetup=_page_setup(True)),
        SimpleNamespace(Name="外汇附表", PageSetup=_page_setup(False)),
    ])
    unchanged = _Book([SimpleNamespace(Name="外汇数据", PageSetup=_page_setup(False))])

    class _BrokenBook:
        @property
        def Worksheets(self):
            raise RuntimeError("COM enumeration failed")

    opened = {"fx.xlsx": changed, "no-change.xlsm": unchanged,
              "process-fail.xlsx": _BrokenBook()}

    class _Workbooks:
        def Open(self, path, **_kwargs):
            name = Path(path).name
            if name == "open-fail.xls":
                raise OSError("locked")
            return opened[name]

    app = SimpleNamespace(Workbooks=_Workbooks())
    closed, quit_apps = [], []
    monkeypatch.setattr(fx_com, "start_excel_app", lambda: app)
    monkeypatch.setattr(fx_com, "safe_close", lambda book: closed.append(book))
    monkeypatch.setattr(fx_com, "safe_quit", lambda current: quit_apps.append(current))

    result = fx_com.run_fx_header_fix_com(paths, tmp_path / "logs")

    assert result == {"files": 2, "saved_files": 1, "fx_sheets": 3,
                      "modified_sheets": 1, "failed_files": 2}
    assert changed.saved is True and unchanged.saved is False
    assert len(closed) == 3 and quit_apps == [app]


def test_run_fx_header_fix_com_returns_when_no_valid_files(tmp_path, monkeypatch):
    path = tmp_path / "report.csv"
    path.touch()
    monkeypatch.setattr(
        fx_com, "start_excel_app", lambda: (_ for _ in ()).throw(AssertionError("not needed"))
    )

    result = fx_com.run_fx_header_fix_com([path, tmp_path / "missing.xlsx"], tmp_path / "logs")

    assert result == {"files": 0, "saved_files": 0, "fx_sheets": 0,
                      "modified_sheets": 0, "failed_files": 0}


class _PrintSheets:
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
        sheet = _PrintSheet(f"Sheet{len(self.items) + 1}", self)
        self.items.append(sheet)
        return sheet


class _PrintSheet:
    def __init__(self, name, collection=None):
        self.Name = name
        self.PageSetup = SimpleNamespace()
        self.collection = collection

    def Delete(self):
        self.collection.items.remove(self)


class _PrintBook:
    def __init__(self):
        self.Worksheets = _PrintSheets([])
        self.Worksheets.items.append(_PrintSheet("Default", self.Worksheets))


@pytest.mark.parametrize("mode", [1, 3])
def test_print_com_comment_workflow_routes_ranges_and_closes_books(tmp_path, monkeypatch, mode):
    from PyCode.core.print_310_com import run_print_fast_by_comment_com, run_print_keep_by_comment_com

    source = tmp_path / "source.xls"
    failed = tmp_path / "failed.xlsx"
    unmarked = tmp_path / "unmarked.xlsx"
    ignored = tmp_path / "ignored.csv"
    missing = tmp_path / "missing.xlsx"
    for path in (source, failed, unmarked, ignored):
        path.touch()

    class _SourceSheet:
        def __init__(self, name):
            self.Name = name
            self.UsedRange = SimpleNamespace(
                Row=1, Column=1, Rows=SimpleNamespace(Count=2),
                Columns=SimpleNamespace(Count=3),
            )

        def Cells(self, *_args):
            return SimpleNamespace(RowHeight=10, ColumnWidth=20)

    marked_book = SimpleNamespace(Worksheets=[
        _SourceSheet("Print"), _SourceSheet("Fallback"), _SourceSheet("Unmarked"),
    ])
    unmarked_book = SimpleNamespace(Worksheets=[_SourceSheet("NoPrint")])
    books = []

    class _Workbooks:
        def Add(self):
            book = _PrintBook()
            books.append(book)
            return book

    app = SimpleNamespace(SheetsInNewWorkbook=3, Workbooks=_Workbooks())
    closed, quit_apps, copies, zero_clears, saved_books = [], [], [], [], []
    monkeypatch.setattr(print_com, "start_excel_app", lambda: app)

    def open_source(_app, path):
        if path == failed:
            raise OSError("locked")
        return unmarked_book if path == unmarked else marked_book

    monkeypatch.setattr(print_com, "open_readonly", open_source)
    monkeypatch.setattr(print_com, "safe_close", lambda book: closed.append(book))
    monkeypatch.setattr(print_com, "safe_quit", lambda current: quit_apps.append(current))
    monkeypatch.setattr(
        print_com, "_a1_comment_text",
        lambda ws: {"Print": "打印 横向", "Fallback": "输出 纵向"}.get(ws.Name, ""),
    )
    monkeypatch.setattr(
        print_com, "_find_comment_ranges_com",
        lambda ws: [(1, 1, 2, 3), (3, 1, 4, 2)] if ws.Name == "Print" else [],
    )
    def copy_block(_src, _dst, bounds, _row, mode):
        copies.append((bounds, mode))
        return bounds[2] - bounds[0] + 1, bounds[3] - bounds[1] + 1

    monkeypatch.setattr(print_com, "_copy_block_com", copy_block)
    monkeypatch.setattr(print_com, "_clear_zero_in_block", lambda *args: zero_clears.append(args))
    monkeypatch.setattr(print_com, "_parse_orientation", lambda _text: None)
    monkeypatch.setattr(print_com, "_apply_print_setup_com", lambda ws, *args: setattr(ws.PageSetup, "args", args))

    def save(_app, book, path):
        saved_books.append(book)
        return path

    monkeypatch.setattr(print_com, "_save_xlsx_with_fallback", save)
    run = run_print_keep_by_comment_com if mode == 1 else run_print_fast_by_comment_com
    result = run([source, failed, unmarked, ignored, missing], tmp_path / "out", tmp_path / "logs",
                 hide_zero=True)

    assert result == {"workbooks_hit": 1, "sheets_hit": 3, "saved_files": 1, "skipped": 1}
    assert copies == [((1, 1, 2, 3), mode), ((3, 1, 4, 2), mode), ((1, 1, 2, 3), mode)]
    assert len(zero_clears) == 3 and app.SheetsInNewWorkbook == 3
    output_sheets = saved_books[0].Worksheets.items
    assert [sheet.Name for sheet in output_sheets] == ["Print_1", "Print_2", "Fallback"]
    assert [sheet.PageSetup.CenterFooter for sheet in output_sheets] == [
        "第 1 页 / 共 3 页", "第 2 页 / 共 3 页", "第 3 页 / 共 3 页",
    ]
    assert len(books) == 2 and len(closed) == 6 and quit_apps == [app]


@pytest.mark.parametrize("mode", [1, 2])
def test_print_com_copy_block_modes_keep_values_and_compact_layout(mode):
    class _SourceRange:
        Value = ((11, 12), (21, 22))

        def __init__(self):
            self.copied = False

        def Copy(self):
            self.copied = True

    source_range = _SourceRange()

    class _Source:
        def Range(self, *_args):
            return source_range

        def Cells(self, row, col):
            return SimpleNamespace(ColumnWidth=12, Value=row * 10 + col)

    class _DestinationRange:
        def __init__(self, book):
            self.book = book
            self.Value = None
            self.EntireColumn = SimpleNamespace(AutoFit=book.autofit)

    class _DestinationCell:
        def __init__(self, formula=False):
            self.HasFormula = formula
            self.Value = "formula" if formula else None
            self.ColumnWidth = 24
            self.pastes = []

        def PasteSpecial(self, **kwargs):
            self.pastes.append(kwargs["Paste"])

    class _Destination:
        def __init__(self):
            self.cells = {}
            self.rows = {}
            self.ranges = []

        def Cells(self, row, col):
            key = (row, col)
            if key not in self.cells:
                self.cells[key] = _DestinationCell(formula=key == (4, 1))
            return self.cells[key]

        def Range(self, *_args):
            result = _DestinationRange(self)
            self.ranges.append(result)
            return result

        def Rows(self, row):
            return self.rows.setdefault(row, SimpleNamespace(RowHeight=12))

        def autofit(self):
            for cell in self.cells.values():
                cell.ColumnWidth = 25

    source = _Source()
    destination = _Destination()

    result = print_com._copy_block_com(source, destination, (2, 2, 3, 3), 4, mode)

    assert result == (2, 2) and source_range.copied
    assert destination.ranges[0].Value == source_range.Value
    assert destination.Cells(4, 1).Value == 22
    assert destination.Cells(1, 1).ColumnWidth == (18 if mode == 2 else 12)
    if mode == 2:
        assert [destination.Rows(row).RowHeight for row in (4, 5)] == [15.0, 15.0]
    else:
        assert destination.Cells(4, 1).pastes == [print_com.XL_PASTE_ALL]


def test_print_com_clear_comments_falls_back_to_deleting_each_comment():
    deleted = []

    class _Comment:
        def Delete(self):
            deleted.append(True)

    class _RangeWithoutClear:
        def ClearComments(self):
            raise RuntimeError("bulk clear unavailable")

    class _Sheet:
        def Range(self, *_args):
            return _RangeWithoutClear()

        def Cells(self, row, col):
            comment = _Comment() if (row, col) != (1, 1) else None
            return SimpleNamespace(Comment=comment)

    print_com._clear_comments_in_block(_Sheet(), 1, 2, 2)

    assert len(deleted) == 3


def test_print_com_sheet_helpers_and_optional_com_properties():
    sheet = _PrintSheet("Target")
    book = SimpleNamespace(Worksheets=_PrintSheets([sheet]))

    assert print_com._ensure_sheet(book, "Target") is sheet
    added = print_com._ensure_sheet(book, "New" * 10)
    assert added.Name == "New" * 10

    cleared = []
    print_com._clear_sheet_contents(SimpleNamespace(Cells=SimpleNamespace(Clear=lambda: cleared.append(True))))
    print_com._clear_sheet_contents(SimpleNamespace(Cells=SimpleNamespace(Clear=lambda: (_ for _ in ()).throw(RuntimeError()))))
    assert cleared == [True]

    class _RestrictedApp:
        Workbooks = SimpleNamespace(Add=lambda: book)

        @property
        def SheetsInNewWorkbook(self):
            raise RuntimeError("property not exposed")

        @SheetsInNewWorkbook.setter
        def SheetsInNewWorkbook(self, _value):
            raise RuntimeError("property is read-only")

    assert print_com._add_clean_workbook(_RestrictedApp()) is book

    class _BrokenPageSetup:
        def __setattr__(self, name, value):
            if name in {"Zoom", "FitToPagesWide", "PaperSize", "PrintArea", "CenterHorizontally"}:
                raise RuntimeError(name)
            object.__setattr__(self, name, value)

    setup = _BrokenPageSetup()
    print_com._apply_print_setup_com(SimpleNamespace(PageSetup=setup), 0, 0, 0, 0, None)
    assert setup.Orientation == print_com.XL_ORIENT_PORTRAIT


def test_print_com_sheet_footer_skips_broken_com_objects():
    good = _PrintSheet("Good")

    class _BrokenWorksheets:
        Count = 2

        def __call__(self, index):
            if index == 1:
                raise RuntimeError("sheet access failed")
            return good

    print_com._apply_sheet_order_footer_com(SimpleNamespace(Worksheets=_BrokenWorksheets()))
    assert good.PageSetup.CenterFooter == "第 2 页 / 共 2 页"

    class _BrokenCount:
        @property
        def Count(self):
            raise RuntimeError("count unavailable")

    print_com._apply_sheet_order_footer_com(SimpleNamespace(Worksheets=_BrokenCount()))
