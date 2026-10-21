from __future__ import annotations

import sys
import types
from pathlib import Path
from types import SimpleNamespace

import pytest

from PyCode.core import com_engine


def _client(monkeypatch, dispatch):
    client = types.ModuleType("win32com.client")
    client.DispatchEx = dispatch
    package = types.ModuleType("win32com")
    package.client = client
    monkeypatch.setitem(sys.modules, "win32com", package)
    monkeypatch.setitem(sys.modules, "win32com.client", client)


def test_array_like_shape_values_and_safe_iat_access() -> None:
    empty = com_engine.ArrayLike([])
    assert empty.shape == (0, 0) and empty.values == []
    data = com_engine.ArrayLike([[1, 2], [3]])
    assert data.shape == (2, 2) and data.values == [[1, 2], [3]]
    assert data.iat[0, 1] == 2 and data.iat[1, 1] is None
    with pytest.raises(IndexError):
        _ = data.iat[2, 0]


def test_debug_flag_and_start_excel_app_fallbacks(monkeypatch) -> None:
    for value, expected in (("1", True), ("TRUE", True), ("yes", True), ("on", True), ("false", False), ("", False)):
        monkeypatch.setenv("PYTOOLS_DEBUG_COM", value)
        assert com_engine._com_debug_enabled() is expected

    calls = []
    app = SimpleNamespace(Visible=True, DisplayAlerts=True)

    def dispatch(progid):
        calls.append(progid)
        if progid != "ket.Application":
            raise OSError(progid)
        return app

    _client(monkeypatch, dispatch)
    started = com_engine.start_excel_app()
    assert started is app and calls == ["Excel.Application", "ket.Application"]
    assert app.Visible is False and app.DisplayAlerts is False
    assert app.ScreenUpdating is False and app.EnableEvents is False
    assert app.Calculation == com_engine.XL_CALC_MANUAL

    _client(monkeypatch, lambda _progid: (_ for _ in ()).throw(OSError("not registered")))
    with pytest.raises(RuntimeError, match="未能启动 Excel COM 应用"):
        com_engine.start_excel_app()

    class RestrictedApp:
        def __setattr__(self, name, value):
            if name in {"ScreenUpdating", "Calculation"}:
                raise OSError("property unsupported")
            object.__setattr__(self, name, value)

    restricted = RestrictedApp()
    _client(monkeypatch, lambda _progid: restricted)
    assert com_engine.start_excel_app() is restricted
    assert restricted.Visible is False and restricted.DisplayAlerts is False

    monkeypatch.setitem(sys.modules, "win32com", None)
    monkeypatch.setitem(sys.modules, "win32com.client", None)
    with pytest.raises(RuntimeError, match="缺少 pywin32"):
        com_engine.start_excel_app()


def test_read_sheet_values_handles_tuple_scalar_empty_and_invalid_used_range() -> None:
    captured = []

    class Sheet:
        UsedRange = SimpleNamespace(
            Row=2, Column=3, Rows=SimpleNamespace(Count=2), Columns=SimpleNamespace(Count=2)
        )

        def Cells(self, row, col):
            return (row, col)

        def Range(self, first, last):
            captured.append((first, last))
            return SimpleNamespace(Value2=(("a", "b"), ("c", "d")))

    matrix = com_engine.read_sheet_values(Sheet())
    assert matrix.values == [["a", "b"], ["c", "d"]]
    assert captured == [((1, 1), (3, 4))]

    class ScalarSheet(Sheet):
        UsedRange = SimpleNamespace(
            Row=1, Column=1, Rows=SimpleNamespace(Count=1), Columns=SimpleNamespace(Count=1)
        )

        def Range(self, *_args):
            return SimpleNamespace(Value2="value")

    assert com_engine.read_sheet_values(ScalarSheet()).values == [["value"]]

    class EmptySheet(ScalarSheet):
        def Range(self, *_args):
            return SimpleNamespace(Value2=None)

    assert com_engine.read_sheet_values(EmptySheet()).shape == (0, 0)

    class InvalidSheet(ScalarSheet):
        UsedRange = SimpleNamespace(
            Row=0, Column=0, Rows=SimpleNamespace(Count=0), Columns=SimpleNamespace(Count=0)
        )

    assert com_engine.read_sheet_values(InvalidSheet()).values == []


def test_read_merged_aware_texts_deduplicates_cells_and_handles_com_errors() -> None:
    class MergedCell:
        MergeCells = True
        MergeArea = SimpleNamespace(Cells=lambda *_args: SimpleNamespace(Value="top left"))
        Value = "covered"

    class PlainCell:
        MergeCells = False
        Value = " plain "

    class BrokenMergeCell:
        @property
        def MergeCells(self):
            raise OSError("MergeCells unavailable")

        Value = " fallback "

    class Sheet:
        def Cells(self, row, col):
            if (row, col) == (1, 1):
                return MergedCell()
            if (row, col) == (2, 2):
                return PlainCell()
            if (row, col) == (3, 3):
                return BrokenMergeCell()
            if (row, col) == (4, 4):
                return SimpleNamespace(MergeCells=False, Value=None)
            raise OSError("missing cell")

    values = com_engine.read_merged_aware_texts(
        Sheet(), [(0, 0), (0, 0), (1, 1), (2, 2), (3, 3), (4, 4), (-1, 0), (8, 8)]
    )
    assert values == {
        (0, 0): "top left", (1, 1): "plain", (2, 2): "fallback", (3, 3): "", (4, 4): "", (8, 8): ""
    }


def test_merge_range_reader_reports_ranges_and_skips_unavailable_merge_areas(monkeypatch, capsys) -> None:
    area = SimpleNamespace(Row=2, Column=3, Rows=SimpleNamespace(Count=2), Columns=SimpleNamespace(Count=4))
    sheet = SimpleNamespace(Name="Data", UsedRange=SimpleNamespace(MergeAreas=[area, area]))
    monkeypatch.setenv("PYTOOLS_DEBUG_COM", "1")
    assert com_engine.get_merge_ranges_com(sheet) == ((1, 2, 2, 5),)
    assert "MergeAreas 路径" in capsys.readouterr().out

    class MissingName:
        @property
        def Name(self):
            raise OSError("unknown name")

        UsedRange = SimpleNamespace(MergeAreas=[])

    assert com_engine.get_merge_ranges_com(MissingName()) == ()

    class BrokenUsed:
        Name = "Broken"

        @property
        def UsedRange(self):
            raise OSError("no UsedRange")

    monkeypatch.setenv("PYTOOLS_DEBUG_COM", "1")
    assert com_engine.get_merge_ranges_com(BrokenUsed()) == ()
    assert "UsedRange 读取失败" in capsys.readouterr().out

    broken_merge = SimpleNamespace(Name="NoMerges", UsedRange=SimpleNamespace(MergeAreas=object()))
    assert com_engine.get_merge_ranges_com(broken_merge) == ()
    assert "已跳过逐格回退扫描" in capsys.readouterr().out


def test_sheet_helpers_open_readonly_and_safe_close_quit(tmp_path: Path) -> None:
    first = SimpleNamespace(Name="One")
    second = SimpleNamespace(Name="Two")
    book = SimpleNamespace(Worksheets=[first, second])
    assert com_engine.list_sheet_names_com(book) == ["One", "Two"]
    assert com_engine.find_sheet_com(book, "Two") is second
    assert com_engine.find_sheet_com(book, "Missing") is None

    calls = []
    app = SimpleNamespace(Workbooks=SimpleNamespace(Open=lambda *args, **kwargs: (calls.append((args, kwargs)) or book)))
    path = tmp_path / "source.xlsx"
    result = com_engine.open_readonly(app, path)
    args, kwargs = calls[0]
    assert result is book and args == (str(path.resolve()),)
    assert kwargs == {
        "ReadOnly": True, "UpdateLinks": 0, "AddToMru": False, "Notify": False,
        "IgnoreReadOnlyRecommended": True, "CorruptLoad": 0, "Local": True,
    }

    events = []
    com_engine.safe_close(SimpleNamespace(Close=lambda **kwargs: events.append(kwargs)))
    com_engine.safe_quit(SimpleNamespace(Quit=lambda: events.append("quit")))
    com_engine.safe_close(None)
    com_engine.safe_quit(None)
    com_engine.safe_close(SimpleNamespace(Close=lambda **_kwargs: (_ for _ in ()).throw(OSError())))
    com_engine.safe_quit(SimpleNamespace(Quit=lambda: (_ for _ in ()).throw(OSError())))
    assert events == [{"SaveChanges": False}, "quit"]
