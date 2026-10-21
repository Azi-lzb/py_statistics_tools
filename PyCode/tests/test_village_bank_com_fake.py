from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

from PyCode.core import feature_1_1_split_village_bank_com as module


class _Cell:
    def __init__(self, sheet: "_Sheet", row: int, col: int):
        self.sheet, self.row, self.col = sheet, row, col

    @property
    def Value(self):
        return self.sheet.values.get((self.row, self.col))

    @Value.setter
    def Value(self, value):
        self.sheet.values[(self.row, self.col)] = value


class _Sheet:
    def __init__(self, name: str, values: dict[tuple[int, int], object], village_row: int | None = None):
        self.Name = name
        self.values = dict(values)
        self.UsedRange = SimpleNamespace(
            Row=1,
            Rows=SimpleNamespace(Count=10),
        )
        if village_row is not None:
            self.values[(village_row, 1)] = module.VILLAGE_BANK_KEY
        self.row_ops: list[tuple[str, int]] = []

    def Cells(self, row: int, col: int):
        return _Cell(self, row, col)

    def Rows(self, row: int):
        return SimpleNamespace(
            Delete=lambda: self.row_ops.append(("delete", row)),
            Insert=lambda: self.row_ops.append(("insert", row)),
        )


class _Book:
    def __init__(self, sheets):
        self.Worksheets = list(sheets)
        self.saved = False

    def Save(self):
        self.saved = True


def _source_sheet(name: str, deposit, loan) -> _Sheet:
    return _Sheet(name, {
        (5, 3): deposit,
        (5, 4): "deposit-note",
        (5, 5): "deposit-kind",
        (5, 87): loan,
        (5, 88): "loan-note",
        (5, 89): "loan-kind",
        (6, 3): 0,
        (6, 87): 0,
    })


def test_split_village_bank_com_fake_workflow_and_empty_matrix_cleanup(tmp_path, monkeypatch):
    current = tmp_path / "current.xlsx"
    previous = tmp_path / "previous.xlsx"
    target = tmp_path / "target.xlsx"
    failed = tmp_path / "failed.xlsx"
    for path in (current, previous, target, failed):
        path.touch()

    current_book = _Book([_source_sheet("村镇银行本外币", "100", "bad-number")])
    previous_book = _Book([_source_sheet("村镇银行本外币", 80, 40)])
    good_sheet = _Sheet("本外币分机构", {}, village_row=3)
    missing_row = _Sheet("人民币分机构", {})
    unknown_type = _Sheet("其他分机构", {}, village_row=2)
    target_book = _Book([good_sheet, missing_row, unknown_type, _Sheet("说明", {})])
    closed, quit_apps, opened = [], [], []

    class _Workbooks:
        def Open(self, path, **_kwargs):
            opened.append(Path(path).name)
            name = Path(path).name
            if name == "current.xlsx":
                return current_book
            if name == "previous.xlsx":
                return previous_book
            if name == "failed.xlsx":
                raise OSError("locked")
            return target_book

    app = SimpleNamespace(Workbooks=_Workbooks())
    monkeypatch.setattr(module, "start_excel_app", lambda: app)
    monkeypatch.setattr(module, "safe_close", lambda book: closed.append(book))
    monkeypatch.setattr(module, "safe_quit", lambda current_app: quit_apps.append(current_app))

    result = module.run_split_village_bank_com(current, previous, [target, failed], tmp_path / "logs")

    assert result == {
        "files": 1,
        "branch_sheets": 3,
        "filled_sheets": 1,
        "skipped_sheets": 2,
        "failed_files": 1,
        "matrix_empty": False,
    }
    assert current_book.Worksheets[0].values[(5, 3)] == "100"
    assert good_sheet.row_ops == [("delete", 3), ("insert", 3), ("insert", 3)]
    assert good_sheet.values[(3, 1)] == module.SPLIT_ROW_1_NAME
    assert good_sheet.values[(4, 1)] == module.SPLIT_ROW_2_NAME
    assert good_sheet.values[(3, 2)] == 100.0
    assert good_sheet.values[(3, 5)] == 25.0
    assert good_sheet.values[(3, 6)] == 0.0
    assert good_sheet.values[(3, 9)] == -100.0
    assert target_book.saved
    assert opened.count("failed.xlsx") == 1
    assert closed.count(current_book) == 1 and closed.count(previous_book) == 1
    assert target_book in closed and quit_apps == [app]


def test_split_village_bank_com_empty_matrix_quits_excel(tmp_path, monkeypatch):
    current = tmp_path / "current.xlsx"
    previous = tmp_path / "previous.xlsx"
    current.touch()
    previous.touch()
    empty_book = _Book([])
    app = SimpleNamespace(Workbooks=SimpleNamespace(Open=lambda *_args, **_kwargs: empty_book))
    closed, quit_apps = [], []
    monkeypatch.setattr(module, "start_excel_app", lambda: app)
    monkeypatch.setattr(module, "safe_close", lambda book: closed.append(book))
    monkeypatch.setattr(module, "safe_quit", lambda current_app: quit_apps.append(current_app))

    result = module.run_split_village_bank_com(current, previous, [], tmp_path / "logs")

    assert result["matrix_empty"] is True
    assert closed == [empty_book, empty_book]
    assert quit_apps == [app]
