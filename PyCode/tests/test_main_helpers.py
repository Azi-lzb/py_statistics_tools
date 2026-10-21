from __future__ import annotations

import builtins
import sys
import types
from pathlib import Path

import pytest

from PyCode import main as app


class _FakeRoot:
    def withdraw(self):
        pass

    def attributes(self, *_args):
        pass

    def lift(self):
        pass

    def focus_force(self):
        pass

    def update(self):
        pass

    def destroy(self):
        pass


def _install_tk(
    monkeypatch,
    *,
    filenames=(),
    filelists=(),
    directories=(),
    answers=(),
):
    tk = types.ModuleType("tkinter")
    filedialog = types.ModuleType("tkinter.filedialog")
    messagebox = types.ModuleType("tkinter.messagebox")
    filenames = iter(filenames)
    filelists = iter(filelists)
    directories = iter(directories)
    answers = iter(answers)
    tk.Tk = _FakeRoot
    filedialog.askopenfilename = lambda **_kwargs: next(filenames, "")
    filedialog.askopenfilenames = lambda **_kwargs: next(filelists, ())
    filedialog.askdirectory = lambda **_kwargs: next(directories, "")
    messagebox.askyesno = lambda *_args, **_kwargs: next(answers, False)
    tk.filedialog = filedialog
    tk.messagebox = messagebox
    monkeypatch.setitem(sys.modules, "tkinter", tk)
    monkeypatch.setitem(sys.modules, "tkinter.filedialog", filedialog)
    monkeypatch.setitem(sys.modules, "tkinter.messagebox", messagebox)


def _disable_tk(monkeypatch):
    original_import = builtins.__import__

    def import_without_tk(name, *args, **kwargs):
        if name == "tkinter":
            raise ImportError("tk unavailable")
        return original_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", import_without_tk)


def test_file_dialog_helpers_accept_selected_files_and_folders(tmp_path: Path, monkeypatch) -> None:
    template, source, second_source = (tmp_path / name for name in ("template.xlsx", "a.xlsx", "b.xlsx"))
    for path in (template, source, second_source):
        path.touch()
    folder_a, folder_b = tmp_path / "a", tmp_path / "b"
    folder_a.mkdir()
    folder_b.mkdir()
    _install_tk(
        monkeypatch,
        filenames=[str(template), str(source)],
        filelists=[(str(source), str(second_source)), (str(source),)],
        directories=[str(folder_a), str(folder_a), str(folder_b), str(folder_b), str(folder_a)],
        answers=[True, True, False],
    )
    monkeypatch.setattr(builtins, "input", lambda _prompt="": "2")

    assert app._pick_files_for_compare() == (template, [source, second_source])
    assert app._pick_source_files("choose") == [source]
    assert app._pick_one_workbook("choose") == source
    assert app._pick_source_folders("folders") == [folder_a, folder_b]
    assert app._pick_archive_paths("archive") == [folder_b]
    assert app._pick_archive_target_root() == folder_a


def test_picker_fallbacks_filter_missing_paths_and_handle_cancellation(tmp_path: Path, monkeypatch) -> None:
    template, source = tmp_path / "template.xlsx", tmp_path / "source.xlsx"
    folder = tmp_path / "folder"
    template.touch()
    source.touch()
    folder.mkdir()
    _disable_tk(monkeypatch)
    answers = iter([
        f'"{template}"', f'"{source}"；{tmp_path / "missing.xlsx"}',
        f'{source};{tmp_path / "missing.csv"}', f'{folder};;{tmp_path / "missing-folder"}',
        "2", str(folder), str(folder), str(source),
    ])
    monkeypatch.setattr(builtins, "input", lambda _prompt="": next(answers))

    assert app._pick_files_for_compare() == (template, [source])
    assert app._pick_source_files() == [source]
    assert app._pick_source_folders("folders") == [folder]
    assert app._pick_archive_paths("archive") == [folder]
    assert app._pick_archive_target_root() == folder
    assert app._pick_one_workbook("workbook") == source


def test_dialog_fallback_cancel_and_missing_selections(tmp_path: Path, monkeypatch) -> None:
    _install_tk(monkeypatch, filenames=["missing.xlsx"], directories=[""])
    assert app._pick_one_workbook("missing") is None
    monkeypatch.setattr(builtins, "input", lambda _prompt="": "0")
    assert app._pick_source_folders("folders") == []

    _disable_tk(monkeypatch)
    monkeypatch.setattr(builtins, "input", lambda _prompt="": "0")
    assert app._pick_source_files("cancel") == []
    assert app._pick_archive_paths("cancel") == []

    answers = iter(["1", "0", "not-a-folder", "0"])
    monkeypatch.setattr(builtins, "input", lambda _prompt="": next(answers))
    monkeypatch.setattr(app, "_pick_source_files", lambda *_args, **_kwargs: [tmp_path / "source.txt"])
    assert app._pick_archive_paths("files") == [tmp_path / "source.txt"]
    assert app._pick_archive_paths("invalid-folder") == []


def test_compare_picker_cancellation_and_manual_entry(tmp_path: Path, monkeypatch) -> None:
    _install_tk(monkeypatch, filenames=[""])
    assert app._pick_files_for_compare() == (None, [])

    _disable_tk(monkeypatch)
    template = tmp_path / "template.xlsx"
    source = tmp_path / "source.xlsx"
    template.touch()
    source.touch()
    answers = iter([
        "0", f'"{template}"', "0", str(template), f'"{source}"；{tmp_path / "missing.xlsx"}',
    ])
    monkeypatch.setattr(builtins, "input", lambda _prompt="": next(answers))
    assert app._pick_files_for_compare() == (None, [])
    assert app._pick_files_for_compare() == (None, [])
    picked_template, picked_sources = app._pick_files_for_compare()
    assert picked_template == template
    assert picked_sources == [source]
    app._print_selected_paths([])


def test_collect_latest_and_auto_route_helpers(tmp_path: Path, monkeypatch) -> None:
    assert app._latest_file_in_dir(tmp_path / "missing") is None
    empty = tmp_path / "empty"
    empty.mkdir()
    assert app._latest_file_in_dir(empty) is None
    old, latest, ignored = (empty / name for name in ("old.xlsx", "latest.xlsx", "~$lock.xlsx"))
    for path in (old, latest, ignored):
        path.touch()
    import os
    os.utime(old, (1, 1))
    os.utime(latest, (2, 2))
    os.utime(ignored, (3, 3))
    assert app._latest_file_in_dir(empty) == latest

    paths, tasks = [old, latest], [{"row": 1}]
    original_mode, original_platform = app.ENGINE_MODE, app.sys.platform
    try:
        app.ENGINE_MODE = "office"
        monkeypatch.setattr(app.sys, "platform", "win32")
        assert app._uses_windows_com()
        assert app._split_sources_for_auto_route(paths) == ([], paths)
        assert app._split_print_tasks_for_auto_route(tasks) == ([], tasks)
        monkeypatch.setattr(app.sys, "platform", "linux")
        assert not app._uses_windows_com()
        assert app._split_sources_for_auto_route(paths) == (paths, [])
        assert app._split_print_tasks_for_auto_route(tasks) == (tasks, [])
    finally:
        app.ENGINE_MODE = original_mode
        monkeypatch.setattr(app.sys, "platform", original_platform)

    folder = tmp_path / "excel"
    folder.mkdir()
    (folder / "book.xlsx").touch()
    monkeypatch.setattr(app, "_pick_excel_folders", lambda _title: [folder])
    assert app._pick_excel_sources_from_folder("excel", recursive=False) == [folder / "book.xlsx"]
    monkeypatch.setattr(app, "_pick_source_folders", lambda _title: [folder])
    assert app._pick_excel_folders("compat") == [folder]
    monkeypatch.setattr(app, "_pick_excel_folders", lambda _title: [])
    assert app._pick_excel_sources_from_folder("empty", recursive=False) == []


def test_source_folder_picker_empty_folder_branch(monkeypatch) -> None:
    monkeypatch.setattr(builtins, "input", lambda _prompt="": "2")
    monkeypatch.setattr(app, "_pick_source_folders", lambda _title: [])
    assert app._pick_source_files_or_folders("folders") == []


def test_file_or_folder_modes_and_excel_folder_compatibility(tmp_path: Path, monkeypatch) -> None:
    folder = tmp_path / "source"
    child = folder / "nested"
    child.mkdir(parents=True)
    (folder / "root.xlsx").touch()
    (child / "nested.csv").touch()
    monkeypatch.setattr(app, "_pick_source_folders", lambda _title: [folder])
    monkeypatch.setattr(app, "_pick_source_files", lambda *_args, **_kwargs: [folder / "root.xlsx"])
    monkeypatch.setattr(app, "_pick_excel_folders", lambda _title: [folder])
    answers = iter(["1", "2", "3", "0"])
    monkeypatch.setattr(builtins, "input", lambda _prompt="": next(answers))

    assert app._pick_source_files_or_folders("files") == [folder / "root.xlsx"]
    assert app._pick_source_files_or_folders("current") == [folder / "root.xlsx"]
    assert app._pick_source_files_or_folders("recursive") == [child / "nested.csv", folder / "root.xlsx"]
    assert app._pick_source_files_or_folders("cancel") == []
    assert app._pick_excel_sources_from_folder("excel", recursive=True) == [folder / "root.xlsx"]


def test_engine_selection_and_main_menu_helpers_cover_defaults_and_invalid_input(monkeypatch, capsys) -> None:
    original_mode = app.ENGINE_MODE
    try:
        app.ENGINE_MODE = "office"
        assert app._select_engine_mode() == "office"
        app.ENGINE_MODE = ""
        answers = iter(["1", "invalid", "2", "bad", "2"])
        monkeypatch.setattr(builtins, "input", lambda _prompt="": next(answers))
        assert app._select_engine_mode() == "office"
        assert app._show_main_menu() == "2"
        assert app._show_sub_menu("test", ("1", "2")) == "2"
        output = capsys.readouterr().out
        assert "仅接受" in output and "输入无效" in output
    finally:
        app.ENGINE_MODE = original_mode


def test_menu_and_engine_helpers_handle_eof(monkeypatch) -> None:
    original_mode = app.ENGINE_MODE
    try:
        app.ENGINE_MODE = ""
        monkeypatch.setattr(builtins, "input", lambda _prompt="": (_ for _ in ()).throw(EOFError()))
        assert app._select_engine_mode() == "python"
        assert app._show_main_menu() == "0"
        assert app._show_sub_menu("menu", ("1",)) == "0"
    finally:
        app.ENGINE_MODE = original_mode


@pytest.mark.parametrize(
    ("mode", "platform", "expected"),
    [
        ("office", "win32", "Excel/WPS COM"),
        ("office", "linux", "LibreOffice"),
        ("office", "darwin", "此系统没有 COM/LibreOffice 适配"),
        ("python", "linux", "非 Office"),
    ],
)
def test_main_announces_selected_engine_mode(mode: str, platform: str, expected: str, monkeypatch, capsys) -> None:
    monkeypatch.setattr(app, "_select_engine_mode", lambda: mode)
    monkeypatch.setattr(app, "_show_main_menu", lambda: "0")
    monkeypatch.setattr(app.sys, "platform", platform)
    monkeypatch.setenv("PYTOOLS_MODE", "")

    app.main()

    assert expected in capsys.readouterr().out


def test_base_dir_dedup_task_conversion_and_config_check(tmp_path: Path, monkeypatch, capsys) -> None:
    monkeypatch.setattr(app.sys, "frozen", False, raising=False)
    assert app._app_base_dir() == Path(app.__file__).resolve().parent
    executable = tmp_path / "bundle" / "pytools.exe"
    monkeypatch.setattr(app.sys, "frozen", True, raising=False)
    monkeypatch.setattr(app.sys, "executable", str(executable))
    assert app._app_base_dir() == executable.parent.resolve()

    converted = app._to_dedup_tasks([
        {"source_wb": None},
        {"row_no": "3", "enabled": 1, "source_wb": tmp_path / "source.xlsx", "key_cols": "AB"},
    ])
    assert len(converted) == 1
    assert converted[0].row_no == 3 and converted[0].enabled is True
    assert converted[0].key_cols == ["A", "B"]

    monkeypatch.setattr(app, "validate_sheets_for_feature", lambda *_args: ["missing config"])
    assert app._check("feature") is False
    assert "missing config" in capsys.readouterr().out
    monkeypatch.setattr(app, "validate_sheets_for_feature", lambda *_args: [])
    assert app._check("feature") is True


def test_open_path_and_archive_folder_fallbacks(tmp_path: Path, monkeypatch, capsys) -> None:
    existing = tmp_path / "existing.txt"
    existing.touch()
    opened = []
    monkeypatch.setattr(app.os, "startfile", lambda path: opened.append(path), raising=False)
    app._open_path(existing)
    assert opened == [str(existing)]
    with pytest.raises(FileNotFoundError):
        app._open_path(tmp_path / "missing")

    _disable_tk(monkeypatch)
    folder = tmp_path / "folder"
    folder.mkdir()
    answers = iter(["2", str(folder), "2", "missing-folder", "0"])
    monkeypatch.setattr(builtins, "input", lambda _prompt="": next(answers))
    assert app._pick_archive_paths("folder") == [folder]
    assert app._pick_archive_paths("invalid-folder") == []
    assert app._pick_archive_target_root() is None
    assert "文件夹不存在" in capsys.readouterr().out
