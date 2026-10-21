from __future__ import annotations

import builtins
import importlib
import runpy
import sys
import types
from pathlib import Path

import pandas as pd

from PyCode import main as app
from PyCode.core import app_identity


def test_script_bootstrap_and_main_entrypoint(monkeypatch, capsys):
    monkeypatch.setenv("PYTOOLS_MODE", "python")
    monkeypatch.setattr(builtins, "input", lambda _prompt="": "0")

    runpy.run_path(str(Path(app.__file__).resolve()), run_name="__main__")

    output = capsys.readouterr().out
    assert "当前模式：非 Office" in output
    assert "再见。" in output


def test_picker_focus_fallback_wrappers_and_manual_cancellation(tmp_path, monkeypatch):
    source = tmp_path / "source.xlsx"
    source.touch()

    class _Root:
        def withdraw(self):
            pass

        def attributes(self, *_args):
            pass

        def lift(self):
            pass

        def focus_force(self):
            raise RuntimeError("focus unsupported")

        def update(self):
            pass

        def destroy(self):
            pass

    tk = types.ModuleType("tkinter")
    filedialog = types.ModuleType("tkinter.filedialog")
    tk.Tk = _Root
    tk.filedialog = filedialog
    filedialog.askopenfilenames = lambda **_kwargs: (str(source),)
    monkeypatch.setitem(sys.modules, "tkinter", tk)
    monkeypatch.setitem(sys.modules, "tkinter.filedialog", filedialog)

    assert app._pick_source_files("focus fallback") == [source]

    folder = tmp_path / "source-folder"
    folder.mkdir()
    original_pick_source_folders = app._pick_source_folders
    monkeypatch.setattr(app, "_pick_source_folders", lambda title: [folder] if title == "legacy" else [])
    assert app._pick_excel_folders("legacy") == [folder]
    monkeypatch.setattr(app, "_pick_source_folders", original_pick_source_folders)

    original_import = builtins.__import__

    def no_tk(name, *args, **kwargs):
        if name == "tkinter":
            raise ImportError("no display")
        return original_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", no_tk)
    monkeypatch.setattr(builtins, "input", lambda _prompt="": "0")
    assert app._pick_source_folders("manual cancel") == []
    assert app._pick_one_workbook("manual cancel") is None


def test_view_config_tolerates_stat_and_reports_excel_open_errors(monkeypatch, capsys):
    class _UnreadableConfig:
        def exists(self):
            return True

        def stat(self):
            raise OSError("stat denied")

        def __str__(self):
            return "unreadable-config.xlsx"

    monkeypatch.setattr(app, "CFG_PATH", _UnreadableConfig())
    monkeypatch.setattr(pd, "ExcelFile", lambda _path: (_ for _ in ()).throw(ValueError("not an xlsx")))

    app._view_config()

    output = capsys.readouterr().out
    assert "配置文件: unreadable-config.xlsx" in output
    assert "[读取失败] not an xlsx" in output


def test_app_title_uses_fallback_when_identity_module_fails(monkeypatch):
    original = app_identity.app_title
    try:
        monkeypatch.setattr(
            app_identity,
            "app_title",
            lambda: (_ for _ in ()).throw(RuntimeError("metadata missing")),
        )
        importlib.reload(app)
        assert app._APP_TITLE == "pytools"
    finally:
        monkeypatch.setattr(app_identity, "app_title", original)
        importlib.reload(app)
