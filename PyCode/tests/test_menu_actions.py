from __future__ import annotations

from types import SimpleNamespace
from pathlib import Path

from PyCode import main as app


def test_config_menu_initializes_and_displays_config(tmp_path: Path, monkeypatch, capsys) -> None:
    config = tmp_path / "config.xlsx"
    monkeypatch.setattr(app, "CFG_PATH", config)

    app.dispatch("c")
    assert config.is_file()
    assert "配置修复" in capsys.readouterr().out

    app.dispatch("cv")
    output = capsys.readouterr().out
    assert "Sheet 列表" in output
    assert "全局配置解析" in output


def test_quick_open_config_and_latest_output(tmp_path: Path, monkeypatch, capsys) -> None:
    config = tmp_path / "config.xlsx"
    config.touch()
    output_dir = tmp_path / "output"
    output_dir.mkdir()
    older = output_dir / "older.xlsx"
    newer = output_dir / "newer.xlsx"
    temporary = output_dir / "~$newer.xlsx"
    older.write_text("older", encoding="utf-8")
    newer.write_text("newer", encoding="utf-8")
    temporary.write_text("temporary", encoding="utf-8")
    older.touch()
    newer.touch()
    temporary.touch()
    monkeypatch.setattr(app, "CFG_PATH", config)
    monkeypatch.setattr(app, "load_global", lambda _cfg: SimpleNamespace(output_dir=output_dir))
    opened: list[Path] = []
    monkeypatch.setattr(app, "_open_path", opened.append)

    app.dispatch("q1")
    app.dispatch("q2")

    assert opened == [config, newer]
    assert "已打开最新输出文件" in capsys.readouterr().out


def test_quick_latest_output_reports_empty_folder(tmp_path: Path, monkeypatch, capsys) -> None:
    output_dir = tmp_path / "empty-output"
    monkeypatch.setattr(app, "load_global", lambda _cfg: SimpleNamespace(output_dir=output_dir))

    app.dispatch("q2")

    assert "未找到输出文件" in capsys.readouterr().out

