from __future__ import annotations

import os
from types import SimpleNamespace
from pathlib import Path

from nocompytools import main as app


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
    os.utime(older, ns=(1_000_000_000, 1_000_000_000))
    os.utime(newer, ns=(2_000_000_000, 2_000_000_000))
    os.utime(temporary, ns=(3_000_000_000, 3_000_000_000))
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


def test_word_conversion_dispatch_uses_global_target_format(tmp_path: Path, monkeypatch, capsys) -> None:
    source = tmp_path / "source.doc"
    source.touch()
    calls = []
    monkeypatch.setattr(app, "_check", lambda _choice: True)
    monkeypatch.setattr(
        app, "load_global",
        lambda _cfg: SimpleNamespace(output_dir=tmp_path / "output", log_dir=tmp_path / "logs"),
    )
    monkeypatch.setattr(
        app, "_pick_source_files_or_folders",
        lambda title, **kwargs: calls.append((title, kwargs)) or [source],
    )
    monkeypatch.setattr(app, "load_global_value", lambda _cfg, key, default: "docx")
    monkeypatch.setattr(
        app, "run_word_convert",
        lambda files, target, log_dir: calls.append((files, target, log_dir))
        or {"ok": 1, "skip": 0, "target_ext": "docx", "engine": "LibreOffice"},
    )

    app.dispatch("t5")

    assert calls[0][1]["suffixes"] == {".doc", ".docx"}
    assert calls[1] == ([source], "docx", tmp_path / "logs")
    assert "引擎=LibreOffice" in capsys.readouterr().out

