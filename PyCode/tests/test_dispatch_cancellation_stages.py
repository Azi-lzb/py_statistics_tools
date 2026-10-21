from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

from PyCode import main as app


def test_dispatch_cancellation_at_each_selection_stage_and_empty_results(monkeypatch, tmp_path, capsys):
    source, com_source = tmp_path / "source.xlsx", tmp_path / "com-source.xls"
    source.touch()
    com_source.touch()
    monkeypatch.setattr(app, "CFG_PATH", tmp_path / "config.xlsx")
    monkeypatch.setattr(app, "_check", lambda _choice: True)
    monkeypatch.setattr(
        app,
        "load_global",
        lambda _path: SimpleNamespace(output_dir=tmp_path / "out", log_dir=tmp_path / "logs"),
    )
    monkeypatch.setattr(app, "load_global_value", lambda *_args, **_kwargs: "")
    monkeypatch.setattr(app, "load_timeline_rules", lambda _path: [object()])
    monkeypatch.setattr(app, "load_path_maps", lambda _path: [])
    monkeypatch.setattr(app, "load_institution_mapping", lambda _path: {"raw": "standard"})
    monkeypatch.setattr(app, "load_foreign_banks", lambda _path: {"foreign"})
    monkeypatch.setattr(app, "load_extract_tasks", lambda _path: [{"area_spec": "A1"}])

    selections = iter([source, source, None, source, source, source, None, source, None, source])
    monkeypatch.setattr(app, "_pick_one_workbook", lambda *_args, **_kwargs: next(selections))
    monkeypatch.setattr(app, "_pick_source_files", lambda *_args, **_kwargs: [])

    for choice in ("s22com", "x11", "x11", "x12", "x13", "x16", "x16com", "x18", "x19", "s22"):
        app.dispatch(choice)

    monkeypatch.setattr(app, "_pick_source_files", lambda *_args, **_kwargs: [source, com_source])
    monkeypatch.setattr(app, "_split_sources_for_auto_route", lambda _paths: ([source], [com_source]))
    empty_result = {
        "matrix_empty": True,
        "files": 0,
        "branch_sheets": 0,
        "filled_sheets": 0,
        "skipped_sheets": 0,
        "failed_files": 0,
    }
    monkeypatch.setattr(app, "run_split_village_bank", lambda *_args: empty_result)
    monkeypatch.setattr(app, "run_split_village_bank_com", lambda *_args: empty_result)
    app.dispatch("x11")

    report = capsys.readouterr().out
    assert report.count("[已取消]") >= 8
    assert "Python 村镇银行计算结果为空" in report
    assert "COM 村镇银行计算结果为空" in report
