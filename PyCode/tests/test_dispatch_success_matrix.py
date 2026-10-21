from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

from PyCode import main as app


def _prepare_dispatch(monkeypatch, tmp_path: Path, source: Path, com_source: Path) -> None:
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
    monkeypatch.setattr(app, "load_print_tasks", lambda _path: [object()])
    monkeypatch.setattr(app, "load_extract_tasks", lambda _path: [{"area_spec": "A1"}])
    monkeypatch.setattr(app, "load_institution_mapping", lambda _path: {"raw": "standard"})
    monkeypatch.setattr(app, "load_foreign_banks", lambda _path: {"foreign"})
    monkeypatch.setattr(app, "_pick_one_workbook", lambda *_args, **_kwargs: source)
    monkeypatch.setattr(app, "_pick_source_files", lambda *_args, **_kwargs: [source, com_source])
    monkeypatch.setattr(app, "_split_sources_for_auto_route", lambda _paths: ([source], [com_source]))


def test_dispatch_success_paths_report_python_and_com_results(monkeypatch, tmp_path, capsys):
    source, com_source = tmp_path / "source.xlsx", tmp_path / "com_source.xls"
    source.touch()
    com_source.touch()
    _prepare_dispatch(monkeypatch, tmp_path, source, com_source)

    monkeypatch.setattr(app, "run_summary_by_usedrange", lambda *_args: {
        "saved": True, "rows": 3, "path": "used.xlsx",
    })
    comment_com_results = iter([
        {"saved": False, "sheets_ok": 1, "sheets_skip": 2, "errors": ["e1", "e2", "e3", "e4"]},
        {"saved": True, "rows": 4, "path": "comment.xlsx", "sheets_ok": 2, "sheets_skip": 1},
    ])
    monkeypatch.setattr(app, "run_summary_by_usedrange_com", lambda *_args: {
        "saved": False, "rows": 0, "path": "",
    })
    monkeypatch.setattr(app, "run_summary_by_comment", lambda *_args: {
        "saved": False, "rows": 0, "path": "",
    })
    monkeypatch.setattr(app, "run_summary_by_comment_com", lambda *_args: next(comment_com_results))
    monkeypatch.setattr(app, "run_split_village_bank", lambda *_args: {
        "matrix_empty": False, "files": 1, "branch_sheets": 2,
        "filled_sheets": 1, "skipped_sheets": 1, "failed_files": 0,
    })
    monkeypatch.setattr(app, "run_split_village_bank_com", lambda *_args: {
        "matrix_empty": False, "files": 1, "branch_sheets": 3,
        "filled_sheets": 2, "skipped_sheets": 1, "failed_files": 0,
    })
    monkeypatch.setattr(app, "run_normalize_institution", lambda *_args: {
        "files": 1, "branch_sheets": 2, "mapped_sheets": 1, "mapped_rows": 3, "failed_files": 0,
    })
    monkeypatch.setattr(app, "run_normalize_institution_com", lambda *_args: {
        "files": 1, "branch_sheets": 2, "mapped_sheets": 2, "mapped_rows": 5, "failed_files": 0,
    })
    monkeypatch.setattr(app, "run_remove_foreign_bank", lambda *_args: {
        "files": 1, "ok_files": 1, "branch_sheets": 2, "deleted_rows": 1, "failed_files": 0,
    })
    monkeypatch.setattr(app, "run_remove_foreign_bank_com", lambda *_args: {
        "files": 1, "ok_files": 1, "branch_sheets": 2, "deleted_rows": 2, "failed_files": 0,
    })
    monkeypatch.setattr(app, "run_fx_header_fix", lambda *_args: {
        "files": 1, "saved_files": 1, "fx_sheets": 2, "modified_sheets": 1, "failed_files": 0,
    })
    monkeypatch.setattr(app, "run_fx_header_fix_com", lambda *_args: {
        "files": 1, "saved_files": 1, "fx_sheets": 1, "modified_sheets": 1, "failed_files": 0,
    })
    monkeypatch.setattr(app, "run_region_sum_check", lambda *_args: {
        "files": 1, "sheets": 2, "errors": 1, "failed_files": 0, "path": "check.xlsx",
    })
    monkeypatch.setattr(app, "run_region_sum_check_com", lambda *_args: {
        "files": 1, "sheets": 1, "errors": 0, "failed_files": 0, "path": "check-com.xlsx",
    })
    extracted = {"files": 1, "extracted_sheets": 2, "output_files": 1,
                 "failed_files": 0, "paths": ["extract.xlsx"]}
    monkeypatch.setattr(app, "run_extract_sheets", lambda *_args, **_kwargs: extracted)
    monkeypatch.setattr(app, "run_extract_sheets_com", lambda *_args, **_kwargs: extracted)
    monkeypatch.setattr(app, "run_adjust_rural_loan", lambda *_args: {
        "error": "", "matched": 2, "fill_total": 10.0, "area_total": 10.0,
        "consistent": True, "unmatched_current": ["A"], "previous_only": ["B"],
    })
    monkeypatch.setattr(app, "run_adjust_rural_loan_quarterly", lambda *_args: {
        "error": "", "matched": 2, "fill_total": 10.0, "area_total": 9.0,
        "consistent": False,
        "regions": [
            {"region": "甲", "error": "地区文件缺失", "fill_total": 0, "area_total": 0,
             "consistent": False, "unmatched_current": [], "previous_only": []},
            {"region": "乙", "error": "", "fill_total": 10.0, "area_total": 9.0,
             "consistent": False, "unmatched_current": ["C"], "previous_only": ["D"]},
        ],
    })

    for choice in (
        "s21com", "s22com", "x11", "x12", "x13", "x14", "x15",
        "x16", "x14com", "x16com",
    ):
        app.dispatch(choice)
    app.dispatch("s22com")
    monkeypatch.setattr(app, "_split_sources_for_auto_route", lambda _paths: ([source], []))
    app.dispatch("x18")
    app.dispatch("x19")
    app.dispatch("s22")

    report = capsys.readouterr().out
    for expected in (
        "COM 无可汇总数据", "跳过异常sheet 4 个", "  - e3",
        "Python 文件=1 分机构sheet=2", "COM 文件=1 分机构sheet=2",
        "删除行=1", "删除行=2", "修改sheet=1", "check-com.xlsx", "extract.xlsx",
        "本期未匹配到上期: A", "上期独有: B", "甲：地区文件缺失", "乙：分机构=10.00",
        "本期未匹配：C", "上期独有：D", "行数=4 输出=comment.xlsx",
    ):
        assert expected in report
