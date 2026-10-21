from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

from PyCode import main as app


def test_dispatch_success_paths_for_all_feature_commands(tmp_path: Path, monkeypatch) -> None:
    config = tmp_path / "config.xlsx"
    source = tmp_path / "source.xlsx"
    output = tmp_path / "result.xlsx"
    global_config = SimpleNamespace(output_dir=tmp_path / "output", log_dir=tmp_path / "logs")
    monkeypatch.setattr(app, "CFG_PATH", config)
    monkeypatch.setattr(app, "load_global", lambda _path: global_config)
    monkeypatch.setattr(app, "load_global_value", lambda _path, _key, default="": "是")
    monkeypatch.setattr(app, "_check", lambda _choice: True)
    monkeypatch.setattr(app, "_pick_one_workbook", lambda *_args, **_kwargs: source)
    monkeypatch.setattr(app, "_pick_source_files", lambda *_args, **_kwargs: [source])
    monkeypatch.setattr(app, "_pick_source_files_or_folders", lambda *_args, **_kwargs: [source])
    monkeypatch.setattr(app, "_pick_files_for_compare", lambda: (source, [source]))
    monkeypatch.setattr(app, "_pick_archive_paths", lambda *_args, **_kwargs: [source])
    monkeypatch.setattr(app, "_pick_archive_target_root", lambda: tmp_path / "archive")
    pair_split_count = 0

    def split_sources(paths):
        nonlocal pair_split_count
        selected = list(paths)
        if len(selected) != 2:
            return selected, selected
        pair_split_count += 1
        if pair_split_count in (1, 4):
            return selected, []
        return [], selected

    monkeypatch.setattr(app, "_split_sources_for_auto_route", split_sources)
    monkeypatch.setattr(app, "_split_print_tasks_for_auto_route", lambda tasks: (list(tasks), list(tasks)))
    monkeypatch.setattr(app, "_to_dedup_tasks", lambda _tasks: [object()])
    monkeypatch.setattr(app, "load_timeline_rules", lambda _path: [SimpleNamespace(name="rule")])
    monkeypatch.setattr(app, "load_path_maps", lambda _path: [object()])
    monkeypatch.setattr(app, "load_dedup_tasks", lambda _path: [object()])
    monkeypatch.setattr(app, "load_print_tasks", lambda _path: [object()])
    monkeypatch.setattr(app, "load_institution_mapping", lambda _path: {"old": "new"})
    monkeypatch.setattr(app, "load_foreign_banks", lambda _path: {"foreign"})
    monkeypatch.setattr(app, "load_extract_tasks", lambda _path: [object()])

    stats = {
        "run_compare": output,
        "run_compare_com": output,
        "run_timeline_slim": output,
        "run_timeline_slim_com": output,
        "run_wide_summary": output,
        "run_wide_summary_com": output,
        "run_check_by_comment": {"hit_sheets": 1, "marked_rows": 1},
        "run_delete_by_comment": {"hit_sheets": 1, "deleted_rows": 1},
        "run_check_by_config": {"task_ok": 1, "task_skip": 0, "marked_rows": 1, "elapsed_sec": 0.1},
        "run_delete_by_config": {"task_ok": 1, "task_skip": 0, "deleted_rows": 1, "elapsed_sec": 0.1},
        "run_append_by_config": {"task_ok": 1, "task_skip": 0, "appended_rows": 1, "deleted_rows": 0, "elapsed_sec": 0.1},
        "run_precheck_by_config": {"ok": 1, "warn": 0, "fail": 0},
        "run_print_keep_by_comment": {"workbooks_hit": 1, "sheets_hit": 1, "saved_files": 1, "skipped": 0},
        "run_print_keep_by_comment_com": {"workbooks_hit": 1, "sheets_hit": 1, "saved_files": 1, "skipped": 0},
        "run_print_fast_by_comment": {"workbooks_hit": 1, "sheets_hit": 1, "saved_files": 1, "skipped": 0},
        "run_print_fast_by_comment_com": {"workbooks_hit": 1, "sheets_hit": 1, "saved_files": 1, "skipped": 0},
        "run_print_config_all_modes": {"task_ok": 1, "task_skip": 0, "written_sheets": 1, "written_rows": 2, "task_total": 1, "elapsed_sec": 0.1, "slow_top": []},
        "run_print_config_all_modes_com": {"task_ok": 1, "task_skip": 0, "written_sheets": 1, "written_rows": 2},
        "run_print_config_precheck": {"ok": 1, "warn": 0, "fail": 0},
        "run_print_config_to_pdf_com": {"task_ok": 1, "task_skip": 0, "written_sheets": 1, "written_rows": 2, "pdf_files": [tmp_path / "result.pdf"]},
        "run_batch_rename_33": {"ok": 1, "skip": 0},
        "run_excel_convert_34": {"target_ext": "xlsx", "engine": "Python", "ok": 1, "skip": 0},
        "run_word_convert_36": {"target_ext": "docx", "engine": "LibreOffice", "ok": 1, "skip": 0},
        "run_batch_rename_sheet_37": {"workbooks": 1, "rename_ok": 1, "skip": 0},
        "run_batch_rename_sheet_37_com": {"workbooks": 1, "rename_ok": 1, "skip": 0, "engine": "Excel.Application"},
        "run_batch_rename_by_content_46": {"config_row": 2, "spec": "sheet!A1", "config_remark": "ok", "ok": 1, "skip": 0},
        "run_batch_unrename_by_content_47": {"config_row": 2, "spec": "sheet!A1", "ok": 1, "skip": 0},
        "run_batch_rename_local_map_48": {"rules": 1, "ok": 1, "skip": 0},
        "run_submission_check_48": {"output": output, "rules": 1, "files": 1, "submitted": 1, "missing": 0, "unmatched": 0, "multi": 0},
        "run_file_sort_prefix_49": {"output": output, "rules": 1, "files": 1, "renamed": 1, "unmatched": 0, "skipped": 0, "failed": 0},
        "run_remove_file_sort_prefix_410": {"output": output, "rules": 1, "files": 1, "renamed": 1, "skipped": 0, "failed": 0},
        "run_archive_by_date_preview": {"output": output, "files": 1, "skipped_ext": 0, "invalid": 0, "target_root": tmp_path},
        "run_archive_by_type_preview": {"output": output, "files": 1, "rules": 1, "skipped_ext": 0, "invalid": 0, "target_root": tmp_path},
        "run_archive_plan_copy": {"output": output, "copied": 1, "skip": 0, "failed": 0},
        "run_archive_by_date_copy": {"output": output, "copied": 1, "skip": 0, "skipped_ext": 0, "failed": 0, "target_root": tmp_path},
        "run_archive_by_type_copy": {"output": output, "copied": 1, "rules": 1, "skip": 0, "skipped_ext": 0, "failed": 0, "target_root": tmp_path},
        "run_summary_by_usedrange": {"saved": True, "rows": 1, "path": output},
        "run_summary_by_usedrange_com": {"saved": True, "rows": 1, "path": output},
        "run_summary_by_comment": {"saved": True, "rows": 1, "path": output},
        "run_summary_by_comment_com": {"saved": True, "rows": 1, "path": output, "sheets_ok": 1, "sheets_skip": 0, "errors": []},
        "run_split_village_bank": {"files": 1, "branch_sheets": 1, "filled_sheets": 1, "skipped_sheets": 0, "failed_files": 0, "matrix_empty": False},
        "run_split_village_bank_com": {"files": 1, "branch_sheets": 1, "filled_sheets": 1, "skipped_sheets": 0, "failed_files": 0, "matrix_empty": False},
        "run_normalize_institution": {"files": 1, "branch_sheets": 1, "mapped_sheets": 1, "mapped_rows": 1, "failed_files": 0},
        "run_normalize_institution_com": {"files": 1, "branch_sheets": 1, "mapped_sheets": 1, "mapped_rows": 1, "failed_files": 0},
        "run_remove_foreign_bank": {"files": 1, "ok_files": 1, "branch_sheets": 1, "deleted_rows": 1, "failed_files": 0},
        "run_remove_foreign_bank_com": {"files": 1, "ok_files": 1, "branch_sheets": 1, "deleted_rows": 1, "failed_files": 0},
        "run_fx_header_fix": {"files": 1, "saved_files": 1, "fx_sheets": 1, "modified_sheets": 1, "failed_files": 0},
        "run_fx_header_fix_com": {"files": 1, "saved_files": 1, "fx_sheets": 1, "modified_sheets": 1, "failed_files": 0},
        "run_region_sum_check": {"files": 1, "sheets": 1, "errors": 0, "failed_files": 0, "path": output},
        "run_region_sum_check_com": {"files": 1, "sheets": 1, "errors": 0, "failed_files": 0, "path": output},
        "run_extract_sheets": {"files": 1, "extracted_sheets": 1, "output_files": 1, "failed_files": 0, "paths": [output]},
        "run_extract_sheets_com": {"files": 1, "extracted_sheets": 1, "output_files": 1, "failed_files": 0, "paths": [output]},
        "run_adjust_rural_loan": {"error": "", "matched": 1, "fill_total": 1.0, "area_total": 1.0, "consistent": True, "unmatched_current": [], "previous_only": []},
        "run_adjust_rural_loan_com": {"error": "", "matched": 1, "fill_total": 1.0, "area_total": 1.0, "consistent": True, "unmatched_current": [], "previous_only": []},
        "run_adjust_rural_loan_quarterly": {"error": "", "regions": [], "matched": 1, "fill_total": 1.0, "area_total": 1.0, "consistent": True},
        "run_adjust_rural_loan_quarterly_com": {"error": "", "regions": [], "matched": 1, "fill_total": 1.0, "area_total": 1.0, "consistent": True},
        "run_precheck": 0,
    }
    calls: dict[str, list[tuple[tuple, dict]]] = {name: [] for name in stats}
    for name, result in stats.items():
        def fake(*args, __name=name, __result=result, **kwargs):
            calls[__name].append((args, kwargs))
            return __result

        monkeypatch.setattr(app, name, fake)

    commands = (
        "5", "1", "1com", "2com", "3com", "4com", "2", "3", "4",
        "6", "7", "8", "9", "a", "b", "p1", "p3", "p7", "p8",
        "p1com", "p3com", "p7com", "ppdf", "t3", "t4", "t5", "t7",
        "t7com", "t8", "t9", "t10", "t11", "t12", "t13", "ar_date",
        "ar_type", "ar_exec", "ar_date_copy", "ar_type_copy", "s21", "s21com",
        "s22", "s22com", "x11", "x12", "x13", "x14", "x15", "x16",
        "x14com", "x16com", "x18", "x19",
    )
    for command in commands:
        app.dispatch(command)
    app.dispatch("x18")
    app.dispatch("x19")

    assert len(commands) == 53
    assert all(calls[name] for name in stats if name != "run_precheck")
    assert calls["run_precheck"]
    assert calls["run_print_config_to_pdf_com"]
    assert calls["run_extract_sheets_com"][0][1]["append_existing"] is True
