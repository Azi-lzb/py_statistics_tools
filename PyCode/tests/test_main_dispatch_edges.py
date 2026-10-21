from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest

from PyCode import main as app


def _prepare_dispatch(monkeypatch, tmp_path):
    cfg = tmp_path / "config.xlsx"
    output, logs = tmp_path / "output", tmp_path / "logs"
    monkeypatch.setattr(app, "CFG_PATH", cfg)
    monkeypatch.setattr(app, "_check", lambda _choice: True)
    monkeypatch.setattr(app, "load_global", lambda _path: SimpleNamespace(output_dir=output, log_dir=logs))
    monkeypatch.setattr(app, "load_global_value", lambda *_args, **_kwargs: "")
    monkeypatch.setattr(app, "load_timeline_rules", lambda _path: [])
    monkeypatch.setattr(app, "load_path_maps", lambda _path: [])
    monkeypatch.setattr(app, "load_dedup_tasks", lambda _path: [])
    monkeypatch.setattr(app, "load_print_tasks", lambda _path: [])
    monkeypatch.setattr(app, "load_extract_tasks", lambda _path: [])
    monkeypatch.setattr(app, "load_institution_mapping", lambda _path: {})
    monkeypatch.setattr(app, "load_foreign_banks", lambda _path: set())
    monkeypatch.setattr(app, "_pick_one_workbook", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(app, "_pick_source_files", lambda *_args, **_kwargs: [])
    monkeypatch.setattr(app, "_pick_source_files_or_folders", lambda *_args, **_kwargs: [])
    monkeypatch.setattr(app, "_pick_archive_paths", lambda *_args, **_kwargs: [])
    monkeypatch.setattr(app, "_pick_archive_target_root", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(app, "run_precheck", lambda *_args: 0)


def test_dispatch_cancellation_and_empty_configuration_returns(monkeypatch, tmp_path, capsys):
    _prepare_dispatch(monkeypatch, tmp_path)
    choices = [
        "1", "1com", "2com", "3com", "4com", "2", "3", "4",
        "6", "7", "p1", "p3", "p1com", "p3com", "t3", "t4", "t5", "t7", "t7com",
        "t8", "t9", "t10", "t11", "t12", "t13", "ar_date", "ar_type", "ar_exec",
        "ar_date_copy", "ar_type_copy", "s21", "s21com", "s22", "s22com",
        "x11", "x12", "x13", "x14", "x15", "x16", "x14com", "x16com", "x18", "x19",
    ]
    for choice in choices:
        app.dispatch(choice)
    report = capsys.readouterr().out
    assert report.count("[已取消]") >= 30
    assert "无启用的时序提取规则" in report
    assert "机构映射表为空" in report
    assert "外资行配置为空" in report
    assert "工作表提取'配置无启用任务" in report


def test_dispatch_handles_not_found_output_and_open_errors(monkeypatch, tmp_path, capsys):
    _prepare_dispatch(monkeypatch, tmp_path)
    app.dispatch("q2")
    assert "未找到输出文件" in capsys.readouterr().out

    output = tmp_path / "output-existing"
    output.mkdir()
    monkeypatch.setattr(app, "load_global", lambda _path: SimpleNamespace(output_dir=output, log_dir=tmp_path / "logs"))
    child = output / "result.xlsx"
    child.touch()
    monkeypatch.setattr(app, "_open_path", lambda _path: (_ for _ in ()).throw(PermissionError("denied")))
    app.dispatch("q2")
    app.dispatch("q3")
    report = capsys.readouterr().out
    assert "打开输出文件失败: denied" in report
    assert "打开输出文件夹失败: denied" in report


def test_dispatch_quick_open_and_config_repair_paths(monkeypatch, tmp_path, capsys):
    _prepare_dispatch(monkeypatch, tmp_path)
    monkeypatch.setattr(app, "initialize_or_repair_config", lambda _path: {
        "created_sheets": 2, "repaired_sheets": 1, "details": ["created", "repaired"]
    })
    app.dispatch("c")

    opened = []
    monkeypatch.setattr(app, "_open_path", lambda path: opened.append(path))
    app.dispatch("q1")
    monkeypatch.setattr(app, "_open_path", lambda _path: (_ for _ in ()).throw(PermissionError("config denied")))
    app.dispatch("q1")
    monkeypatch.setattr(app, "_open_path", lambda path: opened.append(path))
    output = tmp_path / "output"
    output.mkdir()
    latest = output / "latest.xlsx"
    latest.touch()
    monkeypatch.setattr(app, "load_global", lambda _path: SimpleNamespace(output_dir=output, log_dir=tmp_path / "logs"))
    app.dispatch("q2")
    app.dispatch("q3")
    assert opened == [app.CFG_PATH, latest, output]

    monkeypatch.setattr(app, "_check", lambda _choice: False)
    app.dispatch("6")
    app.dispatch("cv")
    report = capsys.readouterr().out
    assert "配置修复" in report
    assert "已打开最新输出文件" in report
    assert "已打开输出文件夹" in report
    assert "打开配置文件失败: config denied" in report


def test_view_config_missing_and_read_error_paths(monkeypatch, tmp_path, capsys):
    _prepare_dispatch(monkeypatch, tmp_path)
    monkeypatch.setattr(app, "CFG_PATH", tmp_path / "missing.xlsx")
    app._view_config()
    assert "请先 5-1 初始化" in capsys.readouterr().out

    cfg = tmp_path / "bad-config.xlsx"
    cfg.touch()
    monkeypatch.setattr(app, "CFG_PATH", cfg)
    import pandas as pd

    class FakeExcelFile:
        sheet_names = ["good", "bad"]

    class FakeFrame:
        shape = (2, 3)

        def __len__(self):
            return 2

    monkeypatch.setattr(pd, "ExcelFile", lambda _path: FakeExcelFile())
    monkeypatch.setattr(pd, "read_excel", lambda _xls, name, **_kwargs: (
        (_ for _ in ()).throw(ValueError("sheet broken")) if name == "bad" else FakeFrame()
    ))
    monkeypatch.setattr(app, "load_global", lambda _path: (_ for _ in ()).throw(ValueError("global broken")))
    monkeypatch.setattr(app, "load_timeline_rules", lambda _path: (_ for _ in ()).throw(ValueError("rules broken")))
    monkeypatch.setattr(app, "load_institution_mapping", lambda _path: (_ for _ in ()).throw(ValueError("map broken")))
    monkeypatch.setattr(app, "load_foreign_banks", lambda _path: set())
    monkeypatch.setattr(app, "load_extract_tasks", lambda _path: (_ for _ in ()).throw(ValueError("tasks broken")))
    app._view_config()
    report = capsys.readouterr().out
    assert "good: 2 行 × 3 列" in report
    assert "bad: 读取失败 sheet broken" in report
    assert "[全局配置解析失败] global broken" in report


def test_view_config_success_loads_all_summaries(monkeypatch, tmp_path, capsys):
    _prepare_dispatch(monkeypatch, tmp_path)
    cfg = tmp_path / "config.xlsx"
    cfg.touch()
    monkeypatch.setattr(app, "CFG_PATH", cfg)
    import pandas as pd

    class FakeExcelFile:
        sheet_names = ["config"]

    class FakeFrame:
        shape = (2, 3)

        def __len__(self):
            return 2

    monkeypatch.setattr(pd, "ExcelFile", lambda _path: FakeExcelFile())
    monkeypatch.setattr(pd, "read_excel", lambda *_args, **_kwargs: FakeFrame())
    monkeypatch.setattr(app, "load_global", lambda _path: SimpleNamespace(
        input_dir=tmp_path / "input", output_dir=tmp_path / "output", log_dir=tmp_path / "logs",
        source_exts=[".xlsx"],
    ))
    monkeypatch.setattr(app, "load_timeline_rules", lambda _path: [SimpleNamespace(
        name="rule", wb_keyword="book", sheet_keyword="sheet"
    )])
    monkeypatch.setattr(app, "load_institution_mapping", lambda _path: {"a": "b"})
    monkeypatch.setattr(app, "load_foreign_banks", lambda _path: {"foreign"})
    monkeypatch.setattr(app, "load_extract_tasks", lambda _path: [{"task": 1}])
    app._view_config()
    report = capsys.readouterr().out
    assert "时序提取规则（启用）: 1 条" in report
    assert "机构映射: 1 条；外资行: 1 个" in report
    assert "工作表提取任务（启用）: 1 条" in report


def test_dispatch_print_slow_task_and_summary_empty_results(monkeypatch, tmp_path, capsys):
    _prepare_dispatch(monkeypatch, tmp_path)
    monkeypatch.setattr(app, "run_print_config_all_modes", lambda *_args: {
        "task_ok": 0, "task_skip": 0, "written_sheets": 0, "written_rows": 0,
        "task_total": 0, "elapsed_sec": 0, "slow_top": [{
            "row_no": 2, "mode": 3, "target": "out.xlsx", "elapsed_sec": 1.25, "status": "slow"
        }],
    })
    app.dispatch("p7")

    monkeypatch.setattr(app, "_pick_source_files", lambda *_args, **_kwargs: [tmp_path / "source.xlsx"])
    monkeypatch.setattr(app, "run_summary_by_usedrange", lambda *_args: {"saved": False})
    app.dispatch("s21")
    routes = iter([([tmp_path / "source.xlsx"], []), ([], [tmp_path / "source.xlsx"])])
    monkeypatch.setattr(app, "_split_sources_for_auto_route", lambda _paths: next(routes))
    app.dispatch("s21com")
    monkeypatch.setattr(app, "_pick_one_workbook", lambda *_args: tmp_path / "template.xlsx")
    monkeypatch.setattr(app, "_pick_source_files", lambda *_args, **_kwargs: [tmp_path / "source.xlsx"])
    monkeypatch.setattr(app, "run_summary_by_comment", lambda *_args: {"saved": False})
    app.dispatch("s22")
    report = capsys.readouterr().out
    assert "最慢任务TOP3" in report
    assert "无可汇总数据，未生成文件" in report


@pytest.mark.parametrize("choice", ["x18", "x19"])
def test_dispatch_rural_adjustment_error_results(monkeypatch, tmp_path, capsys, choice):
    _prepare_dispatch(monkeypatch, tmp_path)
    paths = iter([tmp_path / "current.xlsx", tmp_path / "previous.xlsx"])
    monkeypatch.setattr(app, "_pick_one_workbook", lambda *_args: next(paths))
    monkeypatch.setattr(app, "_split_sources_for_auto_route", lambda items: (list(items), []))
    if choice == "x18":
        monkeypatch.setattr(app, "run_adjust_rural_loan", lambda *_args: {
            "error": "missing sheet", "matched": 0, "fill_total": 0, "area_total": 0,
            "consistent": False, "unmatched_current": [], "previous_only": [],
        })
    else:
        monkeypatch.setattr(app, "run_adjust_rural_loan_quarterly", lambda *_args: {
            "error": "missing sheet", "regions": [], "matched": 0, "fill_total": 0,
            "area_total": 0, "consistent": False,
        })
    app.dispatch(choice)
    assert "[失败] missing sheet" in capsys.readouterr().out


def test_dispatch_timeline_no_rules_cancellation_and_no_match_results(monkeypatch, tmp_path, capsys):
    _prepare_dispatch(monkeypatch, tmp_path)
    monkeypatch.setattr(app, "load_timeline_rules", lambda _path: [object()])
    source = tmp_path / "source.xlsx"
    monkeypatch.setattr(app, "_pick_files_for_compare", lambda: (None, []))
    app.dispatch("1")
    app.dispatch("1com")

    monkeypatch.setattr(app, "_pick_source_files", lambda *_args, **_kwargs: [])
    monkeypatch.setattr(app, "_pick_source_files_or_folders", lambda *_args, **_kwargs: [])
    for choice in ("2com", "3com", "4com", "2", "3", "4"):
        app.dispatch(choice)

    monkeypatch.setattr(app, "_pick_files_for_compare", lambda: (source, [source]))
    monkeypatch.setattr(app, "_pick_source_files", lambda *_args, **_kwargs: [source])
    monkeypatch.setattr(app, "_pick_source_files_or_folders", lambda *_args, **_kwargs: [source])
    monkeypatch.setattr(app, "_split_sources_for_auto_route", lambda paths: (list(paths), []))
    monkeypatch.setattr(app, "run_compare", lambda *_args: None)
    monkeypatch.setattr(app, "run_timeline_slim", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(app, "run_wide_summary", lambda *_args, **_kwargs: None)
    app.dispatch("1com")
    app.dispatch("2com")
    app.dispatch("3com")
    app.dispatch("4com")
    app.dispatch("2")
    app.dispatch("3")
    app.dispatch("4")

    report = capsys.readouterr().out
    assert "无启用的时序提取规则" not in report
    assert "无匹配结果，未生成文件" in report


def test_dispatch_archive_cancellation_after_path_selection(monkeypatch, tmp_path, capsys):
    _prepare_dispatch(monkeypatch, tmp_path)
    source = tmp_path / "source.xlsx"
    monkeypatch.setattr(app, "_pick_archive_paths", lambda *_args: [source])
    monkeypatch.setattr(app, "_pick_archive_target_root", lambda: None)
    for choice in ("ar_date", "ar_type", "ar_date_copy", "ar_type_copy"):
        app.dispatch(choice)
    assert capsys.readouterr().out.count("[已取消]") == 4


def test_dispatch_reports_office_com_runtime_errors(monkeypatch, tmp_path, capsys):
    _prepare_dispatch(monkeypatch, tmp_path)
    source = tmp_path / "source.xlsx"
    source.touch()
    monkeypatch.setattr(app, "load_timeline_rules", lambda _path: [object()])
    monkeypatch.setattr(app, "load_print_tasks", lambda _path: [object()])
    monkeypatch.setattr(app, "load_extract_tasks", lambda _path: [{"area_spec": "A1"}])
    monkeypatch.setattr(app, "load_institution_mapping", lambda _path: {"raw": "standard"})
    monkeypatch.setattr(app, "load_foreign_banks", lambda _path: {"foreign"})
    monkeypatch.setattr(app, "_pick_files_for_compare", lambda: (source, [source]))
    monkeypatch.setattr(app, "_pick_one_workbook", lambda *_args: source)
    monkeypatch.setattr(app, "_pick_source_files", lambda *_args, **_kwargs: [source])
    monkeypatch.setattr(app, "_pick_source_files_or_folders", lambda *_args, **_kwargs: [source])
    monkeypatch.setattr(app, "_pick_archive_paths", lambda *_args: [source])
    monkeypatch.setattr(app, "_pick_archive_target_root", lambda: tmp_path / "archive")
    monkeypatch.setattr(app, "_split_sources_for_auto_route", lambda paths: ([], list(paths)))
    monkeypatch.setattr(app, "_split_print_tasks_for_auto_route", lambda tasks: ([], list(tasks)))
    error = lambda *_args, **_kwargs: (_ for _ in ()).throw(RuntimeError("office unavailable"))
    for name in (
        "run_compare_com", "run_timeline_slim_com", "run_wide_summary_com",
        "run_print_keep_by_comment_com", "run_print_fast_by_comment_com",
        "run_print_config_all_modes_com", "run_print_config_to_pdf_com",
        "run_batch_rename_sheet_37_com", "run_summary_by_usedrange_com",
        "run_summary_by_comment_com", "run_split_village_bank_com",
        "run_normalize_institution_com", "run_remove_foreign_bank_com",
        "run_fx_header_fix_com", "run_region_sum_check_com", "run_extract_sheets_com",
        "run_adjust_rural_loan_com", "run_adjust_rural_loan_quarterly_com",
    ):
        monkeypatch.setattr(app, name, error)

    choices = (
        "1com", "2com", "3com", "4com", "p1com", "p3com", "p7com", "ppdf", "t7com",
        "s21com", "s22com", "x11", "x12", "x13", "x14com", "x15", "x16com", "x18", "x19",
    )
    for choice in choices:
        app.dispatch(choice)
    report = capsys.readouterr().out
    assert report.count("[COM 不可用] office unavailable") >= 15


def test_main_routes_every_submenu_mapping(monkeypatch, capsys):
    main_values = []
    sub_values = []
    expected = []
    menus = {
        "1": [("1", "1com"), ("2", "2com"), ("3", "3com"), ("4", "4com"), ("5", "5")],
        "2": [("1", "6"), ("2", "7"), ("3", "8"), ("4", "9"), ("5", "a"), ("6", "b")],
        "3": [("1", "p1com"), ("2", "p3com"), ("3", "p7com"), ("4", "p8"), ("5", "ppdf")],
        "4": [("1", "t3"), ("2", "t4"), ("3", "t5"), ("4", "t7com"), ("5", "t8"),
              ("6", "t9"), ("7", "t10"), ("8", "t11"), ("9", "t12"), ("10", "t13")],
        "5": [("1", "c"), ("2", "cv")],
        "6": [("1", "s21com"), ("2", "s22com")],
        "7": [("1", "x11"), ("2", "x12"), ("3", "x13"), ("4", "x14com"), ("5", "x15"),
              ("6", "x16com"), ("7", "x18"), ("8", "x19")],
        "8": [("1", "ar_date"), ("2", "ar_type"), ("3", "ar_exec"), ("4", "ar_date_copy"), ("5", "ar_type_copy")],
        "9": [("1", "q1"), ("2", "q2"), ("3", "q3")],
    }
    for menu, items in menus.items():
        for sub, mapped in items:
            main_values.append(menu)
            sub_values.append(sub)
            expected.append(mapped)
    main_values.append("0")
    main_answers, sub_answers = iter(main_values), iter(sub_values)
    called = []
    monkeypatch.setattr(app, "_select_engine_mode", lambda: "python")
    monkeypatch.setattr(app, "_show_main_menu", lambda: next(main_answers))
    monkeypatch.setattr(app, "_show_sub_menu", lambda *_args: next(sub_answers))
    monkeypatch.setattr(app, "dispatch", called.append)

    app.main()

    assert called == expected
    capsys.readouterr()


def test_main_returns_from_each_submenu_and_reports_dispatch_exceptions(monkeypatch, capsys):
    main_answers = iter(["1", "2", "3", "4", "5", "6", "7", "8", "9", "0"])
    monkeypatch.setattr(app, "_select_engine_mode", lambda: "python")
    monkeypatch.setattr(app, "_show_main_menu", lambda: next(main_answers))
    monkeypatch.setattr(app, "_show_sub_menu", lambda *_args: "0")
    monkeypatch.setattr(app, "dispatch", lambda _choice: pytest.fail("cancelled submenu must not dispatch"))
    app.main()

    main_answers = iter(["9", "0"])
    sub_answers = iter(["2"])
    monkeypatch.setattr(app, "_show_main_menu", lambda: next(main_answers))
    monkeypatch.setattr(app, "_show_sub_menu", lambda *_args: next(sub_answers))
    monkeypatch.setattr(app, "dispatch", lambda _choice: (_ for _ in ()).throw(ValueError("dispatch error")))
    app.main()
    captured = capsys.readouterr()
    assert "[运行异常] dispatch error" in captured.out
    assert "ValueError: dispatch error" in captured.err
