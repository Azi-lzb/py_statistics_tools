from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pandas as pd
import pytest

from PyCode.core import config_xlsx as config


def _write(path: Path, sheets: dict[str, pd.DataFrame]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name, index=False)
    return path


def _global_frame(rows=None, columns=("键", "值")) -> pd.DataFrame:
    rows = rows if rows is not None else [
        ["输入目录", "input"], ["输出目录", "output"], ["日志目录", "logs"],
    ]
    return pd.DataFrame(rows, columns=list(columns))


def test_config_scalar_parsers_cover_empty_invalid_and_unicode_separators() -> None:
    assert config._truthy(None) is False
    assert config._truthy(" On ") and config._truthy("YES")
    assert config._truthy("否") is False
    assert config._split(None) == []
    assert config._split(float("nan")) == []
    assert config._split("  ") == []
    assert config._split("A，B; C；D") == ["A", "B", "C", "D"]
    assert config._to_int(None, 7) == 7
    assert config._to_int(float("nan"), 7) == 7
    assert config._to_int("2.8") == 2
    assert config._to_int("bad", 7) == 7
    assert config._to_int_or_none(None) is None
    assert config._to_int_or_none(float("nan")) is None
    assert config._to_int_or_none(" ") is None
    assert config._to_int_or_none("4.9") == 4
    assert config._to_int_or_none("bad") is None
    assert config._to_str(float("nan")) == ""
    assert config._to_str(" x ") == "x"


def test_set_area_parser_validates_format_alias_and_cell_address() -> None:
    assert config._parse_set_items(None) == ([], "")
    assert config._parse_set_items(" name@a1；total@BC99 ") == ([('name', 'A1'), ('total', 'BC99')], "")
    assert "格式非法" in config._parse_set_items("missing-at")[1]
    assert "别名为空" in config._parse_set_items("@A1")[1]
    assert "地址非法" in config._parse_set_items("name@A0")[1]


def test_feature_validation_reports_missing_file_workbook_sheets_columns_and_keys(tmp_path: Path, monkeypatch) -> None:
    assert config.validate_sheets_for_feature("p1", tmp_path / "missing.xlsx")[0].startswith("配置文件不存在")
    corrupt = tmp_path / "corrupt.xlsx"
    corrupt.write_text("not xlsx", encoding="utf-8")
    assert config.validate_sheets_for_feature("p1", corrupt)[0].startswith("无法打开配置文件")

    no_dedup = _write(tmp_path / "no-dedup.xlsx", {"Other": pd.DataFrame({"x": [1]})})
    assert "或 数据处理配置" in config.validate_sheets_for_feature("8", no_dedup)[0]
    assert config.validate_sheets_for_feature("unknown", no_dedup) == []
    has_dedup = _write(tmp_path / "has-dedup.xlsx", {
        config.SHEET_DEDUP_TASK: pd.DataFrame(columns=config.DEDUP_TASK_COLS),
    })
    assert config.validate_sheets_for_feature("8", has_dedup) == []

    one_column = _write(tmp_path / "one-column.xlsx", {
        config.SHEET_GLOBAL: pd.DataFrame({"键": ["输入目录"]}),
    })
    assert any("至少需要 键/值 两列" in item for item in config.validate_sheets_for_feature("p1", one_column))

    missing_keys = _write(tmp_path / "missing-keys.xlsx", {
        config.SHEET_GLOBAL: _global_frame([["输入目录", "in"]]),
    })
    errors = config.validate_sheets_for_feature("p1", missing_keys)
    assert len([item for item in errors if "缺少键" in item]) == 2

    short_headers = _write(tmp_path / "short-headers.xlsx", {
        config.SHEET_GLOBAL: _global_frame(),
        config.SHEET_PATH_MAP: pd.DataFrame(columns=["映射名称"]),
    })
    assert "缺少列" in " ".join(config.validate_sheets_for_feature("2", short_headers))
    assert any("缺少 Sheet" in item for item in config.validate_sheets_for_feature("p8", short_headers))

    calls = 0

    def fail_on_second_open(_path):
        nonlocal calls
        calls += 1
        if calls == 1:
            return SimpleNamespace(sheet_names=[config.SHEET_DEDUP_TASK])
        raise OSError("disappeared during validation")

    monkeypatch.setattr(config.pd, "ExcelFile", fail_on_second_open)
    assert "disappeared during validation" in config.validate_sheets_for_feature("8", no_dedup)[0]


def test_global_config_resolves_relative_absolute_paths_and_defaults(tmp_path: Path) -> None:
    cfg = tmp_path / "config" / "config.xlsx"
    absolute = tmp_path / "absolute-input"
    _write(cfg, {config.SHEET_GLOBAL: _global_frame([
        ["输入目录", "relative/in"], ["输出目录", str(tmp_path / "absolute-out")],
        ["日志目录", ""], ["源文件扩展名", "XLSX； csv, .xlsm"],
        ["时序规则命中策略", " FIRST_MATCH "], ["错误策略", "stop"], ["默认编码", "gbk"],
    ])})

    loaded = config.load_global(cfg)

    base = cfg.parent.parent
    assert loaded.input_dir == (base / "relative/in").resolve()
    assert loaded.output_dir == tmp_path / "absolute-out"
    assert loaded.log_dir == (base / "logs").resolve()
    assert loaded.source_exts == [".xlsx", ".csv", ".xlsm"]
    assert loaded.timeline_rule_match_mode == "first_match"
    assert loaded.error_policy == "stop" and loaded.default_encoding == "gbk"

    _write(cfg, {config.SHEET_GLOBAL: _global_frame([
        ["输入目录", str(absolute)], ["源文件扩展名", ""], ["规则命中策略", "unknown"],
    ])})
    defaults = config.load_global(cfg)
    assert defaults.input_dir == absolute
    assert defaults.source_exts == [".xlsx", ".xlsm", ".csv"]
    assert defaults.timeline_rule_match_mode == "all_match"
    assert config.load_global_value(cfg, "not-there", "fallback") == "fallback"

    one_col = _write(tmp_path / "one-col.xlsx", {config.SHEET_GLOBAL: pd.DataFrame({"键": ["x"]})})
    assert config.load_global_value(one_col, "x", "fallback") == "fallback"


def test_timeline_loader_parses_multiline_headers_optional_columns_and_set_errors(tmp_path: Path) -> None:
    cfg = tmp_path / "timeline.xlsx"
    row1 = {column: "" for column in config.TIMELINE_COLS}
    row1.update({
        "是否启用": "Y", "规则名称": "multi", "行头列": "A,C", "列表头行": "2，4;5",
        "必含列头": "机构；金额", "必含行头": "惠州,惠阳", "数据起始行": "3.8",
        "数据结束行": "bad", "数据起始列": "B", "数据结束列": "D",
        "跳过关键字": "小计；合计", "目标工作簿路径": "out.xlsx", "目标工作表": "总表",
        "启用目标写入": "N", "set区域": "bad-set", "目标去重列": "A；C", "是否输出行序号": "Y",
    })
    row2 = {column: "" for column in config.TIMELINE_REQUIRED_COLS}
    row2.update({"是否启用": "Y", "规则名称": "legacy", "列表头行": "bad,0", "行头列": ""})
    row3 = {column: "" for column in config.TIMELINE_COLS}
    row3.update({"是否启用": "N", "规则名称": "disabled"})
    _write(cfg, {config.SHEET_TIMELINE_RULE: pd.DataFrame([row1, row2, row3])})

    rules = config.load_timeline_rules(cfg)

    assert len(rules) == 2
    assert rules[0].row_header_cols == [1, 3] and rules[0].row_header_col_specified
    assert rules[0].col_header_rows == [2, 4, 5]
    assert rules[0].required_col_headers == ["机构", "金额"]
    assert rules[0].required_row_headers == ["惠州", "惠阳"]
    assert rules[0].data_row_start == 3 and rules[0].data_row_end is None
    assert rules[0].data_col_start == 2 and rules[0].data_col_end == 4
    assert rules[0].skip_keywords == ["小计", "合计"]
    assert rules[0].target_wb_path == Path("out.xlsx") and rules[0].target_sheet == "总表"
    assert rules[0].target_write_enabled is False
    assert rules[0].set_items == [] and "格式非法" in rules[0].set_parse_error
    assert rules[0].target_dedup_cols == ["A", "C"] and rules[0].output_row_no
    assert rules[1].row_header_cols == [] and not rules[1].row_header_col_specified
    assert rules[1].col_header_rows == [1]

    legacy_cfg = _write(tmp_path / "legacy-timeline.xlsx", {
        config.SHEET_TIMELINE_RULE: pd.DataFrame([row2], columns=config.TIMELINE_REQUIRED_COLS),
    })
    assert config.load_timeline_rules(legacy_cfg)[0].target_write_enabled


def test_path_map_dedup_print_and_extract_loaders_cover_fallback_and_legacy_columns(tmp_path: Path) -> None:
    cfg = tmp_path / "config" / "loaders.xlsx"
    path_rows = [
        ["Y", "map", "rule", "wb", "sheet", "", "", "old", "new", "note"],
        ["N", "disabled", "", "", "", "", "", "x", "y", ""],
    ]
    dedup_rows = [
        ["Y", "task", "source.xlsx", "Data", "1,2;bad", "target.xlsx", "Out", "append", "note"],
        ["N", "disabled", "", "", "", "", "", "", ""],
    ]
    print_rows = [["Y", "2", "source.xlsx", "Data", "A1:B2", "out.xlsx", "Output", "bad", "0",
                   "横向", "yes", "否", "Y", "1", "remark"]]
    extract_rows = [
        ["N", "ignored", "", "", "", "", "", "", ""],
        ["Y", "", "", "", "", "", "", "", "empty"],
        ["Y", "", "sheet", "exclude", "", "horizontal", "out", "name", "remark"],
    ]
    _write(cfg, {
        config.SHEET_PATH_MAP: pd.DataFrame(path_rows, columns=config.PATH_MAP_COLS),
        config.SHEET_DEDUP_TASK: pd.DataFrame(dedup_rows, columns=[
            "是否启用", "任务名", "源数据工作簿", "源数据工作表", "标识列序号", "目标工作簿", "目标工作表", "写入方式", "备注"
        ]),
        config.SHEET_PRINT_CONFIG: pd.DataFrame(print_rows, columns=config.PRINT_CONFIG_COLS),
        config.SHEET_EXTRACT_CONFIG: pd.DataFrame(extract_rows, columns=config.EXTRACT_CONFIG_COLS),
    })

    maps = config.load_path_maps(cfg)
    assert len(maps) == 1 and maps[0].target == "两者" and maps[0].match_mode == "精确"
    assert maps[0].original == "old" and maps[0].standard == "new"
    tasks = config.load_dedup_tasks(cfg)
    assert len(tasks) == 2 and tasks[0]["enabled"] and not tasks[1]["enabled"]
    assert tasks[0]["key_cols"] == [1, 2] and tasks[0]["write_mode"] == "append"
    assert tasks[0]["source_wb"] == (cfg.parent.parent / "source.xlsx").resolve()
    prints = config.load_print_tasks(cfg)
    assert len(prints) == 1 and prints[0]["mode"] == 2
    assert prints[0]["fit_wide"] == 1 and prints[0]["fit_tall"] == 0
    assert prints[0]["center_h"] and prints[0]["drop_comments"] and prints[0]["hide_zero"]
    absolute_source = (tmp_path / "absolute-source.xlsx").resolve()
    absolute_target = (tmp_path / "absolute-target.xlsx").resolve()
    absolute_print = _write(tmp_path / "absolute-print.xlsx", {
        config.SHEET_PRINT_CONFIG: pd.DataFrame([[
            "Y", 1, str(absolute_source), "Data", "", str(absolute_target), "Out", 1, 1,
            "", "", "", "", "", "",
        ]], columns=config.PRINT_CONFIG_COLS),
    })
    absolute_task = config.load_print_tasks(absolute_print)[0]
    assert absolute_task["source_wb"] == absolute_source and absolute_task["target_wb"] == absolute_target
    extracts = config.load_extract_tasks(cfg)
    assert len(extracts) == 1 and extracts[0]["row_no"] == 4
    assert extracts[0]["sheet_keywords"] == "sheet" and extracts[0]["output_name"] == "name"

    old_cfg = _write(tmp_path / "legacy.xlsx", {
        config.SHEET_DEDUP_TASK_LEGACY: pd.DataFrame(
            [["Y", "source.xlsx", "Data", "2:3", "target.xlsx", "Out"]],
            columns=["启用", "源工作簿路径", "源工作表名", "去重列", "目标工作簿路径", "目标工作表名"],
        ),
    })
    legacy = config.load_dedup_tasks(old_cfg)
    assert legacy[0]["enabled"] and legacy[0]["key_cols"] == [2, 3]
    assert legacy[0]["task_name"] == "" and legacy[0]["write_mode"] == ""
    assert config.load_dedup_tasks(_write(tmp_path / "no-dedup.xlsx", {"Other": pd.DataFrame({"x": [1]})})) == []
    assert config.load_path_maps(_write(tmp_path / "no-maps.xlsx", {"Other": pd.DataFrame({"x": [1]})})) == []
    assert config.load_print_tasks(_write(tmp_path / "no-print.xlsx", {"Other": pd.DataFrame({"x": [1]})})) == []
    assert config.load_extract_tasks(_write(tmp_path / "no-extract.xlsx", {"Other": pd.DataFrame({"x": [1]})})) == []


def test_institution_loaders_ignore_missing_short_and_invalid_rows(tmp_path: Path) -> None:
    missing = _write(tmp_path / "missing.xlsx", {"Other": pd.DataFrame({"x": [1]})})
    assert config.load_foreign_banks(missing) == set()
    assert config.load_institution_mapping(missing) == {}

    short = _write(tmp_path / "short.xlsx", {
        config.SHEET_INSTITUTION_MAPPING: pd.DataFrame([["source", "standard"]], columns=["raw", "mapped"]),
    })
    assert config.load_foreign_banks(short) == set()
    assert config.load_institution_mapping(short) == {"source": "standard"}

    full = _write(tmp_path / "full.xlsx", {
        config.SHEET_INSTITUTION_MAPPING: pd.DataFrame([
            ["raw1", "mapped1", "1"], ["raw2", "mapped2", "0"], ["", "nameless", "1"],
            ["raw3", "mapped3", "yes"], ["raw4", "", "1"],
        ], columns=config.INSTITUTION_MAPPING_COLS),
    })
    assert config.load_foreign_banks(full) == {"mapped1", "nameless"}
    assert config.load_institution_mapping(full) == {"raw1": "mapped1", "raw2": "mapped2", "raw3": "mapped3"}


def test_helper_path_and_key_column_fallbacks(tmp_path: Path) -> None:
    frame = pd.DataFrame(columns=["secondary"])
    assert config._find_col(frame, ["first", "secondary"]) == "secondary"
    assert config._find_col(frame, ["missing"]) is None
    assert config._parse_key_cols(None) == []
    assert config._parse_key_cols(float("nan")) == []
    assert config._parse_key_cols(" ") == []
    assert config._parse_key_cols("1,, 2；0;bad:-3") == [1, 2]
    base = tmp_path / "base"
    assert config._resolve_any_path("relative/file.xlsx", base) == (base / "relative/file.xlsx").resolve()
    absolute = (tmp_path / "absolute.xlsx").resolve()
    assert config._resolve_any_path(str(absolute), base) == absolute
