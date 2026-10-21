from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pandas as pd
import pytest

from nocompytools.core import io_excel, timeline_396, wide_397_398
from nocompytools.core.config_xlsx import GlobalConfig, PathMapRule, TimelineRule


def _rule(**changes) -> TimelineRule:
    values = {
        "enabled": True,
        "name": "wide",
        "wb_keyword": "book",
        "sheet_keyword": "Data",
        "row_header_col": 1,
        "col_header_row": 1,
        "col_header_rows": [1],
        "data_row_start": 2,
        "data_col_start": 2,
    }
    values.update(changes)
    return TimelineRule(**values)


def _config(tmp_path: Path) -> GlobalConfig:
    return GlobalConfig(tmp_path, tmp_path / "out", tmp_path / "logs")


def _path_map() -> PathMapRule:
    return PathMapRule(True, "map", "*", "", "", "两者", "精确", "raw", "standard")


def test_timeline_deduplicates_cells_and_protects_target_header(tmp_path, monkeypatch):
    source = tmp_path / "book.xlsx"
    source.touch()
    target = tmp_path / "target.xlsx"
    rule = _rule(target_wb_path=target, target_sheet="Timeline", target_write_enabled=True)
    cells = [SimpleNamespace(
        source_file="book.xlsx", sheet_name="Data", rule_name="wide", data_date="2026-05-31",
        row_path="row", col_path="2026-05", value=10,
    )]
    writes, appends = [], []
    monkeypatch.setattr(io_excel, "list_sheet_names", lambda _path: ["Other", "Data"])
    monkeypatch.setattr(timeline_396, "extract_cells", lambda *_args, **_kwargs: cells)
    monkeypatch.setattr(timeline_396, "write_workbook", lambda path, sheets: writes.append((path, sheets)))
    monkeypatch.setattr(timeline_396, "append_to_target", lambda *args, **kwargs: (
        appends.append((args, kwargs)) or {"written": 0, "skipped_reason": "header_mismatch"}
    ))

    output = timeline_396.run_timeline_slim([rule], [_path_map()], _config(tmp_path), [source, source])

    assert output == writes[0][0]
    result = writes[0][1]["时序提取结果"]
    assert len(result) == 1
    assert result.iloc[0]["数值"] == 10
    assert len(appends) == 1
    assert appends[0][1]["required_prefix"] == timeline_396.SLIM_COLS[:6]


@pytest.mark.parametrize("policy,raises", [("continue", False), ("fail_fast", True)])
def test_timeline_sheet_errors_follow_configured_policy(tmp_path, monkeypatch, policy, raises):
    source = tmp_path / "book.xlsx"
    source.touch()
    cfg = _config(tmp_path)
    cfg.error_policy = policy
    rule = _rule(target_write_enabled=False)
    monkeypatch.setattr(
        io_excel,
        "list_sheet_names",
        lambda _path: (_ for _ in ()).throw(OSError("workbook is corrupt")),
    )

    if raises:
        with pytest.raises(OSError, match="corrupt"):
            timeline_396.run_timeline_slim([rule], [], cfg, [source])
    else:
        assert timeline_396.run_timeline_slim([rule], [], cfg, [source]) is None


def test_wide_set_columns_row_numbers_conflicts_and_target_write(tmp_path, monkeypatch):
    source = tmp_path / "book.xlsx"
    source.touch()
    target = tmp_path / "target.xlsx"
    rule = _rule(
        row_header_col_specified=False,
        output_row_no=True,
        set_items=[("机构", "A1")],
        target_wb_path=target,
        target_sheet="Wide",
        target_write_enabled=True,
        target_dedup_cols=["数据日期", "机构"],
    )
    first = SimpleNamespace(
        rule_name="wide", source_file="book.xlsx", sheet_name="Data", data_date="2026-05-31",
        row_path="row", col_path="2026-05", cell_addr="B2", value=10,
    )
    duplicate = SimpleNamespace(**{**vars(first), "cell_addr": "B3", "value": 20})
    written, appends, cleared = [], [], []
    monkeypatch.setattr(wide_397_398, "list_sheet_names", lambda _path: ["Data"])
    monkeypatch.setattr(wide_397_398, "read_sheet_2d", lambda *_args: pd.DataFrame([["机构A", "2026-05"], ["row", 10]]))
    monkeypatch.setattr(wide_397_398, "_load_merge_ranges", lambda *_args: [])
    monkeypatch.setattr(wide_397_398, "compute_wide_col_paths_from_arrays", lambda *_args, **_kwargs: ([], ["2026-05"]))
    monkeypatch.setattr(wide_397_398, "extract_cells", lambda *_args, **_kwargs: [first, duplicate])
    monkeypatch.setattr(wide_397_398, "clear_sheet_cache", lambda: cleared.append(True))
    monkeypatch.setattr(wide_397_398, "write_workbook", lambda path, sheets: written.append((path, sheets)))
    monkeypatch.setattr(wide_397_398, "append_to_target", lambda *args, **kwargs: (
        appends.append((args, kwargs)) or {"written": 2, "input_rows": 2, "batch_dedup_rows": 1,
                                           "existing_filtered_rows": 1}
    ))

    output = wide_397_398.run_wide_summary([rule], [_path_map()], _config(tmp_path), True, [source])

    assert output == written[0][0]
    result = written[0][1]["wide"]
    assert list(result.columns) == ["工作簿名", "工作表名", "数据日期", "机构", "行序号", "2026-05"]
    assert result.iloc[0].tolist() == ["book.xlsx", "Data", "2026-05-31", "机构A", "2", 10]
    assert len(appends) == 1
    assert appends[0][1]["required_prefix"] == ["工作簿名", "工作表名", "数据日期", "机构"]
    assert cleared == [True]


def test_wide_empty_and_invalid_rules_skip_without_output(tmp_path, monkeypatch):
    source = tmp_path / "book.xlsx"
    source.touch()
    bad_set = _rule(name="bad-set", set_parse_error="bad address")
    incompatible = _rule(name="incompatible", row_header_col_specified=False,
                         data_row_start=0, data_col_start=0)
    no_sheet = _rule(name="no-sheet", sheet_keyword="Missing")
    monkeypatch.setattr(wide_397_398, "list_sheet_names", lambda _path: ["Other"])

    assert wide_397_398.run_wide_summary(
        [bad_set, incompatible, no_sheet], [], _config(tmp_path), True, [source]
    ) is None
