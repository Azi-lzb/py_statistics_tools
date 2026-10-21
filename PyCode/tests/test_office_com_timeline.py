from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest

from PyCode.core.compare_393 import run_compare
from PyCode.core.compare_393_com import run_compare_com
from PyCode.core.config_xlsx import load_global, load_path_maps, load_timeline_rules
from PyCode.core.timeline_396 import run_timeline_slim
from PyCode.core.timeline_396_com import run_timeline_slim_com
from PyCode.core.wide_397_398 import run_wide_summary
from PyCode.core.wide_397_398_com import run_wide_summary_com
from PyCode.tests.test_e2e import _make_config, _make_source


pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="Office COM is Windows-only")


def _records(path: Path, sheet: str) -> list[tuple]:
    frame = pd.read_excel(path, sheet_name=sheet, dtype=object)
    frame = frame.where(pd.notna(frame), None)
    records = []
    for row in frame.itertuples(index=False, name=None):
        records.append(tuple(v.date().isoformat() if isinstance(v, datetime) else v for v in row))
    return records


def _first_match(root: Path, pattern: str) -> Path:
    paths = sorted(root.glob(pattern))
    assert paths, f"no output matching {pattern} under {root}"
    return paths[-1]


def test_com_compare_timeline_and_both_wide_modes_match_python(tmp_path: Path) -> None:
    base = tmp_path / "pytools_test"
    input_dir = base / "input"
    output_dir = base / "output"
    log_dir = base / "logs"
    cfg = base / "config" / "config.xlsx"
    _make_config(cfg, input_dir, output_dir, log_dir)

    template_dir = base / "tmpl"
    input_dir.mkdir(parents=True, exist_ok=True)
    template_dir.mkdir(parents=True, exist_ok=True)
    template = template_dir / "template.xlsx"
    data = [
        ["2026年4月本外币存款", "", ""],
        ["项目", "本月余额", "比上月"],
        ["活期存款", 100, 5],
        ["定期存款", 200, -3],
        ["合计", 300, 2],
    ]
    source = input_dir / "本外币存款表_202604.xlsx"
    _make_source(template, "本外币", data)
    _make_source(source, "本外币", data)

    global_cfg = load_global(cfg)
    rules = load_timeline_rules(cfg)
    path_maps = load_path_maps(cfg)
    assert len(rules) == 1

    py_slim = run_timeline_slim(rules, path_maps, global_cfg, source_paths=[source])
    com_slim = run_timeline_slim_com(rules, path_maps, global_cfg, source_paths=[source])
    assert py_slim is not None and com_slim is not None
    assert _records(py_slim, "时序提取结果") == _records(com_slim, "时序提取结果")

    for suffix_enabled in (True, False):
        py_wide = run_wide_summary(
            rules, path_maps, global_cfg, row_suffix_enabled=suffix_enabled, source_paths=[source]
        )
        com_wide = run_wide_summary_com(
            rules, path_maps, global_cfg, row_suffix_enabled=suffix_enabled, source_paths=[source]
        )
        assert py_wide is not None and com_wide is not None
        assert _records(py_wide, "存款规则") == _records(com_wide, "存款规则")

    # 对比一份通过样本和一份增加行头的差异样本，核对差异行内容和汇总计数。
    py_pass = run_compare(template, [source], rules, global_cfg)
    com_pass = run_compare_com(template, [source], rules, global_cfg)
    assert _records(py_pass, "汇总") == _records(com_pass, "汇总")
    assert _records(py_pass, "表头比对结果") == _records(com_pass, "表头比对结果")

    changed = input_dir / "本外币存款表_202607.xlsx"
    changed_data = [*data[:-1], ["新增项目", 1, 1], data[-1]]
    _make_source(changed, "本外币", changed_data)
    py_diff = run_compare(template, [changed], rules, global_cfg)
    com_diff = run_compare_com(template, [changed], rules, global_cfg)
    assert _records(py_diff, "汇总") == _records(com_diff, "汇总")
    assert _records(py_diff, "表头比对结果") == _records(com_diff, "表头比对结果")
