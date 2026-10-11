from __future__ import annotations

from pathlib import Path

from .com_engine import safe_close, safe_quit, start_excel_app
from .feature_1_8_adjust_rural_loan import (
    HEADER_SCAN_ROWS,
    HEADER_TEXT,
    INST_DATA_START_ROW,
    INST_DIFF_COL,
    INST_NAME_COL,
    INST_VALUE_COL,
    NAME_SCAN_COLS,
    TOLERANCE,
    _normalize,
    _to_num,
)
from .feature_1_9_adjust_rural_loan_quarterly import (
    AREA_SHEET_KEY,
    INSTITUTION_SHEET_KEY,
    _region_from_sheet_name,
)
from .logger import get_logger


def _last_row(ws) -> int:
    used = ws.UsedRange
    return int(used.Row) + int(used.Rows.Count) - 1


def _last_col(ws) -> int:
    used = ws.UsedRange
    return int(used.Column) + int(used.Columns.Count) - 1


def _institution_sheets(wb) -> dict[str, object]:
    sheets: dict[str, object] = {}
    for ws in wb.Worksheets:
        title = str(ws.Name)
        if INSTITUTION_SHEET_KEY not in title:
            continue
        region = _region_from_sheet_name(title)
        if not region:
            raise ValueError(f"无法从分机构Sheet名识别地区：{title}")
        if region in sheets:
            raise ValueError(f"同一工作簿存在重复地区分机构Sheet：{region}")
        sheets[region] = ws
    return sheets


def _find_area_sheet(wb):
    hits = [ws for ws in wb.Worksheets if AREA_SHEET_KEY in str(ws.Name)]
    if len(hits) > 1:
        raise ValueError(f"匹配到多个“{AREA_SHEET_KEY}”工作表，请保留唯一目标表。")
    return hits[0] if hits else None


def _build_institution_map(ws) -> tuple[dict[str, float], dict[str, int], list[str]]:
    values: dict[str, float] = {}
    rows: dict[str, int] = {}
    duplicates: list[str] = []
    for row_no in range(INST_DATA_START_ROW, _last_row(ws) + 1):
        name = _normalize(ws.Cells(row_no, INST_NAME_COL).Value)
        if not name:
            continue
        if name in values:
            duplicates.append(name)
            continue
        values[name] = _to_num(ws.Cells(row_no, INST_VALUE_COL).Value)
        rows[name] = row_no
    return values, rows, duplicates


def _find_header_col(ws, header: str) -> int:
    for row_no in range(1, min(_last_row(ws), HEADER_SCAN_ROWS) + 1):
        for col_no in range(1, _last_col(ws) + 1):
            if _normalize(ws.Cells(row_no, col_no).Value) == header:
                return col_no
    return 0


def _find_region_row(ws, region: str) -> int:
    for row_no in range(1, _last_row(ws) + 1):
        for col_no in range(1, NAME_SCAN_COLS + 1):
            if _normalize(ws.Cells(row_no, col_no).Value) == region:
                return row_no
    return 0


def run_adjust_rural_loan_quarterly_com(current_path: Path, previous_path: Path, log_dir: Path) -> dict:
    logger = get_logger("feature_1_9_adjust_rural_loan_quarterly_com", log_dir)
    result = {
        "regions": [], "matched": 0, "fill_total": 0.0, "area_total": 0.0,
        "consistent": False, "error": None,
    }
    app = start_excel_app()
    wb_current = wb_previous = None
    try:
        wb_current = app.Workbooks.Open(str(current_path.resolve()), ReadOnly=False, UpdateLinks=0)
        wb_previous = app.Workbooks.Open(str(previous_path.resolve()), ReadOnly=True, UpdateLinks=0)
        current_sheets = _institution_sheets(wb_current)
        previous_sheets = _institution_sheets(wb_previous)
        if not current_sheets:
            raise ValueError("本期文件未找到包含“涉农贷款分机构”的工作表。")
        if not previous_sheets:
            raise ValueError("上期文件未找到包含“涉农贷款分机构”的工作表。")
        area_sheet = _find_area_sheet(wb_current)
        if area_sheet is None:
            raise ValueError("本期文件未找到包含“涉农贷款分地区统计”的工作表。")
        header_col = _find_header_col(area_sheet, HEADER_TEXT)
        if not header_col:
            raise ValueError(f'未在“{area_sheet.Name}”中找到表头“{HEADER_TEXT}”。')

        for region, current_sheet in current_sheets.items():
            item = {
                "region": region, "sheet": str(current_sheet.Name), "matched": 0,
                "fill_total": 0.0, "area_total": None, "consistent": False,
                "unmatched_current": [], "previous_only": [], "error": None,
            }
            previous_sheet = previous_sheets.get(region)
            if previous_sheet is None:
                item["error"] = f"上期文件缺少{region}分机构表"
                result["regions"].append(item)
                continue
            current_values, current_rows, current_duplicates = _build_institution_map(current_sheet)
            previous_values, _, previous_duplicates = _build_institution_map(previous_sheet)
            if current_duplicates or previous_duplicates:
                item["error"] = "分机构表存在重复机构名称：" + "、".join(current_duplicates + previous_duplicates)
                result["regions"].append(item)
                continue
            for name, row_no in current_rows.items():
                if name not in previous_values:
                    current_sheet.Cells(row_no, INST_DIFF_COL).Value = None
                    item["unmatched_current"].append(name)
                    continue
                difference = current_values[name] - previous_values[name]
                current_sheet.Cells(row_no, INST_DIFF_COL).Value = difference
                item["matched"] += 1
                item["fill_total"] += difference
            item["previous_only"] = [name for name in previous_values if name not in current_values]
            area_row = _find_region_row(area_sheet, region)
            if not area_row:
                item["error"] = f'“{area_sheet.Name}”未找到地区行：{region}'
            else:
                item["area_total"] = _to_num(area_sheet.Cells(area_row, header_col).Value)
                item["consistent"] = abs(item["fill_total"] - item["area_total"]) <= TOLERANCE
            result["matched"] += item["matched"]
            result["fill_total"] += item["fill_total"]
            if item["area_total"] is not None:
                result["area_total"] += item["area_total"]
            result["regions"].append(item)

        wb_current.Save()
        result["consistent"] = bool(result["regions"]) and all(
            item["consistent"] and not item["error"] for item in result["regions"]
        )
    except Exception as exc:
        result["error"] = str(exc)
        logger.error("失败: %s", exc)
    finally:
        safe_close(wb_previous)
        safe_close(wb_current)
        safe_quit(app)
    return result
