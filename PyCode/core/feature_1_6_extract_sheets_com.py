"""1.6 提取工作表数据 — COM 版本

策略：
- full_extract=True：用 src_ws.Copy(After=...) 整表过去，再把公式就地转为值，**保留批注/格式/合并/列宽，公式按源结果落地**
- full_extract=False：用 Range.Value2 批量读值 + 行列过滤 + 紧凑后写入新 sheet
  （部分提取场景需要"剔除空行/空列"，这一步要求纯值视图，因此不再保留格式）
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Iterable

from .com_engine import (
    open_readonly,
    safe_close,
    safe_quit,
    start_excel_app,
)
from .feature_1_6_extract_sheets import (
    _compact,
    _extract_configured_values,
    _is_full_sheet_task,
    _name_matches_exclude,
    _name_matches_keyword_options,
    _output_path_for_task,
    _sanitize_sheet_name,
    _workbook_matches,
)
from .logger import get_logger


XL_OPEN_XML_WORKBOOK = 51


def _list_sheet_names(wb) -> list[str]:
    return [str(ws.Name) for ws in wb.Worksheets]


def _find_matched_sheets_com(wb, task: dict):
    """按新版工作表关键字匹配。分号为 AND，逗号/竖线为 OR。"""
    matched = []
    seen: set[str] = set()
    for ws in wb.Worksheets:
        title = str(ws.Name)
        if title in seen:
            continue
        if _name_matches_exclude(title, task):
            continue
        if _name_matches_keyword_options(title, str(task.get("sheet_keywords") or "")):
            matched.append(ws)
            seen.add(title)
    return matched


def _unique_sheet_name_com(wb, base: str) -> str:
    base = _sanitize_sheet_name(base)
    existing = set(_list_sheet_names(wb))
    if base not in existing:
        return base
    root = base[:28] if len(base) > 28 else base
    for i in range(2, 101):
        cand = f"{root}_{i}"
        if cand not in existing:
            return cand
    return f"{root}_{datetime.now().strftime('%H%M%S')}"


def _add_clean_workbook(app):
    """新建只含 1 张默认 sheet 的工作簿。"""
    old = None
    try:
        old = app.SheetsInNewWorkbook
        app.SheetsInNewWorkbook = 1
    except Exception:
        pass
    wb = app.Workbooks.Add()
    if old is not None:
        try:
            app.SheetsInNewWorkbook = old
        except Exception:
            pass
    return wb


def _read_full_2d(src_ws) -> list[list]:
    """整张 sheet 的值（用于 partial 提取）。"""
    ur = src_ws.UsedRange
    last_row = int(ur.Row) + int(ur.Rows.Count) - 1
    last_col = int(ur.Column) + int(ur.Columns.Count) - 1
    if last_row < 1 or last_col < 1:
        return []
    rng = src_ws.Range(src_ws.Cells(1, 1), src_ws.Cells(last_row, last_col))
    v = rng.Value2
    if v is None:
        return []
    if not isinstance(v, tuple):
        return [[v]]
    return [list(row) for row in v]


def _write_values_com(dst_ws, values: list[list]) -> None:
    if not values:
        return
    nrow = len(values)
    ncol = max((len(r) for r in values), default=0)
    if ncol == 0:
        return
    # 对齐为等长 tuple of tuples，pywin32 需要规整结构
    payload = tuple(
        tuple(row[i] if i < len(row) else None for i in range(ncol))
        for row in values
    )
    dst_ws.Range(dst_ws.Cells(1, 1), dst_ws.Cells(nrow, ncol)).Value2 = payload


def _sheet_has_content_com(ws) -> bool:
    try:
        values = ws.UsedRange.Value2
    except Exception:
        return False
    if values is None:
        return False
    if not isinstance(values, tuple):
        return values not in ("", None)
    for row in values:
        if not isinstance(row, tuple):
            if row not in ("", None):
                return True
            continue
        for value in row:
            if value not in ("", None):
                return True
    return False


def _copy_full_sheet_into_output_com(wb_out, src_sheet, sheet_title: str) -> tuple[bool, str]:
    before_count = int(wb_out.Worksheets.Count)
    try:
        src_sheet.Copy(After=wb_out.Worksheets(before_count))
        if int(wb_out.Worksheets.Count) > before_count:
            new_ws = wb_out.Worksheets(wb_out.Worksheets.Count)
            try:
                new_ws.Name = sheet_title
            except Exception:
                pass
            ur = new_ws.UsedRange
            ur.Value = ur.Value
            if _sheet_has_content_com(new_ws):
                return True, ""
            new_ws.Delete()
    except Exception:
        pass

    new_ws = None
    try:
        new_ws = wb_out.Worksheets.Add(After=wb_out.Worksheets(wb_out.Worksheets.Count))
        try:
            new_ws.Name = sheet_title
        except Exception:
            pass
        src_range = src_sheet.UsedRange
        src_range.Copy(new_ws.Range("A1"))
        first_row = int(src_range.Row)
        first_col = int(src_range.Column)
        for col in range(first_col, first_col + int(src_range.Columns.Count)):
            new_ws.Columns(col).ColumnWidth = src_sheet.Columns(col).ColumnWidth
        for row in range(first_row, first_row + int(src_range.Rows.Count)):
            new_ws.Rows(row).RowHeight = src_sheet.Rows(row).RowHeight
        try:
            ur = new_ws.UsedRange
            ur.Value = ur.Value
        except Exception:
            pass
        if _sheet_has_content_com(new_ws):
            return True, ""
        try:
            new_ws.Delete()
        except Exception:
            pass
        return False, "UsedRange复制后为空"
    except Exception as e:
        if new_ws is not None:
            try:
                new_ws.Delete()
            except Exception:
                pass
        return False, str(e)


def run_extract_sheets_com(target_wbs: Iterable[Path], tasks: list[dict],
                           output_dir: Path, log_dir: Path,
                           append_existing: bool = False,
                           default_output_name: str | None = None) -> dict:
    logger = get_logger("feature_1_6_extract_sheets_com", log_dir)
    if not tasks:
        logger.warning("'工作表提取'配置无启用任务")
        return {"files": 0, "extracted_sheets": 0, "output_files": 0,
                "failed_files": 0, "paths": []}

    ts_default = default_output_name or f"提取结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    output_dir.mkdir(parents=True, exist_ok=True)

    out_wbs: dict[Path, object] = {}        # path -> COM Workbook (in memory)
    out_default_sheet: dict[Path, str] = {}  # path -> default sheet to drop later
    out_loaded_existing: dict[Path, bool] = {}
    stat = {"files": 0, "extracted_sheets": 0, "output_files": 0,
            "failed_files": 0, "paths": []}

    app = start_excel_app()
    try:
        for src in target_wbs:
            src_wb = None
            try:
                src_wb = open_readonly(app, src)
            except Exception as e:
                stat["failed_files"] += 1
                logger.error(f"打开失败: {src} - {e}")
                continue

            stat["files"] += 1
            logger.info(f"[COM] 打开源文件: {src}")
            try:
                for task in tasks:
                    if not _workbook_matches(src, task):
                        continue
                    if _name_matches_exclude(src.name, task):
                        logger.info(f"  [skip] {src.name} 命中排除关键字")
                        continue
                    matched = _find_matched_sheets_com(src_wb, task)
                    if not matched:
                        continue

                    path = _output_path_for_task(task, output_dir, ts_default)
                    if path not in out_wbs:
                        if append_existing and path.exists():
                            wb_existing = app.Workbooks.Open(str(path.resolve()), ReadOnly=False)
                            out_wbs[path] = wb_existing
                            out_loaded_existing[path] = True
                        else:
                            wb_new = _add_clean_workbook(app)
                            out_wbs[path] = wb_new
                            out_default_sheet[path] = str(wb_new.Worksheets(1).Name)
                            out_loaded_existing[path] = False
                    wb_out = out_wbs[path]

                    for src_sheet in matched:
                        src_title = str(src_sheet.Name)
                        sheet_title = _unique_sheet_name_com(wb_out, src_title)
                        if _is_full_sheet_task(task):
                            ok, err = _copy_full_sheet_into_output_com(wb_out, src_sheet, sheet_title)
                            if ok:
                                stat["extracted_sheets"] += 1
                                logger.info(
                                    f"  整表提取: {src.name} | {src_title} -> "
                                    f"{path.name}[{sheet_title}]"
                                )
                                continue
                            logger.warning(f"整表复制失败，回退值提取: {src.name} | {src_title} - {err}")
                            data = _read_full_2d(src_sheet)
                            values = _compact(data)
                            if not values:
                                logger.info(f"  [skip] {src.name} | {src_title} 无数据")
                                continue
                            new_ws = wb_out.Worksheets.Add(
                                After=wb_out.Worksheets(wb_out.Worksheets.Count)
                            )
                            try:
                                new_ws.Name = sheet_title
                            except Exception:
                                pass
                            _write_values_com(new_ws, values)
                            stat["extracted_sheets"] += 1
                            logger.info(
                                f"  整表值提取: {src.name} | {src_title} -> "
                                f"{path.name}[{sheet_title}]"
                            )
                        else:
                            # 指定区域/多区域提取，按值视图
                            data = _read_full_2d(src_sheet)
                            values = _extract_configured_values(
                                data, task.get("area_spec", ""), task.get("stitch_mode", ""),
                            )
                            if values is None:
                                logger.warning(
                                    f"  [skip] 配置行{task.get('row_no')} 提取区域无效: "
                                    f"{task.get('area_spec')}"
                                )
                                continue
                            values = _compact(values)
                            if not values:
                                logger.info(f"  [skip] {src.name} | {src_title} 无数据")
                                continue
                            new_ws = wb_out.Worksheets.Add(
                                After=wb_out.Worksheets(wb_out.Worksheets.Count)
                            )
                            try:
                                new_ws.Name = sheet_title
                            except Exception:
                                pass
                            _write_values_com(new_ws, values)
                            stat["extracted_sheets"] += 1
                            logger.info(
                                f"  部分提取: {src.name} | {src_title} -> "
                                f"{path.name}[{sheet_title}]"
                            )
            finally:
                safe_close(src_wb)

        # 保存所有输出工作簿，先删默认空白 sheet
        for path, wb_out in out_wbs.items():
            try:
                default = out_default_sheet.get(path, "")
                if default:
                    for ws in wb_out.Worksheets:
                        if str(ws.Name) == default and wb_out.Worksheets.Count > 1:
                            ws.Delete()
                            break
                path.parent.mkdir(parents=True, exist_ok=True)
                if out_loaded_existing.get(path):
                    wb_out.Save()
                else:
                    wb_out.SaveAs(str(path.resolve()), FileFormat=XL_OPEN_XML_WORKBOOK)
                stat["output_files"] += 1
                stat["paths"].append(str(path))
                logger.info(f"已保存: {path}")
            except Exception as e:
                logger.warning(f"保存输出失败 {path}: {e}")
            finally:
                safe_close(wb_out)
    finally:
        safe_quit(app)

    return stat
