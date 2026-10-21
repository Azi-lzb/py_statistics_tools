from __future__ import annotations
import sys
import os
import time
from datetime import datetime
from pathlib import Path
if __package__ in (None, ''):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from nocompytools.core.config_xlsx import validate_sheets_for_feature, load_global, load_timeline_rules, load_path_maps, load_dedup_tasks, load_print_tasks, load_global_value, load_institution_mapping, load_foreign_banks, load_extract_tasks
from nocompytools.core.precheck_399 import run_precheck
from nocompytools.core.compare_393 import run_compare
from nocompytools.core.timeline_396 import run_timeline_slim
from nocompytools.core.wide_397_398 import run_wide_summary
from nocompytools.core.dedup_311 import DedupTask, run_check_by_comment, run_delete_by_comment, run_check_by_config, run_delete_by_config, run_append_by_config, run_precheck_by_config
from nocompytools.core.config_init import initialize_or_repair_config
from nocompytools.core.print_310 import run_print_keep_by_comment, run_print_fast_by_comment, run_print_config_all_modes, run_print_config_precheck
from nocompytools.core.summary_321_322 import run_summary_by_usedrange, run_summary_by_comment
from nocompytools.core.convert_33_34_35_37 import run_batch_rename_33, run_batch_rename_sheet_37, run_batch_rename_by_content_46, run_batch_unrename_by_content_47, run_batch_rename_local_map_48, run_submission_check_48, run_file_sort_prefix_49, run_remove_file_sort_prefix_410, run_archive_by_date_preview, run_archive_by_type_preview, run_archive_plan_copy, run_archive_by_date_copy, run_archive_by_type_copy
from nocompytools.core.feature_1_1_split_village_bank import run_split_village_bank
from nocompytools.core.feature_1_2_normalize_institution import run_normalize_institution
from nocompytools.core.feature_1_3_remove_foreign_bank import run_remove_foreign_bank
from nocompytools.core.feature_1_4_fx_header import run_fx_header_fix
from nocompytools.core.feature_1_5_region_sum_check import run_region_sum_check
from nocompytools.core.feature_1_6_extract_sheets import run_extract_sheets
from nocompytools.core.feature_1_8_adjust_rural_loan import run_adjust_rural_loan
from nocompytools.core.feature_1_9_adjust_rural_loan_quarterly import run_adjust_rural_loan_quarterly
from nocompytools.core.app_identity import app_title
from nocompytools.core.xls_to_xlsx import run_xls_to_xlsx
from nocompytools.core.word_convert import run_word_convert

def _app_base_dir() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent
CFG_PATH = Path(os.environ.get('PYTOOLS_CONFIG_PATH', _app_base_dir() / 'config' / 'config.xlsx'))
try:
    _APP_TITLE = app_title()
except Exception:
    _APP_TITLE = 'pytools'
MAIN_MENU = f'\n======== {_APP_TITLE} 菜单 ========\n  1  时序工具\n  2  去重工具\n  3  打印工具\n  4  转换工具\n  5  配置工具\n  6  汇总工具\n  7  一二批处理工具\n  8  归档工具\n  9  快捷工具\n  0  退出\n==============================\n'
ARCHIVE_MENU = '\n------ 归档工具 ------\n  1  按日期归档（预览）\n  2  按类型归档（预览，归档类型配置）\n  3  执行归档计划（归档计划）\n  4  按日期归档（直接复制）\n  5  按类型归档（直接复制，归档类型配置）\n  0  返回\n----------------------\n'
PENDING_MENU = '\n------ 一二批处理工具 ------\n  1  拆分村镇银行数据\n  2  机构名称标准化\n  3  删除分机构表的外资行\n  4  外汇页眉修改\n  5  地区总分校验\n  6  提取工作表数据（按值提取，不复制格式）\n  7  涉农贷款比上月修正\n  8  季度涉农贷款比上月修正及分地区核对\n  0  返回\n'
TIMELINE_MENU = '\n------ 时序工具 ------\n  1  表头路径比对\n  2  时序提取\n  3  宽表汇总（行头加后缀）\n  4  宽表汇总（行头不加后缀）\n  5  预校验\n  0  返回\n'
DEDUP_MENU = '\n------ 去重工具 ------\n  1  按批注检查重复\n  2  按批注删除重复\n  3  按配置检查重复（去重追加数据配置）\n  4  按配置删除重复（去重追加数据配置）\n  5  按配置去重追加到目标（去重追加数据配置）\n  6  按配置预校验（去重追加数据配置）\n  0  返回\n----------------------\n'
CONFIG_MENU = '\n------ 配置工具 ------\n  1  初始化/修复配置表\n  2  查看当前配置（重新读取磁盘）\n  0  返回\n----------------------\n'
CONVERT_MENU = '\n------ 转换工具 ------\n  1  批量重命名文件\n  2  旧版 Excel 转换（xls → xlsx，按系统选择引擎）\n  3  批量修改 Word 格式（doc/docx，按系统选择引擎）\n  4  批量修改 Sheet 名\n  5  根据文件内容重命名\n  6  取消根据内容重命名前缀\n  7  批量局部映射重命名文件\n  8  提交缺失检查\n  9  文件排序前缀\n 10  取消文件排序前缀\n  0  返回\n'
SUMMARY_MENU = '\n------ 汇总工具 ------\n  1  按使用区域汇总（全量）\n  2  按批注汇总（模板+源文件）\n  0  返回\n'
QUICK_MENU = '\n------ 快捷工具 ------\n  1  打开配置文件（config.xlsx）\n  2  打开 output 最新文件\n  3  打开 output 文件夹\n  0  返回\n----------------------\n'
PRINT_MENU = '\n------ 打印工具 ------\n  1  按批注打印（保留源格式）\n  2  按批注打印（快速复制，Calamine 读取，保留框线和合并）\n  3  按配置打印（打印配置）\n  4  按配置打印预校验（打印配置）\n  0  返回\n'

def _print_selected_paths(paths: list[Path], label: str='已选择文件') -> None:
    if not paths:
        return
    print(f'→ {label} {len(paths)} 个：')
    for idx, path in enumerate(paths, start=1):
        print(f'  {idx}. {path}')

def _pick_files_for_compare() -> tuple[Path | None, list[Path]]:
    """弹文件选择框：先模板，再多选源文件。"""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        print('→ 请在弹出窗口中选择 [模板文件]')
        tmpl = filedialog.askopenfilename(title='选择模板文件', filetypes=[('Excel', '*.xlsx *.xlsm'), ('所有', '*.*')])
        if not tmpl:
            root.destroy()
            return (None, [])
        _print_selected_paths([Path(tmpl)], '已选择模板文件')
        print('→ 请在弹出窗口中选择 [源文件]（可多选）')
        srcs = filedialog.askopenfilenames(title='选择源文件（可多选）', filetypes=[('Excel', '*.xlsx *.xlsm'), ('所有', '*.*')])
        root.destroy()
        src_paths = [Path(p) for p in srcs]
        _print_selected_paths(src_paths, '已选择源文件')
        return (Path(tmpl), src_paths)
    except Exception as e:
        print(f'[提示] 文件弹窗不可用：{e}')
    tmpl_raw = input('模板文件路径（输入 0 取消）> ').strip().strip('"').strip("'")
    if not tmpl_raw or tmpl_raw == '0':
        return (None, [])
    src_raw = input('源文件路径（可用分号分隔多个，输入 0 取消）> ').strip()
    if not src_raw or src_raw == '0':
        return (None, [])
    tmpl_path = Path(tmpl_raw)
    src_paths = [Path(item.strip().strip('"').strip("'")) for item in src_raw.replace('；', ';').split(';') if item.strip() and Path(item.strip().strip('"').strip("'")).exists()]
    _print_selected_paths([tmpl_path], '已输入模板文件')
    _print_selected_paths(src_paths, '已输入源文件')
    return (tmpl_path if tmpl_path.exists() else None, src_paths)

def _pick_source_files(title: str='选择源文件（可多选）', filetypes: list[tuple[str, str]] | None=None) -> list[Path]:
    print(f'→ 正在打开文件选择窗口：{title}')
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        root.lift()
        try:
            root.focus_force()
        except Exception:
            pass
        root.update()
        srcs = filedialog.askopenfilenames(title=title, filetypes=filetypes or [('Excel/CSV', '*.xlsx *.xlsm *.csv'), ('所有', '*.*')])
        root.destroy()
        picked = [Path(p) for p in srcs]
        if picked:
            _print_selected_paths(picked, '已选择文件')
            return picked
    except Exception as e:
        print(f'[提示] 文件弹窗不可用：{e}')
    print('→ 未通过弹窗选到文件，可手动输入路径（分号分隔），或输入 0 取消：')
    raw = input('文件路径> ').strip()
    if raw in ('', '0'):
        return []
    parts = [x.strip().strip('"').strip("'") for x in raw.replace('；', ';').split(';') if x.strip()]
    out: list[Path] = []
    for p in parts:
        pp = Path(p)
        if pp.exists():
            out.append(pp)
    _print_selected_paths(out, '已输入文件')
    return out

def _pick_source_folders(title: str) -> list[Path]:
    """连续选择多个源文件夹，弹窗不可用时允许分号输入。"""
    print(f'→ 正在打开文件夹选择窗口：{title}')
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        root.lift()
        root.update()
        folders: list[Path] = []
        while True:
            raw = filedialog.askdirectory(title=title)
            if not raw:
                break
            folder = Path(raw)
            if folder not in folders:
                folders.append(folder)
            if not messagebox.askyesno('继续选择', '是否继续选择其他源文件夹？', parent=root):
                break
        root.destroy()
        if folders:
            _print_selected_paths(folders, '已选择文件夹')
            return folders
    except Exception as e:
        print(f'[提示] 文件夹弹窗不可用：{e}')
    raw = input('文件夹路径（可用分号分隔多个，输入 0 取消）> ').strip()
    if raw in ('', '0'):
        return []
    folders: list[Path] = []
    for text in raw.replace('；', ';').split(';'):
        folder = Path(text.strip().strip('"').strip("'"))
        if not text.strip():
            continue
        if not folder.is_dir():
            print(f'[提示] 文件夹不存在，已跳过: {folder}')
            continue
        if folder not in folders:
            folders.append(folder)
    _print_selected_paths(folders, '已输入文件夹')
    return folders

def _pick_excel_folders(title: str) -> list[Path]:
    """兼容旧调用：连续选择多个 Excel 源文件夹。"""
    return _pick_source_folders(title)

def _collect_files(folder: Path, recursive: bool, suffixes: set[str] | None=None) -> list[Path]:
    """收集文件夹中的文件；suffixes 为 None 时不限制文件类型。"""
    normalized_suffixes = {suffix.lower() for suffix in suffixes} if suffixes else None
    candidates = folder.rglob('*') if recursive else folder.iterdir()
    return sorted((path for path in candidates if path.is_file() and (not path.name.startswith('~$')) and (normalized_suffixes is None or path.suffix.lower() in normalized_suffixes)), key=lambda path: str(path.relative_to(folder)).casefold())

def _collect_files_from_folders(folders: list[Path], recursive: bool, suffixes: set[str] | None=None) -> list[Path]:
    """合并多个文件夹中的文件，并按绝对路径去重。"""
    seen: dict[str, Path] = {}
    for folder in folders:
        for path in _collect_files(folder, recursive, suffixes):
            resolved = path.resolve()
            seen[str(resolved).casefold()] = resolved
    return sorted(seen.values(), key=lambda path: str(path).casefold())

def _collect_excel_files(folder: Path, recursive: bool) -> list[Path]:
    """兼容旧调用：收集文件夹中的 Excel 文件。"""
    return _collect_files(folder, recursive, {'.xlsx', '.xlsm'})

def _collect_excel_files_from_folders(folders: list[Path], recursive: bool) -> list[Path]:
    """兼容旧调用：合并多个文件夹的 Excel 文件。"""
    return _collect_files_from_folders(folders, recursive, {'.xlsx', '.xlsm'})

def _pick_source_files_or_folders(title: str, filetypes: list[tuple[str, str]] | None=None, suffixes: set[str] | None=None) -> list[Path]:
    """选择多个文件，或选择多个文件夹后按当前层/递归收集文件。"""
    print(f'→ {title}')
    print('  1  选择文件（可多选）')
    print('  2  选择文件夹（可多选，仅当前层）')
    print('  3  选择文件夹（可多选，递归全部子目录）')
    print('  0  取消')
    mode = input('选择方式> ').strip()
    if mode == '1':
        return _pick_source_files(title, filetypes=filetypes)
    if mode not in {'2', '3'}:
        return []
    folders = _pick_source_folders(title)
    if not folders:
        return []
    recursive = mode == '3'
    sources = _collect_files_from_folders(folders, recursive, suffixes)
    scope = '含子文件夹' if recursive else '当前层'
    print(f'→ 已选 {len(folders)} 个文件夹，{scope}共识别文件 {len(sources)} 个')
    _print_selected_paths(sources, '参与处理的文件')
    return sources

def _pick_excel_sources_from_folder(title: str, recursive: bool) -> list[Path]:
    folders = _pick_excel_folders(title)
    if not folders:
        return []
    sources = _collect_excel_files_from_folders(folders, recursive)
    scope = '含子文件夹' if recursive else '当前层'
    print(f'→ 已选 {len(folders)} 个文件夹，{scope}共识别 Excel 文件 {len(sources)} 个')
    _print_selected_paths(sources, '参与宽表汇总的文件')
    return sources

def _pick_archive_paths(title: str) -> list[Path]:
    print(f'→ {title}')
    print('  1  选择文件（可多选）')
    print('  2  选择文件夹（递归遍历）')
    print('  0  取消')
    mode = input('选择方式> ').strip()
    if mode == '0':
        return []
    if mode == '1':
        return _pick_source_files(title, filetypes=[('所有文件', '*.*')])
    if mode == '2':
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            root.attributes('-topmost', True)
            root.lift()
            root.update()
            p = filedialog.askdirectory(title=title)
            root.destroy()
            paths = [Path(p)] if p else []
            _print_selected_paths(paths, '已选择文件夹')
            return paths
        except Exception as e:
            print(f'[提示] 文件夹弹窗不可用：{e}')
    raw = input('文件夹路径（输入 0 取消）> ').strip().strip('"').strip("'")
    if raw in ('', '0'):
        return []
    folder = Path(raw)
    if not folder.is_dir():
        print(f'[提示] 文件夹不存在: {folder}')
        return []
    _print_selected_paths([folder], '已输入文件夹')
    return [folder]

def _pick_archive_target_root() -> Path | None:
    print('→ 请选择归档目标根目录')
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        root.lift()
        root.update()
        p = filedialog.askdirectory(title='选择归档目标根目录')
        root.destroy()
        if p:
            picked = Path(p)
            _print_selected_paths([picked], '已选择目标目录')
            return picked
    except Exception as e:
        print(f'[提示] 目录弹窗不可用：{e}')
    raw = input('目标根目录路径（输入 0 取消）> ').strip()
    if raw in ('', '0'):
        return None
    picked = Path(raw.strip('"').strip("'"))
    _print_selected_paths([picked], '已输入目标目录')
    return picked

def _open_path(path: Path) -> None:
    if not path.exists():
        raise FileNotFoundError(str(path))
    import subprocess
    if sys.platform == 'win32':
        os.startfile(str(path))
    else:
        subprocess.Popen(['xdg-open', str(path)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

def _latest_file_in_dir(dir_path: Path) -> Path | None:
    if not dir_path.exists():
        return None
    files = [p for p in dir_path.rglob('*') if p.is_file() and (not p.name.startswith('~$'))]
    if not files:
        return None
    return max(files, key=lambda p: p.stat().st_mtime)

def _show_main_menu() -> str:
    print(MAIN_MENU)
    while True:
        try:
            s = input('主菜单选择> ').strip()
        except EOFError:
            return '0'
        if s in ('1', '2', '3', '4', '5', '6', '7', '8', '9', '0'):
            return s
        print('仅接受 1/2/3/4/5/6/7/8/9/0')

def _show_sub_menu(title: str, valid: tuple[str, ...]) -> str:
    print(title)
    while True:
        try:
            s = input('子菜单选择> ').strip()
        except EOFError:
            return '0'
        if s in valid:
            return s
        print('输入无效，请重试。')

def _pick_one_workbook(title: str) -> Path | None:
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        p = filedialog.askopenfilename(title=title, filetypes=[('Excel', '*.xlsx *.xlsm'), ('所有', '*.*')])
        root.destroy()
    except Exception as e:
        print(f'[提示] 文件弹窗不可用：{e}')
        p = input('Excel 文件路径（输入 0 取消）> ').strip().strip('"').strip("'")
    if not p or p == '0':
        return None
    picked = Path(p)
    if not picked.exists():
        print(f'[提示] 文件不存在: {picked}')
        return None
    _print_selected_paths([picked], '已选择文件')
    return picked

def _to_dedup_tasks(raw_tasks: list[dict]) -> list[DedupTask]:
    tasks: list[DedupTask] = []
    for t in raw_tasks:
        if t.get('source_wb') is None:
            continue
        tasks.append(DedupTask(row_no=int(t.get('row_no') or 0), enabled=bool(t.get('enabled')), source_wb=t['source_wb'], source_ws=str(t.get('source_ws') or ''), key_cols=list(t.get('key_cols') or []), target_wb=t.get('target_wb'), target_ws=t.get('target_ws'), write_mode=str(t.get('write_mode') or ''), task_name=str(t.get('task_name') or '')))
    return tasks

def _view_config() -> None:
    import pandas as pd
    print(f'配置文件: {CFG_PATH}')
    if not CFG_PATH.exists():
        print('  [缺失] 请先 5-1 初始化/修复配置表')
        return
    try:
        mtime = CFG_PATH.stat().st_mtime
        from datetime import datetime
        print(f"最后修改: {datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')}")
    except Exception:
        pass
    try:
        xls = pd.ExcelFile(CFG_PATH)
    except Exception as e:
        print(f'  [读取失败] {e}')
        return
    print(f'Sheet 列表（{len(xls.sheet_names)}）:')
    for s in xls.sheet_names:
        try:
            df = pd.read_excel(xls, s, header=0, dtype=object)
            print(f'  - {s}: {len(df)} 行 × {df.shape[1]} 列')
        except Exception as e:
            print(f'  - {s}: 读取失败 {e}')
    try:
        g = load_global(CFG_PATH)
        print('全局配置解析:')
        print(f'  输入目录: {g.input_dir}')
        print(f'  输出目录: {g.output_dir}')
        print(f'  日志目录: {g.log_dir}')
        print(f'  源文件扩展名: {g.source_exts}')
    except Exception as e:
        print(f'  [全局配置解析失败] {e}')
    try:
        rules = load_timeline_rules(CFG_PATH)
        print(f'时序提取规则（启用）: {len(rules)} 条')
        for r in rules:
            print(f'  · {r.name}  wb={r.wb_keyword}  sheet={r.sheet_keyword}')
    except Exception:
        pass
    try:
        mapping = load_institution_mapping(CFG_PATH)
        foreign = load_foreign_banks(CFG_PATH)
        print(f'机构映射: {len(mapping)} 条；外资行: {len(foreign)} 个')
    except Exception:
        pass
    try:
        tasks = load_extract_tasks(CFG_PATH)
        print(f'工作表提取任务（启用）: {len(tasks)} 条')
    except Exception:
        pass

def _check(feature: str) -> bool:
    errs = validate_sheets_for_feature(feature, CFG_PATH)
    if errs:
        print('[配置缺失，已退出本功能]')
        for e in errs:
            print('  - ' + e)
        return False
    return True

def dispatch(choice: str) -> None:
    if choice == 'c':
        stat = initialize_or_repair_config(CFG_PATH)
        print(f"[完成] 配置修复: 新建Sheet={stat.get('created_sheets', 0)} 修复Sheet={stat.get('repaired_sheets', 0)} 文件={CFG_PATH}")
        for line in stat.get('details', []) or []:
            print(f'  - {line}')
        return
    if choice == 'cv':
        _view_config()
        return
    if choice == 'q1':
        try:
            _open_path(CFG_PATH)
            print(f'[完成] 已打开配置文件: {CFG_PATH}')
        except Exception as e:
            print(f'[失败] 打开配置文件失败: {e}')
        return
    if choice == 'q2':
        g = load_global(CFG_PATH)
        latest = _latest_file_in_dir(g.output_dir)
        if latest is None:
            print(f'[提示] 未找到输出文件: {g.output_dir}')
            return
        try:
            _open_path(latest)
            print(f'[完成] 已打开最新输出文件: {latest}')
        except Exception as e:
            print(f'[失败] 打开输出文件失败: {e}')
        return
    if choice == 'q3':
        g = load_global(CFG_PATH)
        try:
            g.output_dir.mkdir(parents=True, exist_ok=True)
            _open_path(g.output_dir)
            print(f'[完成] 已打开输出文件夹: {g.output_dir}')
        except Exception as e:
            print(f'[失败] 打开输出文件夹失败: {e}')
        return
    if choice == '5':
        run_precheck(CFG_PATH)
        return
    if not _check(choice):
        return
    g = load_global(CFG_PATH)
    g.output_dir.mkdir(parents=True, exist_ok=True)
    g.log_dir.mkdir(parents=True, exist_ok=True)
    hide_zero_global = str(load_global_value(CFG_PATH, '打印零值不输出', '') or load_global_value(CFG_PATH, '打印零值不显示', '')).strip().lower() in ('是', '1', 'true', 'y', 'yes', 'on')
    if choice == '1':
        rules = load_timeline_rules(CFG_PATH)
        if not rules:
            print('[无启用的时序提取规则]')
            return
        tmpl, srcs = _pick_files_for_compare()
        if not tmpl or not srcs:
            print('[已取消]')
            return
        out = run_compare(tmpl, srcs, rules, g)
        print(f'[完成] 输出: {out}')
    elif choice == '2':
        rules = load_timeline_rules(CFG_PATH)
        path_maps = load_path_maps(CFG_PATH)
        if not rules:
            print('[无启用的时序提取规则]')
            return
        srcs = _pick_source_files('选择时序提取源文件（可多选）')
        if not srcs:
            print('[已取消]')
            return
        out = run_timeline_slim(rules, path_maps, g, source_paths=srcs)
        if out is None:
            print('[完成] 无匹配结果，未生成文件')
        else:
            print(f'[完成] 输出: {out}')
    elif choice in ('3', '4'):
        rules = load_timeline_rules(CFG_PATH)
        path_maps = load_path_maps(CFG_PATH)
        if not rules:
            print('[无启用的时序提取规则]')
            return
        srcs = _pick_source_files('选择宽表汇总源文件（可多选）')
        if not srcs:
            print('[已取消]')
            return
        suffix = choice == '3'
        t0 = time.perf_counter()
        out = run_wide_summary(rules, path_maps, g, row_suffix_enabled=suffix, source_paths=srcs)
        elapsed = time.perf_counter() - t0
        if out is None:
            print(f'[完成] 无匹配结果，未生成文件（耗时 {elapsed:.2f}s）')
        else:
            print(f'[完成] 输出: {out}（耗时 {elapsed:.2f}s）')
    elif choice == '6':
        picked = _pick_one_workbook('选择要按批注检查重复的工作簿')
        if not picked:
            print('[已取消]')
            return
        stat = run_check_by_comment(picked, g.log_dir)
        print(f"[完成] 命中Sheet={stat['hit_sheets']} 标红行={stat['marked_rows']}")
    elif choice == '7':
        picked = _pick_one_workbook('选择要按批注删除重复的工作簿')
        if not picked:
            print('[已取消]')
            return
        stat = run_delete_by_comment(picked, g.log_dir)
        print(f"[完成] 命中Sheet={stat['hit_sheets']} 删除行={stat['deleted_rows']}")
    elif choice == '8':
        tasks = _to_dedup_tasks(load_dedup_tasks(CFG_PATH))
        stat = run_check_by_config(tasks, g.log_dir)
        print(f"[完成] 任务成功={stat['task_ok']} 跳过={stat['task_skip']} 标红行={stat['marked_rows']} 耗时={float(stat.get('elapsed_sec', 0.0)):.2f}s")
    elif choice == '9':
        tasks = _to_dedup_tasks(load_dedup_tasks(CFG_PATH))
        stat = run_delete_by_config(tasks, g.log_dir)
        print(f"[完成] 任务成功={stat['task_ok']} 跳过={stat['task_skip']} 删除行={stat['deleted_rows']} 耗时={float(stat.get('elapsed_sec', 0.0)):.2f}s")
    elif choice == 'a':
        tasks = _to_dedup_tasks(load_dedup_tasks(CFG_PATH))
        stat = run_append_by_config(tasks, g.log_dir)
        print(f"[完成] 任务成功={stat['task_ok']} 跳过={stat['task_skip']} 追加={stat['appended_rows']} 去重删除={stat['deleted_rows']} 耗时={float(stat.get('elapsed_sec', 0.0)):.2f}s")
    elif choice == 'b':
        tasks = _to_dedup_tasks(load_dedup_tasks(CFG_PATH))
        stat = run_precheck_by_config(tasks, g.log_dir)
        print(f"[完成] 预校验 OK={stat['ok']} WARN={stat['warn']} FAIL={stat['fail']}")
    elif choice == 'p1':
        srcs = _pick_source_files('选择按批注打印（保留源格式）源文件（可多选）')
        if not srcs:
            print('[已取消]')
            return
        stat = run_print_keep_by_comment(srcs, g.output_dir, g.log_dir, hide_zero=hide_zero_global)
        print(f"[完成] 命中工作簿={stat['workbooks_hit']} 命中sheet={stat['sheets_hit']} 输出文件={stat['saved_files']} 跳过={stat['skipped']}")
    elif choice == 'p3':
        srcs = _pick_source_files('选择 3.10.3 源文件（可多选）')
        if not srcs:
            print('[已取消]')
            return
        stat = run_print_fast_by_comment(srcs, g.output_dir, g.log_dir, hide_zero=hide_zero_global)
        print(f"[完成] 命中工作簿={stat['workbooks_hit']} 命中sheet={stat['sheets_hit']} 输出文件={stat['saved_files']} 跳过={stat['skipped']}")
    elif choice == 'p7':
        tasks = load_print_tasks(CFG_PATH)
        stat = run_print_config_all_modes(tasks, g.log_dir)
        print(f"[完成] 任务成功={stat['task_ok']} 跳过={stat['task_skip']} 写入sheet={stat['written_sheets']} 写入行={stat['written_rows']}")
        total = int(stat.get('task_total', 0) or 0)
        elapsed = float(stat.get('elapsed_sec', 0.0) or 0.0)
        avg = elapsed / total if total > 0 else 0.0
        print(f'[耗时] 总耗时={elapsed:.2f}s 平均每任务={avg:.2f}s (任务总数={total})')
        slow_top = stat.get('slow_top', []) or []
        if slow_top:
            print('[耗时] 最慢任务TOP3:')
            for i, r in enumerate(slow_top, start=1):
                print(f"  {i}. row={r.get('row_no')} mode={r.get('mode')} target={r.get('target')} 耗时={float(r.get('elapsed_sec', 0.0)):.2f}s status={r.get('status')}")
    elif choice == 'p8':
        tasks = load_print_tasks(CFG_PATH)
        stat = run_print_config_precheck(tasks, g.log_dir)
        print(f"[完成] 预校验 OK={stat['ok']} WARN={stat['warn']} FAIL={stat['fail']}")
    elif choice == 't3':
        srcs = _pick_source_files_or_folders('选择要批量重命名的文件或文件夹', filetypes=[('所有文件', '*.*')])
        if not srcs:
            print('[已取消]')
            return
        stat = run_batch_rename_33(CFG_PATH, srcs, g.log_dir)
        print(f"[完成] 成功={stat['ok']} 跳过={stat['skip']}")
    elif choice == 't4':
        srcs = _pick_source_files_or_folders('选择旧版 .xls 文件或文件夹（转换为 .xlsx）', filetypes=[('Excel 97-2003', '*.xls')], suffixes={'.xls'})
        if not srcs:
            print('[已取消]')
            return
        try:
            stat = run_xls_to_xlsx(srcs, g.log_dir)
        except Exception as e:
            print(f"[转换失败] {e}")
            return
        if stat.get("error"):
            print(f"[转换引擎不可用] {stat['error']}")
        print(f"[完成] 目标格式={stat['target_ext']} 引擎={stat.get('engine', '')} 成功={stat['ok']} 跳过/失败={stat['skip']}")
    elif choice == 't5':
        srcs = _pick_source_files_or_folders('选择要批量转换格式的 Word 文件或文件夹', filetypes=[('Word', '*.doc *.docx')], suffixes={'.doc', '.docx'})
        if not srcs:
            print('[已取消]')
            return
        target_fmt = load_global_value(CFG_PATH, 'word目的格式', 'docx') or 'docx'
        stat = run_word_convert(srcs, target_fmt, g.log_dir)
        if stat.get("error"):
            print(f"[转换引擎不可用] {stat['error']}")
        print(f"[完成] 目标格式={stat['target_ext']} 引擎={stat.get('engine', '')} 成功={stat['ok']} 跳过/失败={stat['skip']}")
    elif choice == 't7':
        srcs = _pick_source_files_or_folders('选择要批量修改 Sheet 名的工作簿或文件夹', filetypes=[('Excel', '*.xlsx *.xlsm')], suffixes={'.xlsx', '.xlsm'})
        if not srcs:
            print('[已取消]')
            return
        stat = run_batch_rename_sheet_37(CFG_PATH, srcs, g.log_dir)
        print(f"[完成] 工作簿={stat['workbooks']} 重命名={stat['rename_ok']} 跳过={stat['skip']}")
    elif choice == 't8':
        srcs = _pick_source_files_or_folders('选择要根据内容重命名的 Excel 文件或文件夹', filetypes=[('Excel', '*.xlsx *.xlsm')], suffixes={'.xlsx', '.xlsm'})
        if not srcs:
            print('[已取消]')
            return
        stat = run_batch_rename_by_content_46(CFG_PATH, srcs, g.log_dir)
        remark = f" 备注={stat.get('config_remark', '')}" if stat.get('config_remark') else ''
        print(f"[完成] 配置行={stat.get('config_row', '')} 规则={stat.get('spec', '')}{remark} 成功={stat['ok']} 跳过={stat['skip']}")
    elif choice == 't9':
        srcs = _pick_source_files_or_folders('选择要取消根据内容重命名前缀的 Excel 文件或文件夹', filetypes=[('Excel', '*.xlsx *.xlsm')], suffixes={'.xlsx', '.xlsm'})
        if not srcs:
            print('[已取消]')
            return
        stat = run_batch_unrename_by_content_47(CFG_PATH, srcs, g.log_dir)
        remark = f" 备注={stat.get('config_remark', '')}" if stat.get('config_remark') else ''
        print(f"[完成] 配置行={stat.get('config_row', '')} 规则={stat.get('spec', '')}{remark} 成功={stat['ok']} 跳过={stat['skip']}")
    elif choice == 't10':
        srcs = _pick_source_files_or_folders('选择要按 M:N 局部映射重命名的文件或文件夹（支持任意文件）', filetypes=[('所有文件', '*.*')])
        if not srcs:
            print('[已取消]')
            return
        stat = run_batch_rename_local_map_48(CFG_PATH, srcs, g.log_dir)
        print(f"[完成] 映射规则={stat.get('rules', 0)} 成功={stat['ok']} 跳过={stat['skip']}")
    elif choice == 't11':
        srcs = _pick_source_files_or_folders('选择要检查提交情况的文件或文件夹', filetypes=[('所有文件', '*.*')])
        if not srcs:
            print('[已取消]')
            return
        stat = run_submission_check_48(CFG_PATH, srcs, g.output_dir, g.log_dir)
        print(f"[完成] 输出={stat['output']} 配置项={stat['rules']} 文件={stat['files']} 已提交={stat['submitted']} 缺失={stat['missing']} 未识别文件={stat['unmatched']} 多重命中={stat['multi']}")
    elif choice == 't12':
        srcs = _pick_source_files_or_folders('选择要添加排序前缀的文件或文件夹', filetypes=[('所有文件', '*.*')])
        if not srcs:
            print('[已取消]')
            return
        stat = run_file_sort_prefix_49(CFG_PATH, srcs, g.output_dir, g.log_dir)
        print(f"[完成] 输出={stat['output']} 规则={stat['rules']} 文件={stat['files']} 重命名={stat['renamed']} 未命中={stat['unmatched']} 跳过={stat['skipped']} 失败={stat['failed']}")
    elif choice == 't13':
        srcs = _pick_source_files_or_folders('选择要取消排序前缀的文件或文件夹', filetypes=[('所有文件', '*.*')])
        if not srcs:
            print('[已取消]')
            return
        stat = run_remove_file_sort_prefix_410(CFG_PATH, srcs, g.output_dir, g.log_dir)
        print(f"[完成] 输出={stat['output']} 规则={stat['rules']} 文件={stat['files']} 已取消前缀={stat['renamed']} 跳过={stat['skipped']} 失败={stat['failed']}")
    elif choice == 'ar_date':
        paths = _pick_archive_paths('选择要按日期归档预览的文件或文件夹')
        if not paths:
            print('[已取消]')
            return
        target_root = _pick_archive_target_root()
        if target_root is None:
            print('[已取消]')
            return
        stat = run_archive_by_date_preview(paths, target_root, g.output_dir, g.log_dir, CFG_PATH)
        print(f"[完成] 归档计划={stat['output']} 文件={stat['files']} 后缀跳过={stat.get('skipped_ext', 0)} 无效路径={stat['invalid']} 目标根目录={stat['target_root']}")
    elif choice == 'ar_type':
        paths = _pick_archive_paths('选择要按类型归档预览的文件或文件夹')
        if not paths:
            print('[已取消]')
            return
        target_root = _pick_archive_target_root()
        if target_root is None:
            print('[已取消]')
            return
        stat = run_archive_by_type_preview(CFG_PATH, paths, target_root, g.output_dir, g.log_dir)
        print(f"[完成] 归档计划={stat['output']} 文件={stat['files']} 规则={stat['rules']} 后缀跳过={stat.get('skipped_ext', 0)} 无效路径={stat['invalid']} 目标根目录={stat['target_root']}")
    elif choice == 'ar_exec':
        plan = _pick_one_workbook('选择归档计划 Excel 文件')
        if plan is None:
            print('[已取消]')
            return
        stat = run_archive_plan_copy(plan, g.output_dir, g.log_dir)
        print(f"[完成] 执行结果={stat['output']} 已复制={stat['copied']} 跳过={stat['skip']} 失败={stat['failed']}")
    elif choice == 'ar_date_copy':
        paths = _pick_archive_paths('选择要按日期直接复制归档的文件或文件夹')
        if not paths:
            print('[已取消]')
            return
        target_root = _pick_archive_target_root()
        if target_root is None:
            print('[已取消]')
            return
        stat = run_archive_by_date_copy(paths, target_root, g.output_dir, g.log_dir, CFG_PATH)
        print(f"[完成] 执行结果={stat['output']} 已复制={stat['copied']} 跳过={stat['skip']} 后缀跳过={stat.get('skipped_ext', 0)} 失败={stat['failed']} 目标根目录={stat['target_root']}")
    elif choice == 'ar_type_copy':
        paths = _pick_archive_paths('选择要按类型直接复制归档的文件或文件夹')
        if not paths:
            print('[已取消]')
            return
        target_root = _pick_archive_target_root()
        if target_root is None:
            print('[已取消]')
            return
        stat = run_archive_by_type_copy(CFG_PATH, paths, target_root, g.output_dir, g.log_dir)
        print(f"[完成] 执行结果={stat['output']} 已复制={stat['copied']} 规则={stat['rules']} 跳过={stat['skip']} 后缀跳过={stat.get('skipped_ext', 0)} 失败={stat['failed']} 目标根目录={stat['target_root']}")
    elif choice == 's21':
        srcs = _pick_source_files('选择按使用区域汇总的源文件（可多选）', filetypes=[('Excel', '*.xlsx *.xlsm *.xls *.xlsb'), ('所有', '*.*')])
        if not srcs:
            print('[已取消]')
            return
        stat = run_summary_by_usedrange(srcs, g.output_dir, g.log_dir)
        if not stat.get('saved'):
            print('[完成] 无可汇总数据，未生成文件')
        else:
            print(f"[完成] 行数={stat['rows']} 输出={stat['path']}")
    elif choice == 'x11':
        print('→ 第一次文件选择：选择 本期商业银行共享一批')
        curr = _pick_one_workbook('选择 本期商业银行共享一批')
        if not curr:
            print('[已取消]')
            return
        print('→ 第二次文件选择：选择 去年同期商业银行共享一批')
        prev = _pick_one_workbook('选择 去年同期商业银行共享一批')
        if not prev:
            print('[已取消]')
            return
        print('→ 第三次文件选择：选择 本外币、人民币、外币、外发用文件（修改分机构表）')
        srcs = _pick_source_files('选择 本外币、人民币、外币、外发用文件（修改分机构表）（可多选）')
        if not srcs:
            print('[已取消]')
            return
        normal_srcs = srcs
        if normal_srcs:
            stat = run_split_village_bank(curr, prev, normal_srcs, g.log_dir)
            if stat.get('matrix_empty'):
                print('[已中止] Python 村镇银行计算结果为空')
            else:
                print(f"[完成] Python 文件={stat['files']} 分机构sheet={stat['branch_sheets']} 填入={stat['filled_sheets']} 跳过={stat['skipped_sheets']} 失败={stat['failed_files']}")
    elif choice == 'x12':
        mapping = load_institution_mapping(CFG_PATH)
        if not mapping:
            print("[已中止] 机构映射表为空，请在 config.xlsx 的'机构映射表'中填写 A=原始 B=映射")
            return
        srcs = _pick_source_files("选择包含'分机构'工作表的目标文件（可多选）")
        if not srcs:
            print('[已取消]')
            return
        normal_srcs = srcs
        if normal_srcs:
            stat = run_normalize_institution(normal_srcs, mapping, g.log_dir)
            print(f"[完成] Python 文件={stat['files']} 分机构sheet={stat['branch_sheets']} 命中sheet={stat['mapped_sheets']} 修改行={stat['mapped_rows']} 失败={stat['failed_files']}")
    elif choice == 'x13':
        foreign_set = load_foreign_banks(CFG_PATH)
        if not foreign_set:
            print("[已中止] 外资行配置为空，请在 config.xlsx 的'机构映射表' C 列将外资行标记为 1")
            return
        print('→ 文件选择：选择 外发用报表（删除分机构表的外资行）')
        srcs = _pick_source_files('选择 外发用报表（删除分机构表的外资行）（可多选，将生成(金融局)副本）')
        if not srcs:
            print('[已取消]')
            return
        normal_srcs = srcs
        if normal_srcs:
            stat = run_remove_foreign_bank(normal_srcs, foreign_set, g.log_dir)
            print(f"[完成] Python 文件={stat['files']} 成功={stat['ok_files']} 分机构sheet={stat['branch_sheets']} 删除行={stat['deleted_rows']} 失败={stat['failed_files']}")
    elif choice == 'x14':
        srcs = _pick_source_files('选择要修改外汇页眉的Excel文件（可多选）')
        if not srcs:
            print('[已取消]')
            return
        stat = run_fx_header_fix(srcs, g.log_dir)
        print(f"[完成] 文件={stat['files']} 保存={stat['saved_files']} 外汇sheet={stat['fx_sheets']} 修改sheet={stat['modified_sheets']} 失败={stat['failed_files']}")
    elif choice == 'x15':
        srcs = _pick_source_files('选择要做地区总分校验的Excel文件（可多选）')
        if not srcs:
            print('[已取消]')
            return
        normal_srcs = srcs
        if normal_srcs:
            stat = run_region_sum_check(normal_srcs, g.output_dir, g.log_dir)
            print(f"[完成] Python 文件={stat['files']} 工作表={stat['sheets']} 错误={stat['errors']} 失败={stat['failed_files']} 输出={stat['path']}")
    elif choice == 'x16':
        tasks = load_extract_tasks(CFG_PATH)
        if not tasks:
            print("[已中止] '工作表提取'配置无启用任务")
            return
        srcs = _pick_source_files('选择要提取的源文件（可多选）', filetypes=[('Excel', '*.xlsx *.xlsm *.xls *.xlsb'), ('所有', '*.*')])
        if not srcs:
            print('[已取消]')
            return
        stat = run_extract_sheets(srcs, tasks, g.output_dir, g.log_dir)
        print(f"[完成] 源文件={stat['files']} 提取sheet={stat['extracted_sheets']} 输出文件={stat['output_files']} 失败={stat['failed_files']}")
        for p in stat['paths']:
            print(f'  -> {p}')
    elif choice == 'x18':
        print('→ 第一次文件选择：选择本期 专项二批次、商业银行共享二批（修改涉农贷款比上月）')
        cur = _pick_one_workbook('选择本期 专项二批次、商业银行共享二批（修改涉农贷款比上月）')
        if not cur:
            print('[已取消]')
            return
        print('→ 第二次文件选择：选择上期 专项二批次、商业银行共享二批（修改涉农贷款比上月）')
        prev = _pick_one_workbook('选择上期 专项二批次、商业银行共享二批（修改涉农贷款比上月）')
        if not prev:
            print('[已取消]')
            return
        normal_files = [cur, prev]
        try:
            r = run_adjust_rural_loan(cur, prev, g.log_dir)
        except RuntimeError as e:
            print(f'[Python 处理失败] {e}')
            return
        if r['error']:
            print(f"[失败] {r['error']}")
        else:
            print(f"[完成] 匹配={r['matched']} 分机构合计={r['fill_total']:.2f} 分地区总数={r['area_total']:.2f} {('校验通过' if r['consistent'] else '校验不一致')}")
            if r['unmatched_current']:
                print('  本期未匹配到上期: ' + '、'.join(r['unmatched_current']))
            if r['previous_only']:
                print('  上期独有: ' + '、'.join(r['previous_only']))
    elif choice == 'x19':
        print('→ 第一次文件选择：选择本期季度涉农贷款文件（将就地回填全部地区分机构表）')
        cur = _pick_one_workbook('选择本期季度涉农贷款文件')
        if not cur:
            print('[已取消]')
            return
        print('→ 第二次文件选择：选择上期季度涉农贷款文件')
        prev = _pick_one_workbook('选择上期季度涉农贷款文件')
        if not prev:
            print('[已取消]')
            return
        normal_files = [cur, prev]
        try:
            r = run_adjust_rural_loan_quarterly(cur, prev, g.log_dir)
        except RuntimeError as e:
            print(f'[Python 处理失败] {e}')
            return
        if r['error']:
            print(f"[失败] {r['error']}")
            return
        print(f"[完成] 地区={len(r['regions'])} 匹配机构={r['matched']} 分机构合计={r['fill_total']:.2f} 分地区合计={r['area_total']:.2f} {('校验通过' if r['consistent'] else '存在不一致')}")
        for item in r['regions']:
            if item['error']:
                print(f"  {item['region']}：{item['error']}")
            else:
                print(f"  {item['region']}：分机构={item['fill_total']:.2f} 分地区={item['area_total']:.2f} {('通过' if item['consistent'] else '不一致')}")
                if item['unmatched_current']:
                    print('    本期未匹配：' + '、'.join(item['unmatched_current']))
                if item['previous_only']:
                    print('    上期独有：' + '、'.join(item['previous_only']))
    elif choice == 's22':
        tmpl = _pick_one_workbook('选择批注汇总模板文件')
        if not tmpl:
            print('[已取消]')
            return
        srcs = _pick_source_files('选择按批注汇总的源文件（可多选）', filetypes=[('Excel', '*.xlsx *.xlsm *.xls *.xlsb'), ('所有', '*.*')])
        if not srcs:
            print('[已取消]')
            return
        stat = run_summary_by_comment(tmpl, srcs, g.output_dir, g.log_dir)
        if not stat.get('saved'):
            print('[完成] 无可汇总数据，未生成文件')
        else:
            print(f"[完成] 行数={stat['rows']} 输出={stat['path']}")

def main() -> None:
    if sys.platform == "win32":
        conversion_engines = "Windows 使用 Excel/WPS 表格转换 xls，使用 Word/WPS 文字转换 doc/docx"
    elif sys.platform.startswith("linux"):
        conversion_engines = "Linux 使用 LibreOffice Calc 转换 xls，使用 LibreOffice Writer 转换 doc/docx"
    else:
        conversion_engines = f"当前系统不支持文件格式转换（{sys.platform}）"
    print(f'NoComPyTools：表格处理使用 Python；{conversion_engines}。')
    while True:
        main_choice = _show_main_menu()
        if main_choice == '0':
            print('再见。')
            return
        if main_choice == '1':
            sub = _show_sub_menu(TIMELINE_MENU, ('1', '2', '3', '4', '5', '0'))
            if sub == '0':
                continue
            mapped = {'1': '1', '2': '2', '3': '3', '4': '4', '5': '5'}[sub]
        elif main_choice == '2':
            sub = _show_sub_menu(DEDUP_MENU, ('1', '2', '3', '4', '5', '6', '0'))
            if sub == '0':
                continue
            mapped = {'1': '6', '2': '7', '3': '8', '4': '9', '5': 'a', '6': 'b'}[sub]
        elif main_choice == '3':
            sub = _show_sub_menu(PRINT_MENU, ('1', '2', '3', '4', '0'))
            if sub == '0':
                continue
            mapped = {'1': 'p1', '2': 'p3', '3': 'p7', '4': 'p8'}[sub]
        elif main_choice == '4':
            sub = _show_sub_menu(CONVERT_MENU, ('1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '0'))
            if sub == '0':
                continue
            mapped = {'1': 't3', '2': 't4', '3': 't5', '4': 't7', '5': 't8', '6': 't9', '7': 't10', '8': 't11', '9': 't12', '10': 't13'}[sub]
        elif main_choice == '5':
            sub = _show_sub_menu(CONFIG_MENU, ('1', '2', '0'))
            if sub == '0':
                continue
            mapped = {'1': 'c', '2': 'cv'}[sub]
        elif main_choice == '6':
            sub = _show_sub_menu(SUMMARY_MENU, ('1', '2', '0'))
            if sub == '0':
                continue
            mapped = {'1': 's21', '2': 's22'}[sub]
        elif main_choice == '7':
            sub = _show_sub_menu(PENDING_MENU, ('1', '2', '3', '4', '5', '6', '7', '8', '0'))
            if sub == '0':
                continue
            mapped = {'1': 'x11', '2': 'x12', '3': 'x13', '4': 'x14', '5': 'x15', '6': 'x16', '7': 'x18', '8': 'x19'}[sub]
        elif main_choice == '8':
            sub = _show_sub_menu(ARCHIVE_MENU, ('1', '2', '3', '4', '5', '0'))
            if sub == '0':
                continue
            mapped = {'1': 'ar_date', '2': 'ar_type', '3': 'ar_exec', '4': 'ar_date_copy', '5': 'ar_type_copy'}[sub]
        else:
            sub = _show_sub_menu(QUICK_MENU, ('1', '2', '3', '0'))
            if sub == '0':
                continue
            mapped = {'1': 'q1', '2': 'q2', '3': 'q3'}[sub]
        try:
            dispatch(mapped)
        except Exception as e:
            print(f'[运行异常] {e}')
            import traceback
            traceback.print_exc()
if __name__ == '__main__':
    main()
