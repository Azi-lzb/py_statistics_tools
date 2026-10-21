from __future__ import annotations

from pathlib import Path
import sys

from .libreoffice import convert_with_libreoffice, libreoffice_executable
from .logger import get_logger


def _libreoffice_executable() -> str:
    return libreoffice_executable()


def _convert_one(executable: str, source: Path, output: Path) -> None:
    convert_with_libreoffice(
        executable, source, output, "xlsx", "Calc MS Excel 2007 XML",
    )


def _start_windows_office():
    try:
        from win32com.client import DispatchEx
    except Exception as exc:
        raise RuntimeError("Windows 下使用 Excel/WPS 转换需要 pywin32；请运行 启动WinExe.bat 安装依赖。") from exc

    last_error = None
    for prog_id in ("Excel.Application", "ket.Application", "KET.Application"):
        try:
            app = DispatchEx(prog_id)
            app.Visible = False
            app.DisplayAlerts = False
            for attr, value in (("ScreenUpdating", False), ("EnableEvents", False),
                                ("AskToUpdateLinks", False)):
                try:
                    setattr(app, attr, value)
                except Exception:
                    pass
            return app, prog_id
        except Exception as exc:
            last_error = exc
    raise RuntimeError(f"未找到可用的 Excel/WPS 表格 COM 服务: {last_error}")


def _convert_one_with_office(app, source: Path, output: Path) -> None:
    workbook = app.Workbooks.Open(str(source.resolve()), UpdateLinks=0, ReadOnly=True)
    try:
        workbook.SaveAs(str(output.resolve()), FileFormat=51, CreateBackup=False)
    finally:
        workbook.Close(SaveChanges=False)
    if not output.is_file():
        raise RuntimeError(f"Excel/WPS 未生成目标文件: {output}")


def run_xls_to_xlsx(files: list[Path], log_dir: Path) -> dict[str, object]:
    log = get_logger("xls_to_xlsx", log_dir)
    candidates: list[tuple[Path, Path]] = []
    skipped = 0
    for source in files:
        if not source.is_file() or source.suffix.lower() != ".xls":
            skipped += 1
            log.warning("skip file=%s reason=source_must_be_xls", source)
            continue
        output_dir = source.parent / "xlsx"
        output_dir.mkdir(parents=True, exist_ok=True)
        output = output_dir / f"{source.stem}.xlsx"
        if output.exists():
            index = 1
            while output.exists():
                output = output_dir / f"{source.stem}_{index}.xlsx"
                index += 1
        candidates.append((source, output))

    if not candidates:
        return {"ok": 0, "skip": skipped, "target_ext": "xlsx", "engine": ""}

    succeeded = 0
    if sys.platform == "win32":
        try:
            app, engine = _start_windows_office()
        except Exception as exc:
            for source, _output in candidates:
                log.warning("skip file=%s reason=office_engine_unavailable err=%s", source, exc)
            return {
                "ok": 0, "skip": skipped + len(candidates), "target_ext": "xlsx",
                "engine": "Excel/WPS", "error": str(exc),
            }
        try:
            for source, output in candidates:
                try:
                    _convert_one_with_office(app, source, output)
                    succeeded += 1
                    log.info("ok src=%s out=%s engine=%s", source, output, engine)
                except Exception as exc:
                    skipped += 1
                    log.warning("skip file=%s reason=convert_failed err=%s", source, exc)
        finally:
            try:
                app.Quit()
            except Exception:
                pass
    elif sys.platform.startswith("linux"):
        try:
            executable = _libreoffice_executable()
        except Exception as exc:
            for source, _output in candidates:
                log.warning("skip file=%s reason=libreoffice_unavailable err=%s", source, exc)
            return {
                "ok": 0, "skip": skipped + len(candidates), "target_ext": "xlsx",
                "engine": "LibreOffice", "error": str(exc),
            }
        engine = "LibreOffice"
        for source, output in candidates:
            try:
                _convert_one(executable, source, output)
                succeeded += 1
                log.info("ok src=%s out=%s engine=%s", source, output, engine)
            except Exception as exc:
                skipped += 1
                log.warning("skip file=%s reason=convert_failed err=%s", source, exc)
    else:
        error = f"当前系统不支持 .xls 转换: {sys.platform}"
        for source, _output in candidates:
            log.warning("skip file=%s reason=unsupported_platform", source)
        return {
            "ok": 0, "skip": skipped + len(candidates), "target_ext": "xlsx",
            "engine": "", "error": error,
        }
    return {"ok": succeeded, "skip": skipped, "target_ext": "xlsx", "engine": engine}
