from __future__ import annotations

from pathlib import Path
import sys

from .libreoffice import convert_with_libreoffice, libreoffice_executable
from .logger import get_logger


_WORD_FILTERS = {
    "doc": "MS Word 97",
    "docx": "Office Open XML Text",
}


def _next_output_path(path: Path) -> Path:
    if not path.exists():
        return path
    index = 1
    while True:
        candidate = path.with_name(f"{path.stem}_{index}{path.suffix}")
        if not candidate.exists():
            return candidate
        index += 1


def _start_windows_word_app():
    try:
        from win32com.client import DispatchEx
    except Exception as exc:
        raise RuntimeError("Windows 下使用 Word/WPS 转换需要 pywin32；请运行 启动WinExe.bat 安装依赖。") from exc

    last_error = None
    for prog_id in ("Word.Application", "kwps.Application", "KWPS.Application", "wps.Application"):
        try:
            app = DispatchEx(prog_id)
            app.Visible = False
            try:
                app.DisplayAlerts = 0
            except Exception:
                pass
            engine = "Microsoft Word" if prog_id == "Word.Application" else "WPS Writer"
            return app, engine
        except Exception as exc:
            last_error = exc
    raise RuntimeError(f"未找到可用的 Word/WPS 文字 COM 服务: {last_error}")


def _convert_with_windows_word(app, source: Path, output: Path, target_format: str) -> None:
    document = app.Documents.Open(
        str(source.resolve()), ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False,
    )
    try:
        output.parent.mkdir(parents=True, exist_ok=True)
        file_format = 0 if target_format == "doc" else 12
        try:
            document.SaveAs2(str(output.resolve()), FileFormat=file_format)
        except Exception:
            document.SaveAs(str(output.resolve()), FileFormat=file_format)
    finally:
        document.Close(SaveChanges=False)
    if not output.is_file():
        raise RuntimeError(f"Word/WPS 未生成目标文件: {output}")


def run_word_convert(files: list[Path], target_format: str, log_dir: Path) -> dict[str, object]:
    log = get_logger("word_convert", log_dir)
    ext = str(target_format or "docx").strip().lower().lstrip(".")
    if ext not in _WORD_FILTERS:
        ext = "docx"

    candidates: list[tuple[Path, Path]] = []
    skipped = 0
    for source in files:
        if not source.is_file() or source.suffix.lower() not in {".doc", ".docx"}:
            skipped += 1
            log.warning("skip file=%s reason=source_must_be_doc_or_docx", source)
            continue
        if source.suffix.lower() == f".{ext}":
            skipped += 1
            log.info("skip file=%s reason=source_equals_target", source)
            continue
        output_dir = source.parent / ext
        output = _next_output_path(output_dir / f"{source.stem}.{ext}")
        candidates.append((source, output))

    if not candidates:
        return {"ok": 0, "skip": skipped, "target_ext": ext, "engine": ""}

    succeeded = 0
    if sys.platform == "win32":
        try:
            app, engine = _start_windows_word_app()
        except Exception as exc:
            for source, _output in candidates:
                log.warning("skip file=%s reason=word_engine_unavailable err=%s", source, exc)
            return {
                "ok": 0, "skip": skipped + len(candidates), "target_ext": ext,
                "engine": "Word/WPS", "error": str(exc),
            }
        try:
            for source, output in candidates:
                try:
                    _convert_with_windows_word(app, source, output, ext)
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
        return {"ok": succeeded, "skip": skipped, "target_ext": ext, "engine": engine}

    if sys.platform.startswith("linux"):
        try:
            executable = libreoffice_executable()
        except Exception as exc:
            for source, _output in candidates:
                log.warning("skip file=%s reason=libreoffice_unavailable err=%s", source, exc)
            return {
                "ok": 0, "skip": skipped + len(candidates), "target_ext": ext,
                "engine": "LibreOffice Writer", "error": str(exc),
            }
        for source, output in candidates:
            try:
                convert_with_libreoffice(executable, source, output, ext, _WORD_FILTERS[ext])
                succeeded += 1
                log.info("ok src=%s out=%s engine=LibreOffice Writer", source, output)
            except Exception as exc:
                skipped += 1
                log.warning("skip file=%s reason=convert_failed err=%s", source, exc)
        return {"ok": succeeded, "skip": skipped, "target_ext": ext, "engine": "LibreOffice Writer"}

    error = f"当前系统不支持 Word 格式转换: {sys.platform}"
    for source, _output in candidates:
        log.warning("skip file=%s reason=unsupported_platform", source)
    return {
        "ok": 0, "skip": skipped + len(candidates), "target_ext": ext,
        "engine": "", "error": error,
    }
