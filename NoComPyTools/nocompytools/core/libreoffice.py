from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def libreoffice_executable() -> str:
    executable = shutil.which("soffice") or shutil.which("libreoffice")
    if executable:
        return executable

    if sys.platform == "win32":
        roots = [os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")]
        for root in roots:
            if root:
                for name in ("soffice.exe", "soffice.com"):
                    candidate = Path(root) / "LibreOffice" / "program" / name
                    if candidate.is_file():
                        return str(candidate)
    else:
        for candidate in sorted(Path("/opt").glob("libreoffice*/program/soffice"), reverse=True):
            if candidate.is_file() and os.access(candidate, os.X_OK):
                return str(candidate)

    raise RuntimeError("此格式转换需要 LibreOffice；请安装 LibreOffice Calc/Writer 后重试。")


def convert_with_libreoffice(
    executable: str, source: Path, output: Path, target_ext: str, output_filter: str,
) -> None:
    with tempfile.TemporaryDirectory(prefix="nocompytools-lo-") as temporary:
        root = Path(temporary)
        target_dir = root / "out"
        target_dir.mkdir()
        result = subprocess.run(
            [
                executable, "--headless", "--nologo", "--nodefault", "--nolockcheck",
                "--nofirststartwizard", f"-env:UserInstallation={(root / 'profile').as_uri()}",
                "--convert-to", f"{target_ext}:{output_filter}",
                "--outdir", str(target_dir), str(source),
            ],
            capture_output=True,
            text=True,
            timeout=180,
            check=False,
        )
        converted = target_dir / f"{source.stem}.{target_ext}"
        if result.returncode != 0 or not converted.is_file():
            raise RuntimeError(
                (result.stderr or result.stdout).strip()
                or f"LibreOffice 未生成 {converted.name}"
            )
        output.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(converted), str(output))
