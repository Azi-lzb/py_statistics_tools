from __future__ import annotations

import shutil
import sys
import zipfile
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import Workbook

from PyCode import main
from PyCode.core import convert_33_34_35_37 as convert


def _libreoffice_is_installed() -> bool:
    if shutil.which("soffice") or shutil.which("libreoffice"):
        return True
    return sys.platform.startswith("linux") and next(
        Path("/opt").glob("libreoffice*/program/soffice"), None
    ) is not None


def _make_workbook(path: Path) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "Data"
    sheet.append(["name", "value"])
    sheet.append(["alpha", 7])
    book.save(path)


def _make_minimal_docx(path: Path) -> None:
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''
    relationships = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    document = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:r><w:t>pytools office conversion smoke</w:t></w:r></w:p><w:sectPr/>
</w:body></w:document>'''
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", relationships)
        archive.writestr("word/document.xml", document)


def test_python_mode_excel_xlsx_to_csv(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("PYTOOLS_MODE", "python")
    source = tmp_path / "source.xlsx"
    _make_workbook(source)

    stat = convert.run_excel_convert_34([source], "csv", tmp_path / "logs")

    output = tmp_path / "csv" / "source.csv"
    assert stat["engine"] == "Python"
    assert stat["ok"] == 1
    assert pd.read_csv(output).iloc[0].to_dict() == {"name": "alpha", "value": 7}


def test_python_mode_rejects_word_conversion(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("PYTOOLS_MODE", "python")
    source = tmp_path / "source.docx"
    source.write_bytes(b"fixture")

    with pytest.raises(RuntimeError, match="不提供 Word 格式转换"):
        convert.run_word_convert_36([source], "doc", tmp_path / "logs")


def test_engine_routing_follows_mode_and_platform(monkeypatch: pytest.MonkeyPatch) -> None:
    paths = [Path("one.xls"), Path("two.xlsx")]

    monkeypatch.setattr(main, "ENGINE_MODE", "office")
    monkeypatch.setattr(main.sys, "platform", "win32")
    assert main._split_sources_for_auto_route(paths) == ([], paths)

    monkeypatch.setattr(main, "ENGINE_MODE", "python")
    assert main._split_sources_for_auto_route(paths) == (paths, [])

    monkeypatch.setattr(main, "ENGINE_MODE", "office")
    monkeypatch.setattr(main.sys, "platform", "linux")
    assert main._split_sources_for_auto_route(paths) == (paths, [])


def test_libreoffice_executable_finds_official_opt_install(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    executable = tmp_path / "libreoffice25.8" / "program" / "soffice"
    executable.parent.mkdir(parents=True)
    executable.touch()

    opt_root = Path("/opt")
    original_glob = Path.glob

    def fake_glob(path: Path, pattern: str):
        if path == opt_root and pattern == "libreoffice*/program/soffice":
            return iter([executable])
        return original_glob(path, pattern)

    monkeypatch.setattr(convert.shutil, "which", lambda _name: None)
    monkeypatch.setattr(Path, "glob", fake_glob)
    monkeypatch.setattr(convert.os, "access", lambda *_args: True)

    assert convert._libreoffice_executable() == str(executable)


@pytest.mark.skipif(sys.platform != "win32", reason="Excel COM integration test runs on Windows")
def test_windows_office_mode_excel_xlsx_xls_roundtrip(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("PYTOOLS_MODE", "office")
    source = tmp_path / "source.xlsx"
    _make_workbook(source)

    to_xls = convert.run_excel_convert_34([source], "xls", tmp_path / "logs")
    xls_file = tmp_path / "xls" / "source.xls"
    to_xlsx = convert.run_excel_convert_34([xls_file], "xlsx", tmp_path / "logs")
    xlsx_file = tmp_path / "xls" / "xlsx" / "source.xlsx"

    assert to_xls["engine"] in {"Excel.Application", "ket.Application", "KET.Application"}
    assert to_xls["ok"] == 1 and to_xlsx["ok"] == 1
    assert pd.read_excel(xlsx_file).iloc[0].to_dict() == {"name": "alpha", "value": 7}


@pytest.mark.skipif(not sys.platform.startswith("linux"), reason="LibreOffice integration test runs on Linux")
@pytest.mark.skipif(not _libreoffice_is_installed(), reason="LibreOffice is not installed")
def test_linux_office_mode_excel_xlsx_xls_roundtrip(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("PYTOOLS_MODE", "office")
    source = tmp_path / "source.xlsx"
    _make_workbook(source)

    to_xls = convert.run_excel_convert_34([source], "xls", tmp_path / "logs")
    xls_file = tmp_path / "xls" / "source.xls"
    to_xlsx = convert.run_excel_convert_34([xls_file], "xlsx", tmp_path / "logs")
    xlsx_file = tmp_path / "xls" / "xlsx" / "source.xlsx"

    assert to_xls["engine"] == "LibreOffice" and to_xls["ok"] == 1
    assert to_xlsx["engine"] == "LibreOffice" and to_xlsx["ok"] == 1
    assert pd.read_excel(xlsx_file).iloc[0].to_dict() == {"name": "alpha", "value": 7}


@pytest.mark.skipif(not sys.platform.startswith("linux"), reason="LibreOffice integration test runs on Linux")
@pytest.mark.skipif(not _libreoffice_is_installed(), reason="LibreOffice is not installed")
def test_linux_office_mode_word_docx_doc_roundtrip(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("PYTOOLS_MODE", "office")
    source = tmp_path / "source.docx"
    _make_minimal_docx(source)

    to_doc = convert.run_word_convert_36([source], "doc", tmp_path / "logs")
    doc_file = tmp_path / "doc" / "source.doc"
    to_docx = convert.run_word_convert_36([doc_file], "docx", tmp_path / "logs")

    assert to_doc["engine"] == "LibreOffice" and to_doc["ok"] == 1
    assert to_docx["engine"] == "LibreOffice" and to_docx["ok"] == 1
    assert (tmp_path / "doc" / "docx" / "source.docx").is_file()
