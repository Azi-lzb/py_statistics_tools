from pathlib import Path

import pytest

from nocompytools.core import word_convert


def test_word_conversion_uses_libreoffice_and_suffixes_duplicate_outputs(tmp_path: Path, monkeypatch) -> None:
    source_doc = tmp_path / "legacy.doc"
    source_doc.write_bytes(b"doc")
    source_docx = tmp_path / "already.docx"
    source_docx.write_bytes(b"docx")
    invalid = tmp_path / "note.txt"
    invalid.write_text("not word", encoding="utf-8")
    calls = []

    monkeypatch.setattr(word_convert, "libreoffice_executable", lambda: "soffice")

    def fake_convert(executable, source, output, target_ext, output_filter):
        calls.append((executable, source, output, target_ext, output_filter))
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_bytes(b"converted")

    monkeypatch.setattr(word_convert, "convert_with_libreoffice", fake_convert)
    result = word_convert.run_word_convert(
        [source_doc, source_docx, invalid], "docx", tmp_path / "logs",
    )

    assert result == {"ok": 1, "skip": 2, "target_ext": "docx", "engine": "LibreOffice"}
    assert calls == [(
        "soffice", source_doc, tmp_path / "docx" / "legacy.docx",
        "docx", "Office Open XML Text",
    )]
    assert (tmp_path / "docx" / "legacy.docx").read_bytes() == b"converted"

    second = word_convert.run_word_convert([source_doc], "docx", tmp_path / "logs")
    assert second["ok"] == 1
    assert (tmp_path / "docx" / "legacy_1.docx").is_file()


def test_word_conversion_empty_or_same_format_input_does_not_require_libreoffice(
    tmp_path: Path, monkeypatch,
) -> None:
    same_format = tmp_path / "source.docx"
    same_format.write_bytes(b"docx")
    monkeypatch.setattr(
        word_convert, "libreoffice_executable",
        lambda: pytest.fail("LibreOffice should not be discovered when nothing can be converted"),
    )

    result = word_convert.run_word_convert([same_format], "docx", tmp_path / "logs")

    assert result == {"ok": 0, "skip": 1, "target_ext": "docx", "engine": ""}
