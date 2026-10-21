from pathlib import Path

from nocompytools.main import (
    _collect_excel_files,
    _collect_excel_files_from_folders,
    _collect_files,
    _collect_files_from_folders,
)


def test_collect_excel_files_current_folder_only_uses_xlsx_and_xlsm(tmp_path: Path):
    (tmp_path / "a.xlsx").touch()
    (tmp_path / "b.xlsm").touch()
    (tmp_path / "c.xls").touch()
    (tmp_path / "~$lock.xlsx").touch()
    (tmp_path / "note.csv").touch()
    nested = tmp_path / "nested"
    nested.mkdir()
    (nested / "d.xlsx").touch()

    assert [p.name for p in _collect_excel_files(tmp_path, recursive=False)] == [
        "a.xlsx", "b.xlsm",
    ]


def test_collect_excel_files_recursively_excludes_legacy_xls(tmp_path: Path):
    (tmp_path / "root.xlsx").touch()
    nested = tmp_path / "nested"
    nested.mkdir()
    (nested / "child.xls").touch()

    assert [p.relative_to(tmp_path).as_posix() for p in _collect_excel_files(tmp_path, recursive=True)] == [
        "root.xlsx",
    ]


def test_collect_excel_files_from_multiple_folders_deduplicates_nested_sources(tmp_path: Path):
    parent = tmp_path / "parent"
    child = parent / "child"
    child.mkdir(parents=True)
    (parent / "root.xlsx").touch()
    (child / "nested.xlsm").touch()

    sources = _collect_excel_files_from_folders([parent, child], recursive=True)

    assert [p.relative_to(tmp_path).as_posix() for p in sources] == [
        "parent/child/nested.xlsm", "parent/root.xlsx",
    ]


def test_collect_files_supports_all_files_and_suffix_filter(tmp_path: Path):
    (tmp_path / "root.xlsx").touch()
    (tmp_path / "note.txt").touch()
    child = tmp_path / "child"
    child.mkdir()
    (child / "child.docx").touch()

    assert [p.name for p in _collect_files(tmp_path, recursive=False)] == ["note.txt", "root.xlsx"]
    assert [p.name for p in _collect_files(tmp_path, recursive=True, suffixes={".docx"})] == ["child.docx"]
    assert [p.name for p in _collect_files_from_folders([tmp_path, child], recursive=True)] == [
        "child.docx", "note.txt", "root.xlsx",
    ]
