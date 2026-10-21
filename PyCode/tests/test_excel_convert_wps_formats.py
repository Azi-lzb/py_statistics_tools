from pathlib import Path
import sys

from PyCode.core import convert_33_34_35_37 as convert


class _FakeApp:
    def Quit(self):
        pass


def test_excel_like_accepts_wps_formats():
    assert convert._excel_like(Path("source.et"))
    assert convert._excel_like(Path("template.ett"))


def test_wps_source_uses_platform_office_engine(tmp_path, monkeypatch):
    monkeypatch.setenv("PYTOOLS_MODE", "office")
    source = tmp_path / "source.et"
    source.touch()
    calls = []

    def fake_start_excel_app(prefer_wps=False):
        calls.append(prefer_wps)
        return _FakeApp(), "ket.Application"

    monkeypatch.setattr(convert, "_start_excel_app", fake_start_excel_app)
    monkeypatch.setattr(convert, "_excel_convert_one", lambda app, src, out, ext: None)
    if sys.platform.startswith("linux"):
        def fake_libreoffice_convert(executable, src, out, target_ext, output_filter):
            out.touch()

        monkeypatch.setattr(convert, "_libreoffice_executable", lambda: "soffice")
        monkeypatch.setattr(convert, "_libreoffice_convert_one", fake_libreoffice_convert)

    result = convert.run_excel_convert_34([source], "xlsx", tmp_path / "logs")

    if sys.platform.startswith("linux"):
        assert calls == []
        assert result["engine"] == "LibreOffice"
    else:
        assert calls == [True]
    assert result["target_ext"] == "xlsx"
    assert result["ok"] == 1
