from pathlib import Path

from PyCode.core import convert_33_34_35_37 as convert


class _FakeApp:
    def Quit(self):
        pass


def test_excel_like_accepts_wps_formats():
    assert convert._excel_like(Path("source.et"))
    assert convert._excel_like(Path("template.ett"))


def test_wps_source_prefers_wps_com(tmp_path, monkeypatch):
    source = tmp_path / "source.et"
    source.touch()
    calls = []

    def fake_start_excel_app(prefer_wps=False):
        calls.append(prefer_wps)
        return _FakeApp(), "ket.Application"

    monkeypatch.setattr(convert, "_start_excel_app", fake_start_excel_app)
    monkeypatch.setattr(convert, "_excel_convert_one", lambda app, src, out, ext: None)

    result = convert.run_excel_convert_34([source], "xlsx", tmp_path / "logs")

    assert calls == [True]
    assert result["target_ext"] == "xlsx"
    assert result["ok"] == 1
