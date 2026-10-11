from types import SimpleNamespace

from PyCode import main


def test_quick_output_folder_creates_and_opens_configured_directory(tmp_path, monkeypatch):
    output_dir = tmp_path / "output"
    opened = []
    monkeypatch.setattr(main, "load_global", lambda _cfg: SimpleNamespace(output_dir=output_dir))
    monkeypatch.setattr(main, "_open_path", lambda path: opened.append(path))

    main.dispatch("q3")

    assert output_dir.is_dir()
    assert opened == [output_dir]
