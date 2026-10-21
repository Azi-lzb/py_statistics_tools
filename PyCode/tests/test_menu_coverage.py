from __future__ import annotations

from PyCode import main as app


def test_every_main_menu_item_routes_to_dispatch(monkeypatch) -> None:
    main_choices = [
        *(["1"] * 5), *(["2"] * 6), *(["3"] * 5), *(["4"] * 10),
        *(["5"] * 2), *(["6"] * 2), *(["7"] * 8), *(["8"] * 5),
        *(["9"] * 3), "0",
    ]
    submenu_choices = [
        *[str(n) for n in range(1, 6)],
        *[str(n) for n in range(1, 7)],
        *[str(n) for n in range(1, 6)],
        *[str(n) for n in range(1, 11)],
        "1", "2", "1", "2",
        *[str(n) for n in range(1, 9)],
        *[str(n) for n in range(1, 6)],
        *[str(n) for n in range(1, 4)],
    ]
    main_iter = iter(main_choices)
    sub_iter = iter(submenu_choices)
    routed: list[str] = []
    monkeypatch.setattr(app, "_select_engine_mode", lambda: "python")
    monkeypatch.setattr(app, "_show_main_menu", lambda: next(main_iter))
    monkeypatch.setattr(app, "_show_sub_menu", lambda *_args, **_kwargs: next(sub_iter))
    monkeypatch.setattr(app, "dispatch", routed.append)

    app.main()

    assert routed == [
        "1com", "2com", "3com", "4com", "5",
        "6", "7", "8", "9", "a", "b",
        "p1com", "p3com", "p7com", "p8", "ppdf",
        "t3", "t4", "t5", "t7com", "t8", "t9", "t10", "t11", "t12", "t13",
        "c", "cv", "s21com", "s22com",
        "x11", "x12", "x13", "x14com", "x15", "x16com", "x18", "x19",
        "ar_date", "ar_type", "ar_exec", "ar_date_copy", "ar_type_copy",
        "q1", "q2", "q3",
    ]
    assert list(sub_iter) == []


def test_main_menu_rejects_removed_survey_shortcut(monkeypatch, capsys) -> None:
    answers = iter(["bad", "s", "0"])
    monkeypatch.setattr("builtins.input", lambda _prompt="": next(answers))

    assert app._show_main_menu() == "0"
    output = capsys.readouterr().out
    assert "仅接受" in output
    assert "调研选项统计" not in output
