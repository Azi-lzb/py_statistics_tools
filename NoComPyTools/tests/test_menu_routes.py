from __future__ import annotations

from nocompytools import main as app


def test_every_current_menu_action_routes_to_python_dispatch(monkeypatch) -> None:
    menus = [
        ("1", app.TIMELINE_MENU, ("1", "2", "3", "4", "5"), ("1", "2", "3", "4", "5")),
        ("2", app.DEDUP_MENU, ("1", "2", "3", "4", "5", "6"), ("6", "7", "8", "9", "a", "b")),
        ("3", app.PRINT_MENU, ("1", "2", "3", "4"), ("p1", "p3", "p7", "p8")),
        ("4", app.CONVERT_MENU, ("1", "2", "3", "4", "5", "6", "7", "8", "9", "10"),
         ("t3", "t4", "t5", "t7", "t8", "t9", "t10", "t11", "t12", "t13")),
        ("5", app.CONFIG_MENU, ("1", "2"), ("c", "cv")),
        ("6", app.SUMMARY_MENU, ("1", "2"), ("s21", "s22")),
        ("7", app.PENDING_MENU, ("1", "2", "3", "4", "5", "6", "7", "8"),
         ("x11", "x12", "x13", "x14", "x15", "x16", "x18", "x19")),
        ("8", app.ARCHIVE_MENU, ("1", "2", "3", "4", "5"),
         ("ar_date", "ar_type", "ar_exec", "ar_date_copy", "ar_type_copy")),
        ("9", app.QUICK_MENU, ("1", "2", "3"), ("q1", "q2", "q3")),
    ]
    main_answers = [main_choice for main_choice, _, choices, _ in menus for _ in choices]
    sub_answers = [choice for _, _, choices, _ in menus for choice in choices]
    expected = [command for _, _, _, commands in menus for command in commands]
    main_answers.append("0")
    main_iter, sub_iter = iter(main_answers), iter(sub_answers)
    seen_menus = []
    routed = []

    monkeypatch.setattr(app, "_show_main_menu", lambda: next(main_iter))

    def show_submenu(menu, choices):
        seen_menus.append((menu, tuple(choices)))
        return next(sub_iter)

    monkeypatch.setattr(app, "_show_sub_menu", show_submenu)
    monkeypatch.setattr(app, "dispatch", routed.append)

    app.main()

    expected_menus = [
        (menu, choices + ("0",))
        for _, menu, choices, _ in menus
        for _ in choices
    ]
    assert seen_menus == expected_menus
    assert routed == expected
    assert list(main_iter) == []
    assert list(sub_iter) == []
    assert "3  按配置打印（打印配置）" in app.PRINT_MENU
    assert "按配置提取数据" not in app.PRINT_MENU


def test_survey_stats_option_is_not_available(monkeypatch, capsys) -> None:
    answers = iter(("s", "0"))
    monkeypatch.setattr("builtins.input", lambda _prompt: next(answers))

    assert app._show_main_menu() == "0"
    output = capsys.readouterr().out
    assert "调研选项统计" not in app.MAIN_MENU
    assert "仅接受 1/2/3/4/5/6/7/8/9/0" in output
