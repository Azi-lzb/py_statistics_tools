from __future__ import annotations

import json
from pathlib import Path

import pytest

from PyCode.core import app_identity
from PyCode.core.config_xlsx import PathMapRule
from PyCode.core.path_map import _match_kw, standardize


def _rule(**changes) -> PathMapRule:
    values = {
        "enabled": True,
        "name": "map",
        "applicable_rule": "",
        "wb_keyword": "",
        "sheet_keyword": "",
        "target": "两者",
        "match_mode": "精确",
        "original": "旧路径",
        "standard": "新路径",
    }
    values.update(changes)
    return PathMapRule(**values)


@pytest.mark.parametrize(
    ("keyword", "name", "expected"),
    [
        ("", "anything", True),
        (" ;；，, ", "anything", True),
        ("银行；报表,2026", "2026银行月报表", True),
        ("银行,不存在", "银行月报表", False),
        ("银行", None, False),
    ],
)
def test_match_keywords(keyword, name, expected):
    assert _match_kw(keyword, name) is expected


def test_standardize_filters_rule_target_keywords_and_match_mode():
    path = "机构/旧路径/旧路径"
    maps = [
        _rule(applicable_rule="other"),
        _rule(target="列头"),
        _rule(wb_keyword="另一个文件"),
        _rule(sheet_keyword="资产负债"),
        _rule(match_mode="包含", original="旧路径", standard="标准"),
    ]
    assert standardize(path, "行头", "rule", "2026银行.xlsx", "利润表", maps) == "机构/标准/标准"

    assert standardize("旧路径", "列头", "rule", "book.xlsx", "表", [_rule()]) == "新路径"
    assert standardize("old", "行头", "rule", "book.xlsx", "表", [_rule()]) == "old"
    assert standardize("旧路径", "行头", "rule", "book.xlsx", "表", [_rule(standard="标准")]) == "标准"
    assert standardize("", "行头", "rule", "book.xlsx", "表", [_rule()]) == ""


@pytest.mark.parametrize("frozen", [False, True])
def test_manifest_path_uses_source_or_frozen_root(monkeypatch, tmp_path, frozen):
    root = tmp_path / ("frozen" if frozen else "source")
    monkeypatch.setattr(app_identity.sys, "frozen", frozen, raising=False)
    if frozen:
        monkeypatch.setattr(app_identity.sys, "_MEIPASS", str(root), raising=False)
    else:
        monkeypatch.setattr(app_identity, "__file__", str(root / "PyCode" / "core" / "app_identity.py"))
    assert app_identity.manifest_path() == root / "版本管理" / "VERSION_MANIFEST.json"


def test_product_info_validates_manifest_and_formats_titles(tmp_path, monkeypatch, capsys):
    manifest = tmp_path / "manifest.json"
    manifest.write_text(json.dumps({"product": {"name": "pytools", "version": "2.4.1"}}), encoding="utf-8")
    assert app_identity.load_product_info(manifest) == {
        "name": "pytools", "version": "2.4.1", "title": "pytools V2.4.1"
    }
    monkeypatch.setattr(app_identity, "manifest_path", lambda: manifest)
    assert app_identity.app_version() == "V2.4.1"
    assert app_identity.app_title() == "pytools V2.4.1"

    monkeypatch.setattr(app_identity.sys, "argv", ["app_identity", "--title"])
    assert app_identity.main() == 0
    assert capsys.readouterr().out.strip() == "pytools V2.4.1"
    monkeypatch.setattr(app_identity.sys, "argv", ["app_identity", "--version"])
    assert app_identity.main() == 0
    assert capsys.readouterr().out.strip() == "2.4.1"


@pytest.mark.parametrize(
    "payload",
    [
        {"product": {"name": "", "version": "1.0.0"}},
        {"product": {"name": "bad/name", "version": "1.0.0"}},
        {"product": {"name": "bad.", "version": "1.0.0"}},
        {"product": {"name": "pytools", "version": "1.0"}},
        {"product": {"name": 1, "version": "1.0.0"}},
        {"product": None},
        [],
    ],
)
def test_product_info_rejects_invalid_manifests(tmp_path, payload):
    manifest = tmp_path / "manifest.json"
    manifest.write_text(json.dumps(payload), encoding="utf-8")
    with pytest.raises(RuntimeError, match="产品版本清单不可用"):
        app_identity.load_product_info(manifest)


def test_product_info_rejects_missing_or_malformed_file(tmp_path):
    with pytest.raises(RuntimeError):
        app_identity.load_product_info(tmp_path / "missing.json")
    malformed = tmp_path / "malformed.json"
    malformed.write_text("{", encoding="utf-8")
    with pytest.raises(RuntimeError):
        app_identity.load_product_info(malformed)
