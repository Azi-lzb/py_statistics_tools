"""产品名称与发布版本的统一读取（源码与打包后的 exe 共用）。"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

VERSION_RE = re.compile(r"\d+\.\d+\.\d+")


def manifest_path() -> Path:
    """源码和冻结包均读取构建时打包的同一份版本清单。"""
    if getattr(sys, "frozen", False):
        root = Path(sys._MEIPASS)
    else:
        root = Path(__file__).resolve().parents[2]
    return root / "版本管理" / "VERSION_MANIFEST.json"


def load_product_info(path: Path | None = None) -> dict[str, str]:
    source = Path(path) if path is not None else manifest_path()
    try:
        product = json.loads(source.read_text(encoding="utf-8"))["product"]
        name, version = product["name"], product["version"]
        if not all(isinstance(v, str) and v for v in (name, version)):
            raise ValueError("产品字段必须为非空文字")
        if not VERSION_RE.fullmatch(version):
            raise ValueError("产品版本必须为 MAJOR.MINOR.PATCH 三段数字")
        if re.search(r'[<>:"/\\|?*\x00-\x1f]', name) or name.endswith((".", " ")):
            raise ValueError("产品名称含文件名不支持的字符")
    except (OSError, ValueError, KeyError, TypeError) as exc:
        raise RuntimeError(
            f"产品版本清单不可用：{source.name}。请使用包含完整版本清单的程序包。"
        ) from exc
    return {"name": name, "version": version, "title": f"{name} V{version}"}


def app_version() -> str:
    return "V" + load_product_info()["version"]


def app_title() -> str:
    return load_product_info()["title"]


def main() -> int:
    import argparse

    parser = argparse.ArgumentParser(description="读取 pytools 的统一产品信息")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--version", action="store_true", help="仅版本号，如 1.0.0")
    group.add_argument("--title", action="store_true", help="完整标题，如 pytools V1.0.0")
    args = parser.parse_args()
    if args.title:
        print(app_title())
    else:
        print(load_product_info()["version"])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
