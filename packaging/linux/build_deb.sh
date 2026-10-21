#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
ROOT=$(cd -- "$SCRIPT_DIR/../.." && pwd)
DIST=${PYTOOLS_LINUX_DIST_DIR:-$ROOT/dist/linux}

STANDALONE_TAG=${STANDALONE_TAG:-20260924}
PYTHON_VERSION=${PYTHON_VERSION:-3.12.14}
PYINSTALLER_VERSION=${PYINSTALLER_VERSION:-6.22.3}
STANDALONE_MIRROR=${STANDALONE_MIRROR:-https://mirror.nju.edu.cn/github-release/astral-sh/python-build-standalone}
STANDALONE_FALLBACK=${STANDALONE_FALLBACK:-https://github.com/astral-sh/python-build-standalone/releases/download}
GLIBC_TARGET=${GLIBC_TARGET:-2.28}
# Kylin V10 SP1 server packages include GCC 7.3/libstdc++; stay at its GLIBCXX ABI ceiling.
GLIBCXX_TARGET=${GLIBCXX_TARGET:-3.4.24}

for tool in dpkg dpkg-deb objdump tar sha256sum; do
    command -v "$tool" >/dev/null || { echo "缺少构建工具：$tool" >&2; exit 1; }
done
FETCH=$(command -v curl || command -v wget || true)
[ -n "$FETCH" ] || { echo '需要 curl 或 wget 下载独立 Python。' >&2; exit 1; }

ARCH=$(dpkg --print-architecture)
case "$ARCH" in
    amd64) STANDALONE_ARCH=x86_64 ;;
    arm64) STANDALONE_ARCH=aarch64 ;;
    *) echo "暂不支持的 Debian 架构：$ARCH（支持 amd64/arm64）" >&2; exit 1 ;;
esac

CACHE_HOME=${XDG_CACHE_HOME:-${HOME:?}/.cache}
CACHE_DIR=$CACHE_HOME/pytools/linux
RUNTIME_DIR=$CACHE_DIR/runtime
ASSET="cpython-${PYTHON_VERSION}+${STANDALONE_TAG}-${STANDALONE_ARCH}-unknown-linux-gnu-install_only.tar.gz"
TARBALL=$RUNTIME_DIR/$ASSET
EXTRACT_DIR=$CACHE_DIR/${ASSET%.tar.gz}
BUILD_ENV=$CACHE_DIR/buildenv-${ARCH}-py${PYTHON_VERSION%.*}

mkdir -p "$RUNTIME_DIR" "$CACHE_DIR"
if [ ! -f "$TARBALL" ]; then
    echo "[Python] 下载 CPython $PYTHON_VERSION（$STANDALONE_ARCH，glibc 2.17 基线）..."
    encoded_asset=${ASSET/+/%2B}
    downloaded=0
    for url in "$STANDALONE_MIRROR/$STANDALONE_TAG/$encoded_asset" "$STANDALONE_FALLBACK/$STANDALONE_TAG/$ASSET"; do
        if [ "${FETCH##*/}" = curl ]; then
            curl -fL --retry 2 --connect-timeout 30 -o "$TARBALL.part" "$url" && downloaded=1
        else
            wget --tries=2 --timeout=60 -O "$TARBALL.part" "$url" && downloaded=1
        fi
        [ "$downloaded" -eq 1 ] && break
    done
    [ "$downloaded" -eq 1 ] || { echo "Python 下载失败，缓存位置：$TARBALL" >&2; exit 1; }
    mv "$TARBALL.part" "$TARBALL"
fi

if [ ! -x "$EXTRACT_DIR/python/bin/python3" ]; then
    mkdir -p "$EXTRACT_DIR"
    tar -xzf "$TARBALL" -C "$EXTRACT_DIR"
fi
STANDALONE_PY=$EXTRACT_DIR/python/bin/python3

if [ ! -x "$BUILD_ENV/bin/python" ]; then
    "$STANDALONE_PY" -m venv "$BUILD_ENV"
fi
BUILD_PY=$BUILD_ENV/bin/python
"$BUILD_PY" -m pip install --disable-pip-version-check -r "$SCRIPT_DIR/requirements.txt"
"$BUILD_PY" -c 'import pandas, openpyxl, xlrd, pyxlsb, PyInstaller; print("[依赖] 导入检查通过")'

BUILD_ROOT=$(mktemp -d "${TMPDIR:-/tmp}/pytools-deb.XXXXXX")
trap 'rm -rf -- "$BUILD_ROOT"' EXIT
SOURCE=$BUILD_ROOT/source
PACKAGE=$BUILD_ROOT/package
mkdir -p "$SOURCE" "$PACKAGE"
cp -a "$ROOT/PyCode" "$SOURCE/PyCode"
cp -a "$ROOT/pytools_entry.py" "$SOURCE/"
cp -a "$ROOT/版本管理" "$SOURCE/版本管理"
rm -rf -- "$SOURCE/PyCode/tests" "$SOURCE/PyCode/__pycache__" "$SOURCE/PyCode/core/__pycache__"

APP_VERSION=$($BUILD_PY -c 'import json,sys; print(json.load(open(sys.argv[1], encoding="utf-8"))["product"]["version"])' "$SOURCE/版本管理/VERSION_MANIFEST.json")
echo "[PyInstaller] 构建 pytools V$APP_VERSION（$ARCH）..."
(
    cd "$SOURCE"
    "$BUILD_PY" -m PyInstaller --clean --noconfirm --onedir --name pytools \
        --add-data "版本管理/VERSION_MANIFEST.json:版本管理" \
        --collect-all pandas --hidden-import xlrd --hidden-import pyxlsb \
        pytools_entry.py >"$BUILD_ROOT/pyinstaller.log" 2>&1 || {
        tail -n 60 "$BUILD_ROOT/pyinstaller.log" >&2
        exit 1
    }
)
BUNDLE=$SOURCE/dist/pytools
[ -x "$BUNDLE/pytools" ] || { echo 'PyInstaller 未生成 Linux 可执行文件。' >&2; exit 1; }

# Let Kylin provide its own ABI-compatible C++ runtime instead of bundling the
# build host's newer libstdc++/libgcc. Keep the frozen payload self-contained otherwise.
rm -f "$BUNDLE/_internal/libstdc++.so.6" "$BUNDLE/_internal/libgcc_s.so.1"
printf '2\n0\n' | "$BUNDLE/pytools" >"$BUILD_ROOT/smoke.log" 2>&1 || {
    cat "$BUILD_ROOT/smoke.log" >&2
    echo '移除构建机 C++ 运行库后，冻结程序启动检查失败。' >&2
    exit 1
}
SYMBOLS=$(find "$BUNDLE" -type f -print0 | xargs -0 -r objdump -T 2>/dev/null | grep -oE 'GLIBC_[0-9]+(\.[0-9]+)?|GLIBCXX_[0-9]+(\.[0-9]+)+' || true)
MAX_GLIBC=$(printf '%s\n' "$SYMBOLS" | sed -n 's/^GLIBC_//p' | sort -Vu | tail -n 1)
MAX_GLIBCXX=$(printf '%s\n' "$SYMBOLS" | sed -n 's/^GLIBCXX_//p' | sort -Vu | tail -n 1)
[ -n "$MAX_GLIBC" ] || { echo 'ELF 审计未发现 GLIBC 版本需求，无法确认兼容性。' >&2; exit 1; }
echo "[审计] 最高 GLIBC 符号：$MAX_GLIBC；最高 GLIBCXX 符号：${MAX_GLIBCXX:-无}"

version_gt() {
    [ "$(printf '%s\n%s\n' "$1" "$2" | sort -V | tail -n 1)" = "$1" ] && [ "$1" != "$2" ]
}
if version_gt "$MAX_GLIBC" "$GLIBC_TARGET" || { [ -n "$MAX_GLIBCXX" ] && version_gt "$MAX_GLIBCXX" "$GLIBCXX_TARGET"; }; then
    echo "冻结产物超过兼容上限 GLIBC <= $GLIBC_TARGET / GLIBCXX <= $GLIBCXX_TARGET。" >&2
    while IFS= read -r -d '' file; do
        file_symbols=$(objdump -T "$file" 2>/dev/null || true)
        file_glibc=$(printf '%s\n' "$file_symbols" | grep -oE 'GLIBC_[0-9]+(\.[0-9]+)?' | sed 's/^GLIBC_//' | sort -Vu | tail -n 1 || true)
        file_glibcxx=$(printf '%s\n' "$file_symbols" | grep -oE 'GLIBCXX_[0-9]+(\.[0-9]+)+' | sed 's/^GLIBCXX_//' | sort -Vu | tail -n 1 || true)
        if { [ -n "$file_glibc" ] && version_gt "$file_glibc" "$GLIBC_TARGET"; } || { [ -n "$file_glibcxx" ] && version_gt "$file_glibcxx" "$GLIBCXX_TARGET"; }; then
            printf '  %s GLIBC<=%s GLIBCXX<=%s\n' "${file#"$BUNDLE"/}" "${file_glibc:-无}" "${file_glibcxx:-无}" >&2
        fi
    done < <(find "$BUNDLE" -type f -print0)
    exit 1
fi

mkdir -p "$PACKAGE/opt/pytools" "$PACKAGE/usr/bin" "$PACKAGE/usr/share/pytools/config" \
    "$PACKAGE/usr/share/applications" "$PACKAGE/usr/share/doc/pytools" "$PACKAGE/DEBIAN"
cp -a "$BUNDLE/." "$PACKAGE/opt/pytools/"
cp "$ROOT/PyCode/config/config.xlsx" "$PACKAGE/usr/share/pytools/config/config.xlsx"
"$BUILD_PY" - "$PACKAGE/usr/share/pytools/config/config.xlsx" <<'PY'
import sys
from openpyxl import load_workbook

path = sys.argv[1]
book = load_workbook(path)
sheet = book["全局配置"]
for row in sheet.iter_rows(min_row=2):
    if row[0].value in {"输入目录", "输出目录", "日志目录"}:
        row[1].value = None
book.save(path)
PY

cp "$SCRIPT_DIR/pytools-launcher" "$PACKAGE/usr/bin/pytools"
cp "$SCRIPT_DIR/pytools.desktop" "$PACKAGE/usr/share/applications/pytools.desktop"
chmod 755 "$PACKAGE/usr/bin/pytools" "$PACKAGE/opt/pytools/pytools"

BUILD_DATE=$(date -u +%Y%m%d)
DEB_VERSION="$APP_VERSION-$BUILD_DATE"
cat >"$PACKAGE/DEBIAN/control" <<EOF
Package: pytools
Version: $DEB_VERSION
Section: utils
Priority: optional
Architecture: $ARCH
Depends: libc6 (>= $MAX_GLIBC), libstdc++6, libgcc-s1
Suggests: libreoffice-calc, libreoffice-writer
Maintainer: pytools local build <noreply@localhost>
Description: Excel and WPS report batch processing tools
 Interactive report processing with pandas and openpyxl engines.
EOF

BUILD_GLIBC=$(getconf GNU_LIBC_VERSION 2>/dev/null | sed 's/^glibc[[:space:]]*//' || true)
cat >"$PACKAGE/usr/share/doc/pytools/build-info.txt" <<EOF
Product version: $APP_VERSION
Debian version: $DEB_VERSION
Architecture: $ARCH
Python runtime: $PYTHON_VERSION (python-build-standalone, $STANDALONE_TAG)
PyInstaller: $PYINSTALLER_VERSION
Build machine glibc: ${BUILD_GLIBC:-unknown}
Highest required GLIBC symbol: $MAX_GLIBC (target <= $GLIBC_TARGET)
Highest required GLIBCXX symbol: ${MAX_GLIBCXX:-none} (target <= $GLIBCXX_TARGET; system libstdc++ required)
Excel/Word conversion engine in Office mode on Linux: LibreOffice (optional system dependency)
EOF

mkdir -p "$DIST"
ARTIFACT="pytools_${DEB_VERSION}_glibc${MAX_GLIBC}_${ARCH}.deb"
DEB_OUT=$DIST/$ARTIFACT
dpkg-deb --build --root-owner-group "$PACKAGE" "$DEB_OUT" >/dev/null
dpkg-deb --info "$DEB_OUT" >/dev/null
(
    cd -- "$DIST"
    sha256sum "$ARTIFACT" >"$ARTIFACT.sha256"
)
echo "[完成] $DEB_OUT"
echo "[校验] $(cat "$DEB_OUT.sha256")"
