# 版本管理

参考 Open-Data-Audit 的版本管理体系建立的精简版。

## 组成

- `VERSION_MANIFEST.json` — 版本唯一来源：产品名 + 当前版本号 + 版本号规则。
- `CHANGELOG.md` — 每个发布版本的变更记录。
- `PyCode/core/app_identity.py` — 统一读取模块，源码和打包后的 exe 都从清单读版本；菜单横幅显示 `pytools Vx.y.z`，构建脚本用它命名产物。

## 发新版本的流程

1. 改代码，跑 `pytest` 确认测试通过。
2. 在 `VERSION_MANIFEST.json` 里把 `version` 改成新版本号：
   - 修复/优化不改功能 → PATCH（1.0.0 → 1.0.1）
   - 新增功能 → MINOR（1.0.1 → 1.1.0）
   - 不兼容改动 → MAJOR（1.1.0 → 2.0.0）
3. 在 `CHANGELOG.md` 顶部加一节新版本说明。
4. 跑 `build_exe.bat`，产物为 `dist\pytools V{版本}_{时间戳}.exe`（同时保留 `dist\pytools.exe` 作为最新版副本）。
5. 提交到 `laptop` 分支并推送，稳定后同步 `main`。
