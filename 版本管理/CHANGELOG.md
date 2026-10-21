# pytools 更新日志

所有对外可见的变更记录在此文件。版本号遵循 `版本管理/VERSION_MANIFEST.json` 中的 scheme。

## V1.2.0 — 2026-10-21

新增功能（MINOR）。

- 新增 Office / 非 Office 运行模式：Windows Office 模式使用 Excel/WPS COM，Linux Office 模式用 LibreOffice 执行 Excel/Word 格式转换；非 Office 模式只使用 Python 库。
- 新增 Linux amd64/arm64 `.deb` 打包入口，使用独立 CPython 运行时，并在构建时审计 GLIBC/GLIBCXX 符号版本；麒麟 V10 SP1 目标基线为 GLIBC 2.28 / GLIBCXX 3.4.24。
- Linux 用户配置存放于 XDG 配置目录；非 Office Excel 格式转换支持导出 `.xlsx` 和 `.csv`。

## V1.1.0 — 2026-10-12

新增功能（MINOR）。

- 宽表汇总新增"是否输出行序号"开关：`时序提取规则` 表新增 S 列 `是否输出行序号`（默认空/否，不输出，历史配置输出表头不变）。开启后宽表输出增加 `行序号` 列（源表行号），用于行头重复的报表（如两行同名"其他情况导致"）区分行身份；配合 `目标去重列 = 数据日期;行序号` 可让目标写入正确保留同名行。非 COM 与 COM 引擎行为一致。
- 新增测试 `test_wide_row_no_output.py` 覆盖开关解析与输出列行为。

## V1.0.0 — 2026-10-11

版本管理体系建立的基线版本。

- 全部既有功能入库：时序工具、去重工具、打印工具、转换工具、地区总分校验、机构标准化、外资行剔除、外币表头、区域汇总核对、工作表提取、农贷调整（月度/季度）、问卷统计等。
- 源码包由 `pytools/` 更名为 `PyCode/`，包引用同步更新。
- 清理仓库中的构建产物、日志与 `__pycache__`，新增 `.gitignore`。
- 移除已废弃的"校验规则生成"功能遗留：`校验规则映射表` 配置 Sheet 与相关死代码，配置修复时自动清除该历史 Sheet。
- 打包产物命名带版本与时间戳：`pytools V{版本}_{时间戳}.exe`。
