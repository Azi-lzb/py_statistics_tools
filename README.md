# pytools（通用汇总工具）

面向 Excel 场景的批处理工具，提供时序提取、去重、打印、转换、汇总、归档和专项处理能力。默认使用 Office 模式；可在 `4` → `11` 切换到非 Office 模式。

## 目录结构

- `pytools/`：核心源码
- `pytools_entry.py`：程序入口
- `pytools/config/config.xlsx`：配置文件
- `dist/`：打包后的可执行程序与运行目录
- `build_exe.bat`：一键打包脚本（PyInstaller）

## 运行方式

### 1) 直接运行源码

```bash
python pytools_entry.py
```

Windows 上也可双击根目录的 `启动开发版.bat`，它会选择已安装运行依赖的 Python 环境并直接运行源码，不启动 `dist/` 中的打包程序。

Windows 开发环境建议使用 Python 3.10 或更新版本，在项目根目录运行：

```bat
py -3 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
```

`requirements.txt` 是跨平台运行依赖；`requirements-dev.txt` 另含测试和 PyInstaller 打包工具。Windows Office 模式通过 `pywin32` 调用 Excel/WPS COM。

运行后按菜单选择功能：

- `1` 时序工具
- `2` 去重工具
- `3` 打印工具
- `4` 转换工具
- `5` 配置工具
- `6` 汇总工具
- `7` 一二批处理工具
- `8` 归档工具
- `9` 快捷工具

### 2) 运行打包程序

进入 `dist/` 目录，直接双击或命令行运行 `.exe` 文件。

## 打包

```bash
build_exe.bat
```

默认使用 PyInstaller 生成单文件程序（`--onefile`）。可编辑的配置文件单独放在 `dist/config/config.xlsx`；分发时请将 exe 与 `config` 目录一起提供。打包脚本会在该配置缺失时复制默认配置，已有配置不会被覆盖。

### Linux / 银河麒麟 V10 SP1

Linux `.deb` 按构建机架构生成（amd64 或 arm64）。为兼容麒麟 V10 SP1，冻结文件审计上限设为 GLIBC 2.28 / GLIBCXX 3.4.24：

```bash
bash packaging/linux/build_deb.sh
```

脚本在 `dist/linux/` 生成 `.deb` 与 SHA-256 文件。安装时使用：

```bash
sudo apt install ./dist/linux/pytools_*.deb
pytools
```

Linux 首次启动会将默认配置复制到 `~/.config/pytools/config/config.xlsx`。程序默认使用 Office 模式：

- **Office 模式**：Windows 的表格操作按功能使用 Excel/WPS COM；Linux 的 Excel/Word 格式转换使用 LibreOffice Calc/Writer（安装 `libreoffice-calc`、`libreoffice-writer`）。Linux 上没有对应 LibreOffice 实现的操作会使用现有 Python 实现。
- **非 Office 模式**：只使用 pandas/openpyxl/xlrd/pyxlsb 等 Python 库，不调用 COM 或 LibreOffice。Excel 转换支持导出 `.xlsx` 和 `.csv`；Word 格式转换不在此模式提供。

运行中可在 `4` → `11` 切换到非 Office 模式；该切换仅对当前运行有效，重新启动后默认回到 Office 模式。

Linux 上 LibreOffice 是可选系统依赖；仅 Office 模式的 Excel/Word 格式转换需要它。Python 模式转换 `.xlsm` 为 `.xlsx` 会丢弃 VBA 宏；导出 CSV 只取第一个工作表，且不保留单元格格式。

## 功能说明

### 1. 时序工具

- 表头路径比对（自动路由）
- 时序提取（自动路由）
- 宽表汇总（行头加后缀 / 不加后缀，自动路由）
- 预校验

### 2. 去重工具

- 按批注检查重复 / 删除重复
- 按配置检查重复 / 删除重复
- 按配置去重并追加到目标
- 按配置预校验

### 3. 打印工具

- 按批注打印（保留源格式 / 快速复制，自动路由）
- 按配置打印（执行全部模式，自动路由）
- 按配置预校验
- 按配置打印并导出 PDF（打印预览，自动路由 COM）

### 4. 转换工具

- 批量重命名文件
- 批量 Excel 格式转换（按所选模式）
- 批量 Word 格式转换（Office 模式）
- 批量修改 Sheet 名（按所选模式）
- 根据文件内容重命名 / 取消内容前缀
- 批量局部映射重命名文件

### 5. 配置工具

- 初始化/修复 `config.xlsx`
- 查看当前配置并重新读取

### 6. 汇总工具

- 按使用区域汇总（全量，自动路由）
- 按批注汇总（模板 + 源文件，自动路由）

### 7. 一二批处理工具

- 拆分村镇银行数据
- 机构名称标准化
- 删除分机构表外资行
- 外汇页眉修改（普通 / COM）
- 地区总分校验
- 提取工作表数据（普通 / COM）
- 涉农贷款“比上月”修正
- 季度涉农贷款“比上月”修正及分地区核对

### 8. 归档工具

- 按日期归档（预览 / 直接复制）
- 按类型归档（预览 / 直接复制）
- 执行归档计划（复制）
- 文件夹选择会递归遍历；未识别日期/类型会进入兜底目录

### 9. 快捷工具

- 打开配置文件
- 打开 output 最新文件
- 打开 output 文件夹

## 配置与输出

- 主配置文件：`pytools/config/config.xlsx`
- 输入/输出/日志目录由配置表控制
- 程序会按功能自动生成日志，便于追溯

## 注意事项

- COM 功能依赖本机可用的 Office/Excel COM 环境。
- 涉及文件选择时，程序会弹出选择窗口；若弹窗不可用，可手动输入路径。
- 首次使用建议先执行“配置工具 -> 初始化/修复配置表”。
