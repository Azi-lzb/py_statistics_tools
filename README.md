# pytools（通用汇总工具）

面向 Excel 场景的批处理工具，提供时序提取、去重、打印、转换、汇总、归档和专项处理能力。常用 Excel 功能已改为自动路由：`.xlsx/.xlsm` 优先走普通处理逻辑，`.xls` 自动走 COM 逻辑。

## 下载与项目版本

从 [GitHub Releases](https://github.com/Azi-lzb/py_statistics_tools/releases) 下载程序和 SHA-256 校验文件。Windows 用户下载 `pytools_*.zip`，解压后运行 `pytools.exe`，保留同级 `config/config.xlsx`。

本次发布为母项目 **V1.1.0 laptop 稳定基线**，发布标签包含 `laptop` 和发布日期；日期表示重新打包日期。母项目面向 Windows，自动路由功能的 `.xls` 使用 COM，其余格式使用 Python；明确标注 COM 的功能仍使用 COM。COM 功能需要安装相应的 Excel/WPS 或 Word/WPS 组件。当前 Windows 包使用 Python 3.12 运行时，包含 Python，无需另装解释器；未验收 Windows 7。

[NoComPyTools 子项目](NoComPyTools/README.md) 独立维护，面向 Windows 和 Linux：普通表格处理使用 Python、python-calamine 和 openpyxl；xls→xlsx 与 doc/docx 转换在 Windows 使用对应的 Office/WPS，在 Linux 使用 LibreOffice Calc/Writer。其 Linux deb、Windows 包与母项目分别构建和验收。新建纯 Python 功能默认在子项目开发。此次 Release 仅提供母项目 Windows 包，不将历史子项目产物标为本次构建。

源码的稳定恢复提交位于 `desktop` 分支；GitHub 首页 `main` 的本次更新只同步 README。需要对应发布源码时，请下载 Release 标签下的 Source code。

更新时先备份用户配置，把新包解压到新目录，再复制原来的 `config/config.xlsx`；不要用压缩包内的默认配置覆盖已维护的配置。日志和输出目录按配置表设置。

## 目录结构

- `PyCode/`：母项目核心源码
- `pytools_entry.py`：程序入口
- `PyCode/config/config.xlsx`：源码默认配置文件
- `NoComPyTools/`：独立的跨平台子项目
- `dist/`：打包后的可执行程序与运行目录
- `build_exe.bat`：一键打包脚本（PyInstaller）

## 运行方式

### 1) 直接运行源码

```bash
python pytools_entry.py
```

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

默认使用 PyInstaller 生成单文件程序（`--onefile`）。

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
- 批量 Excel 格式转换（COM）
- 批量 Word 格式转换（COM）
- 批量修改 Sheet 名（自动路由 COM）
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

- 源码主配置文件：`PyCode/config/config.xlsx`；发布包配置位于 exe 同级的 `config/config.xlsx`
- 输入/输出/日志目录由配置表控制
- 程序会按功能自动生成日志，便于追溯

## 注意事项

- COM 功能依赖本机可用的 Office/Excel COM 环境。
- 涉及文件选择时，程序会弹出选择窗口；若弹窗不可用，可手动输入路径。
- 首次使用建议先执行“配置工具 -> 初始化/修复配置表”。
