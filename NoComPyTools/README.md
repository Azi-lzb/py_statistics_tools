# NoComPyTools

跨平台精简 Python 版项目。Windows 可直接运行 Python 源码，Linux 提供 deb 包。工作簿业务处理使用 pandas、openpyxl 和 python-calamine。格式转换按系统选择：Windows 使用 Excel/WPS 表格转换 xls、使用 Word/WPS 文字转换 doc/docx；Linux 使用 LibreOffice Calc/Writer。转换后，Excel 工具继续读写 xlsx。

python-calamine 是 Rust Excel 读取器的 Python 绑定，Linux amd64 manylinux 轮子要求 glibc 2.17。它用于读取 xlsx/xlsm/xls/xlsb 的值和合并区域；工作簿写入和样式修改继续由 openpyxl 完成。需要修改的工作簿输入限定为 xlsx/xlsm。

安装依赖：`python -m pip install -r requirements.txt`。启动开发版可使用 `bash start.sh`（Linux）或双击/运行 `启动WinExe.bat`（Windows）。

`打包WinExe.bat` 使用本机 Windows Python 和 PyInstaller 打包 Windows 程序，产物目录名包含版本号和本地秒级时间戳，并附带 `build-info.txt`；分发时请带上完整目录。`build.bat` 是它的兼容入口。Linux 上运行 `bash build.sh` 可生成 deb，产物位于 `dist/nocompytools/`，包版本和文件名包含版本号及 UTC 秒级时间戳；从 Windows 构建 deb 时运行 `打包LinuxDeb.bat`，它通过 WSL 调用原有 `build.sh`。两个 bat 分别构建 Windows exe 和 Linux deb，不会互相覆盖产物。普通功能无需 LibreOffice。4-2 的 xls 转 xlsx 和批量 Word 格式转换会按系统选择引擎：Windows 使用 Microsoft Excel/WPS 表格处理 xls，使用 Microsoft Word/WPS 文字处理 doc/docx；需要安装相应 Office 组件，Windows 条件依赖由 `启动WinExe.bat` 安装 pywin32。Linux（包括麒麟、统信、Ubuntu、Debian）使用 LibreOffice Calc 转换 xls，使用 LibreOffice Writer 转换 doc/docx。

Windows 程序使用 PyInstaller onefile 布局。每个版本目录包含一个 exe 和同级的 `config/config.xlsx`；配置可直接编辑，更新程序时替换 exe 并保留 `config` 目录即可。分发时请携带整个版本目录。

打印的格式保留模式由 openpyxl 复制边框、合并单元格等样式；快速复制模式使用 python-calamine 读取值，再由 openpyxl 写出格式并保留边框和合并结构。工作表提取和两种汇总均由 python-calamine 读取值，再用 openpyxl 写入新文件；工作表提取整表时保留行列位置，不复制源格式、合并单元格或行列尺寸。只读提取和汇总可直接读取 xls/xlsb。输出工作表沿用源 Sheet 名（重名时自动加序号）。

开发测试：安装 `requirements-dev.txt` 后运行 `pytest -q`。在安装了 LibreOffice Calc 的虚拟机中，`test_libreoffice_converts_real_xls_values_and_merged_headers` 会检查真实 BIFF xls 的值和合并表头。安装 deb 后可运行 `nocompytools` 打开菜单，用业务 xls 执行转换，并用输出的 xlsx 完成后续读取任务。
