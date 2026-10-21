@echo off
chcp 936 >nul
setlocal EnableExtensions
pushd "%~dp0"

set "VENV_PY=%~dp0.venv\Scripts\python.exe"
if exist "%VENV_PY%" goto check_dependencies

where py >nul 2>&1
if errorlevel 1 goto use_python
echo Creating the NoComPyTools Windows build environment...
py -3 -m venv "%~dp0.venv"
if errorlevel 1 goto venv_failed
goto check_dependencies

:use_python
where python >nul 2>&1
if errorlevel 1 goto missing_python
echo Creating the NoComPyTools Windows build environment...
python -m venv "%~dp0.venv"
if errorlevel 1 goto venv_failed

:check_dependencies
"%VENV_PY%" -c "import pandas, openpyxl, python_calamine, win32com.client, PyInstaller" >nul 2>&1
if not errorlevel 1 goto read_version

echo Installing NoComPyTools build dependencies. This may take a few minutes...
"%VENV_PY%" -m pip install --disable-pip-version-check -r "%~dp0requirements.txt"
if errorlevel 1 goto dependencies_failed

:read_version
set "APP_VERSION="
for /f "usebackq delims=" %%V in (`"%VENV_PY%" -m nocompytools.core.app_identity --version 2^>nul`) do set "APP_VERSION=%%V"
if not defined APP_VERSION goto version_failed
set "BUILD_TIMESTAMP="
for /f %%T in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss" 2^>nul') do set "BUILD_TIMESTAMP=%%T"
if not defined BUILD_TIMESTAMP goto timestamp_failed
set "BUILD_NAME=NoComPyTools_V%APP_VERSION%_%BUILD_TIMESTAMP%"

echo [Version] NoComPyTools V%APP_VERSION%
echo [Build time] %BUILD_TIMESTAMP% local
echo Building the Windows application with native Python...
"%VENV_PY%" -m PyInstaller --clean --noconfirm --onefile --name "%BUILD_NAME%" --distpath "%~dp0dist\windows\%BUILD_NAME%" --workpath "%~dp0build\windows\work" --specpath "%~dp0build\windows\spec" --add-data "%~dp0版本管理\VERSION_MANIFEST.json;版本管理" --hidden-import python_calamine --hidden-import pandas.io.excel._calamine nocompytools_entry.py
if errorlevel 1 goto build_failed

if not exist "%~dp0dist\windows\%BUILD_NAME%\config" mkdir "%~dp0dist\windows\%BUILD_NAME%\config"
copy /y "%~dp0nocompytools\config\config.xlsx" "%~dp0dist\windows\%BUILD_NAME%\config\config.xlsx" >nul
if errorlevel 1 goto config_copy_failed

(
echo Product: NoComPyTools
echo Version: %APP_VERSION%
echo Build timestamp local: %BUILD_TIMESTAMP%
) > "%~dp0dist\windows\%BUILD_NAME%\build-info.txt"
echo.
echo Windows build completed: %~dp0dist\windows\%BUILD_NAME%\%BUILD_NAME%.exe
set "EXIT_CODE=0"
goto pause_and_exit

:missing_python
echo Python 3 was not found. Install Python 3, then run this file again.
set "EXIT_CODE=1"
goto pause_and_exit

:venv_failed
echo Could not create the local Python environment.
set "EXIT_CODE=1"
goto pause_and_exit

:dependencies_failed
echo Dependency installation failed. Check the network and the error above, then run this file again.
set "EXIT_CODE=1"
goto pause_and_exit

:version_failed
echo Could not read the product version from 版本管理\VERSION_MANIFEST.json.
set "EXIT_CODE=1"
goto pause_and_exit

:timestamp_failed
echo Could not generate the build timestamp.
set "EXIT_CODE=1"
goto pause_and_exit

:build_failed
echo Windows EXE build failed with exit code %ERRORLEVEL%. See the error above.
set "EXIT_CODE=1"
goto pause_and_exit

:config_copy_failed
echo Could not copy config.xlsx into the Windows package directory.
set "EXIT_CODE=1"

:pause_and_exit
echo.
pause
popd
exit /b %EXIT_CODE%