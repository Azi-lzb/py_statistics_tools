@echo off
setlocal
cd /d "%~dp0"

set "PY_EXE_PROJECT=%~dp0.venv\Scripts\python.exe"
set "PY_EXE_HZPBC=%UserProfile%\miniconda3\envs\HZPBC\python.exe"
set "PY_EXE_PY310=%LocalAppData%\Programs\Python\Python310\python.exe"
set "PY_EXE="
set "PY_ARGS="
set "FOUND_PYTHON="

if exist "%PY_EXE_PROJECT%" set "PY_EXE=%PY_EXE_PROJECT%"
if not defined PY_EXE if exist "%PY_EXE_HZPBC%" set "PY_EXE=%PY_EXE_HZPBC%"
if defined PY_EXE (
    set "FOUND_PYTHON=1"
    call :check_dependencies
    if not errorlevel 1 goto python_ready
    set "PY_EXE="
)

if exist "%PY_EXE_PY310%" set "PY_EXE=%PY_EXE_PY310%"
if defined PY_EXE (
    set "FOUND_PYTHON=1"
    call :check_dependencies
    if not errorlevel 1 goto python_ready
    set "PY_EXE="
)

where python >nul 2>nul
if not errorlevel 1 set "PY_EXE=python"
if defined PY_EXE (
    set "FOUND_PYTHON=1"
    call :check_dependencies
    if not errorlevel 1 goto python_ready
    set "PY_EXE="
)

where py >nul 2>nul
if not errorlevel 1 (
    set "PY_EXE=py"
    set "PY_ARGS=-3"
)
if defined PY_EXE (
    set "FOUND_PYTHON=1"
    call :check_dependencies
    if not errorlevel 1 goto python_ready
)

if defined FOUND_PYTHON goto missing_dependencies
goto missing_python

:python_ready
set "VERSION="
for /f "usebackq delims=" %%V in (`"%PY_EXE%" %PY_ARGS% -m PyCode.core.app_identity --version 2^>nul`) do set "VERSION=%%V"
if defined VERSION echo [DEV] Starting pytools V%VERSION% from source...
if not defined VERSION echo [DEV] Starting pytools from source...
"%PY_EXE%" %PY_ARGS% pytools_entry.py %*
set "EXIT_CODE=%ERRORLEVEL%"
if "%EXIT_CODE%"=="0" exit /b 0

echo.
echo Development run failed with exit code %EXIT_CODE%.
pause
exit /b %EXIT_CODE%

:check_dependencies
"%PY_EXE%" %PY_ARGS% -c "import sys, numpy, pandas, openpyxl, xlrd, pyxlsb; __import__('win32com.client') if sys.platform == 'win32' else None" >nul 2>nul
exit /b %ERRORLEVEL%

:missing_dependencies
echo [ERROR] Python was found, but no environment has all required packages: numpy, pandas, openpyxl, xlrd, pyxlsb.
echo Windows Office mode also requires pywin32.
echo Install the parent project's runtime dependencies, then run this launcher again.
pause
exit /b 1

:missing_python
echo [ERROR] Python 3 was not found. Install Python 3 and the parent project's runtime dependencies first.
pause
exit /b 1
