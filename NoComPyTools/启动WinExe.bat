@echo off
setlocal EnableExtensions
pushd "%~dp0"

set "VENV_PY=%~dp0.venv\Scripts\python.exe"
if exist "%VENV_PY%" goto check_dependencies

where py >nul 2>&1
if errorlevel 1 goto use_python
echo Creating the NoComPyTools development environment...
py -3 -m venv "%~dp0.venv"
if errorlevel 1 goto venv_failed
goto check_dependencies

:use_python
where python >nul 2>&1
if errorlevel 1 goto missing_python
echo Creating the NoComPyTools development environment...
python -m venv "%~dp0.venv"
if errorlevel 1 goto venv_failed

:check_dependencies
"%VENV_PY%" -c "import pandas, openpyxl, python_calamine, win32com.client" >nul 2>&1
if not errorlevel 1 goto launch

echo Installing NoComPyTools dependencies. This may take a few minutes...
"%VENV_PY%" -m pip install --disable-pip-version-check -r "%~dp0requirements.txt"
if errorlevel 1 goto dependencies_failed

:launch
"%VENV_PY%" nocompytools_entry.py %*
set "EXIT_CODE=%ERRORLEVEL%"
if "%EXIT_CODE%"=="0" goto success
echo.
echo NoComPyTools exited with code %EXIT_CODE%.
goto pause_and_exit

:success
popd
exit /b 0

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

:pause_and_exit
echo.
pause
popd
exit /b %EXIT_CODE%