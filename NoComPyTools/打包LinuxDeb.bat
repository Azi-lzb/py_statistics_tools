@echo off
setlocal

where wsl.exe >nul 2>&1
if errorlevel 1 goto missing_wsl

set "WSL_PROJECT_DIR="
set "WIN_PROJECT_DIR=%~dp0"
set "WIN_PROJECT_DIR=%WIN_PROJECT_DIR:\=/%"
for /f "usebackq delims=" %%P in (`wsl.exe --exec wslpath -a "%WIN_PROJECT_DIR%"`) do set "WSL_PROJECT_DIR=%%P"
if not defined WSL_PROJECT_DIR goto resolve_failed

echo Building the Linux deb package in WSL...
wsl.exe --cd "%WSL_PROJECT_DIR%" bash ./build.sh
set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" goto build_failed

echo.
echo Linux deb build completed. See the package path and checksum above.
goto pause_and_exit

:missing_wsl
echo WSL is required to build the Linux deb package. Install and configure WSL first.
set "EXIT_CODE=1"
goto pause_and_exit

:resolve_failed
echo Could not resolve the project directory inside WSL.
set "EXIT_CODE=1"
goto pause_and_exit

:build_failed
echo.
echo Linux deb build failed with exit code %EXIT_CODE%. See the error above.

:pause_and_exit
echo.
pause
exit /b %EXIT_CODE%
