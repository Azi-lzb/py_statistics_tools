@echo off
setlocal
cd /d "%~dp0"

set "PY_EXE_HZPBC=%UserProfile%\miniconda3\envs\HZPBC\python.exe"
set "PY_EXE=%LocalAppData%\Programs\Python\Python310\python.exe"
set "PY_CMD="

if exist "%PY_EXE_HZPBC%" (
  set "PY_CMD=%PY_EXE_HZPBC%"
) else if exist "%PY_EXE%" (
  set "PY_CMD=%PY_EXE%"
) else (
  where python >nul 2>nul
  if not errorlevel 1 (
    set "PY_CMD=python"
  )
)

rem Read product version (VERSION_MANIFEST.json is the single source of truth)
set "VERSION="
if defined PY_CMD (
  for /f "usebackq delims=" %%i in (`"%PY_CMD%" -m PyCode.core.app_identity --version 2^>nul`) do set "VERSION=%%i"
) else (
  for /f "usebackq delims=" %%i in (`py -3 -m PyCode.core.app_identity --version 2^>nul`) do set "VERSION=%%i"
)
if not defined VERSION (
  echo [WARN] failed to read version, timestamped copy keeps legacy naming
)

rem Keep the previous exe until the new build succeeds
if exist "dist\pytools.exe" move /y "dist\pytools.exe" "dist\pytools_prev.tmp" >nul

if defined PY_CMD (
  "%PY_CMD%" -m PyInstaller --clean --noconfirm pytools.spec
) else (
  py -3 -m PyInstaller --clean --noconfirm pytools.spec
)

if errorlevel 1 (
  if exist "dist\pytools_prev.tmp" move /y "dist\pytools_prev.tmp" "dist\pytools.exe" >nul
  echo [ERROR] PyInstaller build failed
  exit /b 1
)
if exist "dist\pytools_prev.tmp" del /q "dist\pytools_prev.tmp" >nul

if not exist "dist\pytools.exe" (
  echo [WARN] dist\pytools.exe not found
  exit /b 1
)

set "TS_FILE=%TEMP%\pytools_build_ts.txt"
set "TS="
powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss" > "%TS_FILE%"
if exist "%TS_FILE%" (
  set /p TS=<"%TS_FILE%"
  del /q "%TS_FILE%" >nul 2>&1
)

if defined TS if defined VERSION (
  copy /y "dist\pytools.exe" "dist\pytools V%VERSION%_%TS%.exe" >nul
  echo [OK] built dist\pytools.exe
  echo [OK] built dist\pytools V%VERSION%_%TS%.exe
)
if defined TS if not defined VERSION (
  copy /y "dist\pytools.exe" "dist\pytools_%TS%.exe" >nul
  echo [OK] built dist\pytools.exe
  echo [OK] built dist\pytools_%TS%.exe
)
if not defined TS (
  echo [OK] built dist\pytools.exe
  echo [WARN] timestamp suffix copy skipped
)

exit /b 0
