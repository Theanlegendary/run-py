@echo off
setlocal EnableDelayedExpansion
title Metfone Bot - Dependency Installer

echo ============================================================
echo   METFONE EXPRESS BOT - DEPENDENCY INSTALLER
echo ============================================================
echo.

:: 1. Detect Python
echo [1/4] Checking Python installation...
set "PY_CMD="

where python >nul 2>&1
if not errorlevel 1 (
    for /f "tokens=*" %%i in ('where python') do (
        if not defined PY_CMD set "PY_CMD=%%i"
    )
)

if not defined PY_CMD (
    where py >nul 2>&1
    if not errorlevel 1 (
        set "PY_CMD=py -3"
    )
)

if not defined PY_CMD (
    for /d %%p in ("%LocalAppData%\Programs\Python\Python3*") do (
        if exist "%%p\python.exe" set "PY_CMD=%%p\python.exe"
    )
)

if not defined PY_CMD (
    for /d %%p in ("C:\Program Files\Python3*") do (
        if exist "%%p\python.exe" set "PY_CMD=%%p\python.exe"
    )
)

if not defined PY_CMD (
    for /d %%p in ("C:\Python3*") do (
        if exist "%%p\python.exe" set "PY_CMD=%%p\python.exe"
    )
)

if not defined PY_CMD (
    echo.
    echo [ERROR] Python was not found on this computer!
    echo.
    echo Would you like to automatically download and install Python 3.11 now?
    set /p "INSTALL_CHOICE=Install Python now? (Y/N): "
    if /i "!INSTALL_CHOICE!"=="Y" (
        call "%~dp0INSTALL_PYTHON_3_11.bat"
        echo Please restart this installer after Python installation completes.
        pause
        exit /b 0
    ) else (
        echo Please install Python 3.10+ and re-run this script.
        pause
        exit /b 1
    )
)

echo [OK] Using Python: %PY_CMD%
%PY_CMD% --version
echo.

:: 2. Upgrade pip, setuptools, wheel
echo [2/4] Ensuring pip, wheel, and setuptools are up to date...
%PY_CMD% -m pip install --upgrade pip setuptools wheel --no-warn-script-location
echo.

:: 3. Install requirements
echo [3/4] Installing required packages from requirements.txt...
cd /d "%~dp0"
if exist requirements.txt (
    %PY_CMD% -m pip install -r requirements.txt --no-warn-script-location
) else (
    echo [ERROR] requirements.txt not found in %~dp0
    pause
    exit /b 1
)
echo.

:: 4. Verify imports
echo [4/4] Verifying all installed modules...
%PY_CMD% -c "
import sys
modules = ['telegram', 'requests', 'openpyxl', 'pandas', 'PIL', 'apscheduler', 'tabulate', 'flask', 'win32com.client', 'fitz']
failed = []
for m in modules:
    try:
        __import__(m)
        print(f'  [OK] {m}')
    except ImportError as e:
        failed.append(m)
        print(f'  [FAIL] {m}: {e}')

if failed:
    sys.exit(1)
"

if errorlevel 1 (
    echo.
    echo [WARNING] Some modules failed to load.
    echo Please check the error messages above.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   [SUCCESS] All dependencies are successfully installed!
echo   You can now run 'restart_bot.bat' to start the bot.
echo ============================================================
echo.
pause
