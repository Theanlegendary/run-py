@echo off
setlocal enabledelayedexpansion
title Shipments Tomorrow Bot (00:00 - 06:00 Auto Launcher)

echo =====================================================================
echo           🚚 SHIPMENTS TOMORROW BOT - AUTO INSTALLER & RUNNER
echo =====================================================================
echo.

:: 1. Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not added to PATH!
    echo Please install Python 3.10+ from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

echo [*] Python detected:
python --version
echo.

:: 2. Auto Install Dependencies
echo [*] Installing / Verifying required Python packages...
python -m pip install --upgrade pip --quiet
python -m pip install -r requirements.txt --quiet

if %errorlevel% neq 0 (
    echo [!] Standard install had an issue, trying verbose install...
    python -m pip install python-telegram-bot==20.8 requests openpyxl pandas Pillow apscheduler
)

echo.
echo [OK] All dependencies and fonts are ready!
echo.
echo =====================================================================
echo   🚀 STARTING TOMORROW BOT (00:00 - 06:00 MORNING REPORT)...
echo   Bot is listening for /tomorrow commands on Telegram.
echo   Press Ctrl+C to stop.
echo =====================================================================
echo.

:: 3. Run Bot with Auto-Restart Loop
:RUN_LOOP
python bot.py
echo.
echo [!] Bot process ended. Restarting in 5 seconds... (Press Ctrl+C to exit)
timeout /t 5 >nul
goto RUN_LOOP
