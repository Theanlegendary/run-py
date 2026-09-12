@echo off
cd /d "%~dp0"
echo ========================================
echo  SAFE RESTART BOT (no duplicate check)
echo ========================================
echo.

echo [1/3] Killing ALL running bot.py instances...
taskkill /F /IM python.exe /T >nul 2>&1
timeout /t 3 >nul

echo [2/3] Verifying no duplicates remain...
tasklist /FI "IMAGENAME eq python.exe" | findstr python.exe >nul
if %ERRORLEVEL% EQU 0 (
    echo [!] Still found Python processes. Force killing again...
    taskkill /F /IM python.exe /T >nul 2>&1
    timeout /t 2 >nul
) else (
    echo [OK] No Python processes running. Safe to start.
)

echo [3/3] Starting bot.py fresh...
echo.
python bot.py

echo.
echo Bot stopped.
pause
