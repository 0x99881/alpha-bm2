@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set "APP_URL=http://127.0.0.1:5000"
set "PY_CMD="

where py >nul 2>nul
if not errorlevel 1 (
    set "PY_CMD=py -3"
) else (
    where python >nul 2>nul
    if not errorlevel 1 (
        set "PY_CMD=python"
    )
)

if not defined PY_CMD (
    echo Python was not found. Please install Python first.
    pause
    exit /b 1
)

if /i "%~1"=="--check" (
    echo Start check passed
    echo Current folder: %cd%
    echo Python command: %PY_CMD%
    echo URL: %APP_URL%
    exit /b 0
)

for /f "tokens=5" %%p in ('netstat -ano ^| findstr /r /c:":5000 .*LISTENING"') do (
    echo Stopping old local server on port 5000...
    taskkill /PID %%p /F >nul 2>nul
    timeout /t 1 /nobreak >nul
)

start "" powershell -NoProfile -WindowStyle Hidden -Command "Start-Sleep -Seconds 2; Start-Process '%APP_URL%'"
%PY_CMD% app.py

if errorlevel 1 (
    echo.
    echo Start failed. Please read the message above.
    pause
)
