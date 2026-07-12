@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

if /i "%~1"=="--check" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0launcher\start_or_open.ps1" -Check
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0launcher\start_or_open.ps1" %*
)
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
    echo.
    echo Start failed. Please read the message above.
    pause
)

exit /b %EXIT_CODE%
