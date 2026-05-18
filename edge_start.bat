@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
cd /d "%~dp0"

set "APP_URL=http://127.0.0.1:5000/scores"
set "PY_CMD="
set "EDGE_EXE="

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

if exist "%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe" (
    set "EDGE_EXE=%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"
) else if exist "%ProgramFiles%\Microsoft\Edge\Application\msedge.exe" (
    set "EDGE_EXE=%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"
) else if exist "%LocalAppData%\Microsoft\Edge\Application\msedge.exe" (
    set "EDGE_EXE=%LocalAppData%\Microsoft\Edge\Application\msedge.exe"
) else (
    for /f "delims=" %%e in ('where msedge 2^>nul') do (
        if not defined EDGE_EXE set "EDGE_EXE=%%e"
    )
)

if /i "%~1"=="--check" (
    echo Edge start check passed
    echo Current folder: %cd%
    echo Python command: !PY_CMD!
    echo Edge path: !EDGE_EXE!
    echo URL: %APP_URL%
    exit /b 0
)

netstat -ano | findstr /r /c:":5000 .*LISTENING" >nul 2>nul
if errorlevel 1 (
    start "BM2 Local Server" /min cmd /c "cd /d ""%~dp0"" && !PY_CMD! app.py"
    timeout /t 2 /nobreak >nul
)

if defined EDGE_EXE (
    start "" "!EDGE_EXE!" "%APP_URL%"
) else (
    start "" microsoft-edge:%APP_URL%
)
