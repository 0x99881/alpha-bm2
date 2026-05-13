@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set "APP_URL=http://127.0.0.1:5000/scores"
set "PY_CMD="
set "CHROME_EXE="

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

if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" (
    set "CHROME_EXE=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
) else if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" (
    set "CHROME_EXE=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
) else if exist "%LocalAppData%\Google\Chrome\Application\chrome.exe" (
    set "CHROME_EXE=%LocalAppData%\Google\Chrome\Application\chrome.exe"
) else (
    for /f "delims=" %%c in ('where chrome 2^>nul') do (
        if not defined CHROME_EXE set "CHROME_EXE=%%c"
    )
)

if /i "%~1"=="--check" (
    echo Browser start check passed
    echo Current folder: %cd%
    echo Python command: %PY_CMD%
    echo Chrome path: %CHROME_EXE%
    echo URL: %APP_URL%
    exit /b 0
)

netstat -ano | findstr /r /c:":5000 .*LISTENING" >nul 2>nul
if errorlevel 1 (
    start "BM2 Local Server" /min cmd /c "cd /d ""%~dp0"" && %PY_CMD% app.py"
    timeout /t 2 /nobreak >nul
)

if defined CHROME_EXE (
    start "" "%CHROME_EXE%" "%APP_URL%"
) else (
    start "" "%APP_URL%"
)
