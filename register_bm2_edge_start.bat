@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

reg import "%~dp0register_bm2_edge_start.reg"
if errorlevel 1 (
    echo Register failed.
    pause
    exit /b 1
)

echo BM2 Edge link registered.
echo Open this in Edge:
echo bm2://start
echo.
pause
