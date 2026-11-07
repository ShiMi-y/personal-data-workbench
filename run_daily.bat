@echo off
chcp 65001 >nul
cd /d "%~dp0"
py run_daily.py %*
echo.
pause

