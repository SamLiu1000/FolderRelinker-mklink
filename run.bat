@echo off
:: Check for admin privileges
net session >nul 2>&1
if %errorlevel% == 0 (
    echo Running with Administrator privileges...
    cd /d "%~dp0"
    python mklink_manager.py
) else (
    echo Requesting Administrator privileges...
    powershell -Command "Start-Process python -ArgumentList '\"%~dp0mklink_manager.py\"' -Verb RunAs"
)
pause
