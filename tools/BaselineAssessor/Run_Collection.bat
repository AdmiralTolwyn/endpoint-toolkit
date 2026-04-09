@echo off
title BaselinePilot - Data Collection
cd /d "%~dp0"
echo ============================================
echo   BaselinePilot - Data Collection
echo   Requires local administrator privileges
echo ============================================
echo.
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: This script must be run as Administrator.
    echo Right-click and select "Run as administrator".
    echo.
    pause
    exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Invoke-BaselineCollection.ps1"
echo.
pause
