@echo off
:: ---------------------------------------------------------------------
:: DumpPilot Launcher
:: Requires PowerShell 7 (pwsh.exe) on PATH. STA needed for WPF.
:: ---------------------------------------------------------------------
title DumpPilot
cd /d "%~dp0"

where pwsh.exe >NUL 2>&1
if errorlevel 1 (
    echo PowerShell 7 pwsh.exe was not found on PATH.
    echo Install it from https://github.com/PowerShell/PowerShell/releases
    pause
    exit /b 1
)

:: Unblock all files so dot-sourced helpers and XAML load without issues
pwsh.exe -NoProfile -Command "Get-ChildItem -LiteralPath '%~dp0' -Recurse | Unblock-File"

:: Launch the WPF GUI - STA required, hide the console
start "" /MAX pwsh.exe -NoProfile -ExecutionPolicy Bypass -Sta -WindowStyle Hidden -File "DumpPilot.ps1"

exit
