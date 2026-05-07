@echo off
:: Device Decommissioner launcher
cd /d "%~dp0"

set "PS_EXE=powershell.exe"
where pwsh >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    set "PS_EXE=pwsh"
)

%PS_EXE% -NoProfile -Command "Get-ChildItem -LiteralPath '%~dp0' -Recurse | Unblock-File"

start "" %PS_EXE% -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "DeviceDecommissioner.ps1"
exit
