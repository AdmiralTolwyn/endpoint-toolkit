@echo off
:: ---------------------------------------------------------------------
:: Windows 365 Assessor Launcher
:: Automatically detects PowerShell 7 (Core) and defaults to it.
:: Falls back to Windows PowerShell (Legacy) if Core is missing.
:: ---------------------------------------------------------------------

:: Set the current directory to the folder where this batch file is located
cd /d "%~dp0"

:: Default to Legacy PowerShell
set "PS_EXE=powershell.exe"

:: Check if PowerShell 7 (pwsh) exists
where pwsh >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    set "PS_EXE=pwsh"
)

:: FORCE UNBLOCK
:: Recursively unblocks all files so the background worker can load them.
%PS_EXE% -NoProfile -Command "Get-ChildItem -LiteralPath '%~dp0' -Recurse | Unblock-File"

:: Launch the script using the best available PowerShell version
start "" /MAX %PS_EXE% -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "W365Assessor.ps1"

exit
