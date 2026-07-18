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
:: Unblocks the tool's own script files (non-recursive) so MOTW-marked downloads load cleanly.
%PS_EXE% -NoProfile -STA -Command "Get-ChildItem -Path '%~dp0*' -Include *.ps1,*.xaml,*.bat -File | Unblock-File"

:: Launch the script using the best available PowerShell version
start "" %PS_EXE% -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "W365Assessor.ps1"

exit
