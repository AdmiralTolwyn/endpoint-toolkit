@echo off
:: ---------------------------------------------------------------------
:: AVD Assessor Launcher
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
:: Unblocks only the script/UI files needed to launch (not assessments/reports/_backups).
%PS_EXE% -STA -NoProfile -Command "Get-ChildItem -Path '%~dp0*' -Include *.ps1,*.xaml,*.bat -File | Unblock-File"

:: Launch the script using the best available PowerShell version
start "" %PS_EXE% -STA -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "AvdAssessor.ps1"

exit
