@echo off
title BaselinePilot
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0BaselinePilot.ps1"
