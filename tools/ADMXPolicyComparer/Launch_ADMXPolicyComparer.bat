@echo off
title ADMX Policy Comparer
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Sta -File "%~dp0ADMXPolicyComparer.ps1"
