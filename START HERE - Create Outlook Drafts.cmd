@echo off
setlocal
title Water Cooler Outlook Draft Importer
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Create Outlook Drafts.ps1"
echo.
pause
