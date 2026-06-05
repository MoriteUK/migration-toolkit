@echo off
REM Right-click "Run as administrator" recommended for system-wide Node install.
REM Launches Setup.ps1 in the same folder.

cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Setup.ps1"
pause
