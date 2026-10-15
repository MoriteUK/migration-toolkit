@echo off
:: Migration Toolkit — Prerequisites Installer
:: Double-click this file (or run from an Admin command prompt).

:: Self-elevate to Administrator if not already
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting administrator privileges...
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

:: Prefer pwsh.exe (PS7); fall back to Windows PowerShell 5.1
set PS_EXE=powershell.exe
where pwsh.exe >nul 2>&1 && set PS_EXE=pwsh.exe

echo.
echo ================================================
echo   Migration Toolkit - Prerequisites Installer
echo ================================================
echo.
echo Running with: %PS_EXE%
echo.

%PS_EXE% -NoProfile -ExecutionPolicy Bypass -File "%~dp0VGMigrations\Install-Prerequisites.ps1" %*

echo.
pause
