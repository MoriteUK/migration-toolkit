@echo off
REM Migration Toolkit - Quick Start Script
echo.
echo ================================================
echo   Migration Toolkit - Web Edition
echo ================================================
echo.

REM Check if node_modules exists
if not exist "node_modules" (
    echo [1/2] Installing dependencies...
    echo.
    call npm install
    if errorlevel 1 (
        echo.
        echo ERROR: npm install failed
        echo Make sure Node.js is installed: winget install OpenJS.NodeJS.LTS
        pause
        exit /b 1
    )
    echo.
    echo Dependencies installed successfully!
    echo.
)

echo [2/2] Launching Migration Toolkit...
echo.
call npm start

if errorlevel 1 (
    echo.
    echo ERROR: Failed to start application
    pause
)
