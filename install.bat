@echo off
:: DataLink installer — double-click this file or run it from any terminal.
:: Requires Node.js 16+ (https://nodejs.org)
node "%~dp0setup.js"
if %errorlevel% neq 0 (
    echo.
    echo Setup failed. Make sure Node.js 16+ is installed.
    pause
)
