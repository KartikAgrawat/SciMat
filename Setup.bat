@echo off
title Markdown Formula Converter Installer
echo ======================================================
echo    Installing Markdown Formula Converter for Microsoft Word
echo    Created by: Kartik Agrawat
echo ======================================================
echo.

:: Define source and destination paths
set "ADDIN_NAME=MarkdownConverter.dotm"
set "STARTUP_DIR=%APPDATA%\Microsoft\Word\Startup"

:: Check if the source file exists in the current folder
if not exist "%~dp0%ADDIN_NAME%" (
    echo [ERROR] Could not find %ADDIN_NAME% in this directory.
    echo Please make sure setup.bat and %ADDIN_NAME% are in the same folder.
    echo.
    pause
    exit /b
)

:: Create Word Startup directory if it doesn't exist yet
if not exist "%STARTUP_DIR%" (
    echo Creating Word Startup directory structure...
    mkdir "%STARTUP_DIR%"
)

:: Close Word if it's currently open to prevent file locking issues
echo Closing Microsoft Word instances if running...
taskkill /f /im winword.exe >nul 2>&1

:: Copy the Add-in file to the target Startup folder
echo Deploying macro components...
copy /y "%~dp0%ADDIN_NAME%" "%STARTUP_DIR%\" >nul

if %errorlevel% equ 0 (
    echo.
    echo ======================================================
    echo [SUCCESS] Installation Complete!
    echo.
    echo A permanent conversion button has been added to your
    echo Microsoft Word interface.
    echo.
    echo Look at the very top left of your Word desktop app
    echo [above the Home ribbon] to find your custom icon.
    echo.
    echo Open any old document containing raw symbols, or a new 
    echo Word blank page and paste your text containing raw symbols 
    echo instead of formula, click that butterfly icon, 
    echo and watch the formulas instantly fix themselves!
    echo ======================================================
) else (
    echo.
    echo [ERROR] Installation failed. Please try running this script as Administrator.
)

echo.
pause