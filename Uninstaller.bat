@echo off
title Markdown Formula Converter Uninstaller
echo ======================================================
echo    Removing Markdown Formula Converter from Microsoft Word
echo    Uninstaller for tool created by: Kartik Agrawat
echo ======================================================
echo.

:: Define source and destination paths
set "ADDIN_NAME=MarkdownConverter.dotm"
set "STARTUP_DIR=%APPDATA%\Microsoft\Word\Startup"

:: Check if the tool is actually installed
if not exist "%STARTUP_DIR%\%ADDIN_NAME%" (
    echo [INFO] The Markdown Formula Converter is not found in the Word Startup folder.
    echo It appears to be uninstalled already.
    echo.
    pause
    exit /b
)

:: Close Word if it's currently open to prevent file locking issues
echo Closing Microsoft Word instances to safely remove files...
taskkill /f /im winword.exe >nul 2>&1

:: Delete the Add-in file from the target Startup folder
echo Removing macro components...
del /f /q "%STARTUP_DIR%\%ADDIN_NAME%" >nul

if %errorlevel% equ 0 (
    echo.
    echo ======================================================
    echo [SUCCESS] Removal Complete!
    echo.
    echo The conversion button has been completely removed from
    echo your Microsoft Word interface.
    echo.
    echo Thank you for using the Markdown Formula Converter.
    echo ======================================================
) else (
    echo.
    echo [ERROR] Removal failed. Please try running this script as Administrator.
)

echo.
pause