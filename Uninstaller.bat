@echo off
title Markdown Formula Converter Uninstaller
echo ======================================================
echo    Removing...
echo.
echo    Markdown Formula Converter V1.1.0 from Microsoft Word
echo    Uninstaller for tool created by: Kartik Agrawat
echo    Updated on 21 June 2026
echo ======================================================
echo.

:: Define source and destination paths
set "ADDIN_NAME=MarkdownConverter.dotm"
set "STARTUP_DIR=%APPDATA%\Microsoft\Word\Startup"

:: Check if the tool is actually installed
echo Scanning system architecture for installed modules...
timeout /t 2 /nobreak >nul

if not exist "%STARTUP_DIR%\%ADDIN_NAME%" (
    echo [INFO] The Markdown Formula Converter was not found in the Word Startup folder.
    echo It appears the tool is already uninstalled or was never installed on this system.
    echo.
    pause
    exit /b
)

:: Close Word if it's currently open to prevent file locking issues
echo Closing Microsoft Word instances to release file locks...
taskkill /f /im winword.exe >nul 2>&1
timeout /t 2 /nobreak >nul

:: Delete the Add-in file from the target Startup folder
echo Purging macro components from Word STARTUP directory...
del /f /q "%STARTUP_DIR%\%ADDIN_NAME%" >nul
timeout /t 1 /nobreak >nul

echo Clearing residual system configurations...
timeout /t 3 /nobreak >nul

if %errorlevel% equ 0 (
    echo.
    echo ======================================================
    echo [SUCCESS] Removal Complete!
    echo.
    echo The conversion logic has been completely purged from
    echo your Microsoft Word environment.
    echo.
    echo Thank you for using the Markdown Formula Converter.
    echo ======================================================
) else (
    echo.
    echo [ERROR] Removal failed. Please run this script as Administrator.
)

echo.
pause