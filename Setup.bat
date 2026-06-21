@echo off
title Markdown Formula Converter Installer
echo ======================================================
echo    Installing...
echo.   
echo    Markdown Formula Converter V1.1 for Microsoft Word
echo    Created by: Kartik Agrawat 
echo    Updated on 21 June 2026
echo ======================================================
echo.

:: Define source and destination paths
set "ADDIN_NAME=MarkdownConverter.dotm"
set "STARTUP_DIR=%APPDATA%\Microsoft\Word\Startup"

:: Check if the source file exists in the current folder
if not exist "%~dp0%ADDIN_NAME%" (
    call :Typewriter "[ERROR] Could not find %ADDIN_NAME% in this directory."
    call :Typewriter "Please make sure setup.bat and %ADDIN_NAME% are in the same folder."
    echo.
    pause
    goto EndOfScript
)

:: Create Word Startup directory if it doesn't exist yet
if not exist "%STARTUP_DIR%" (
    call :Typewriter "Creating Word Startup directory structure..."
    mkdir "%STARTUP_DIR%"
)

:: Close Word if it's currently open to prevent file locking issues
call :Typewriter "Closing Microsoft Word instances to prevent file locks..."
taskkill /f /im winword.exe >nul 2>&1
timeout /t 1 /nobreak >nul

:: Remove the Mark of the Web security block
call :Typewriter "Scanning file security protocols..."
powershell -Command "Unblock-File -Path '%~dp0%ADDIN_NAME%'"
call :Typewriter "Security clearance granted."
timeout /t 1 /nobreak >nul

:: Copy the Add-in file to the target Startup folder
call :Typewriter "Deploying macro logic into system architecture..."
copy /y "%~dp0%ADDIN_NAME%" "%STARTUP_DIR%\" >nul
timeout /t 1 /nobreak >nul

call :Typewriter "Registering Quick Access Toolbar shortcuts..."
timeout /t 1 /nobreak >nul

if %errorlevel% equ 0 (
    echo.
    echo ======================================================
    call :Typewriter "[SUCCESS] Installation Complete!"
    echo.
    call :Typewriter "A permanent conversion button has been added to your"
    call :Typewriter "Microsoft Word interface."
    echo.
    call :Typewriter "Look at the very top left of your Word desktop app"
    call :Typewriter "[above the Home ribbon] to find your custom icon."
    echo.
    call :Typewriter "Click that butterfly icon, "
    call :Typewriter "and watch the formulas instantly fix themselves!"
    echo ======================================================
) else (
    echo.
    call :Typewriter "[ERROR] Installation failed. Please run this script as Administrator."
)

echo.
pause
goto EndOfScript

:: ======================================================
:: THE TYPEWRITER ENGINE (Do not edit below this line)
:: ======================================================
:Typewriter
powershell -NoProfile -Command "$text='%~1'; foreach($c in $text.ToCharArray()){Write-Host -NoNewline $c; Start-Sleep -Milliseconds 20}; Write-Host ''"
exit /b

:EndOfScript