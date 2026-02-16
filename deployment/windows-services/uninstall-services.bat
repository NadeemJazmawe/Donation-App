@echo off
echo ========================================
echo Uninstalling Donation App Services
echo ========================================
echo.

REM Check if running as Administrator
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo ERROR: This script must be run as Administrator!
    echo Right-click and select "Run as administrator"
    pause
    exit /b 1
)

echo This will remove both Backend and Frontend services.
echo.
set /p CONFIRM="Are you sure you want to uninstall? (Y/N): "
if /i not "%CONFIRM%"=="Y" (
    echo Uninstallation cancelled.
    pause
    exit /b 0
)

echo.
echo Stopping services...
net stop "DonationApp-Backend" >nul 2>&1
if %errorLevel% equ 0 (
    echo Backend service stopped.
) else (
    echo Backend service was not running.
)

net stop "DonationApp-Frontend" >nul 2>&1
if %errorLevel% equ 0 (
    echo Frontend service stopped.
) else (
    echo Frontend service was not running.
)

echo.
echo Removing services...

REM Get the directory where this script is located
set SCRIPT_DIR=%~dp0
set APP_DIR=%SCRIPT_DIR%..\..\
set NSSM=%APP_DIR%nssm\nssm-2.24\win64\nssm.exe

if exist "%NSSM%" (
    "%NSSM%" remove "DonationApp-Backend" confirm
    if %errorLevel% equ 0 (
        echo Backend service removed successfully.
    ) else (
        echo Warning: Failed to remove backend service. It may not exist.
    )
    
    "%NSSM%" remove "DonationApp-Frontend" confirm
    if %errorLevel% equ 0 (
        echo Frontend service removed successfully.
    ) else (
        echo Warning: Failed to remove frontend service. It may not exist.
    )
) else (
    echo ERROR: NSSM not found. Cannot remove services.
    echo Attempting to remove using sc delete...
    sc delete "DonationApp-Backend" >nul 2>&1
    sc delete "DonationApp-Frontend" >nul 2>&1
    echo Services removed using sc delete.
)

echo.
echo ========================================
echo Uninstallation Complete!
echo ========================================
echo.
pause
