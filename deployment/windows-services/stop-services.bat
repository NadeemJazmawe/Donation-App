@echo off
echo ========================================
echo Stopping Donation App Services
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

echo Stopping Backend service...
net stop "DonationApp-Backend"
if %errorLevel% equ 0 (
    echo [OK] Backend service stopped.
) else (
    echo [INFO] Backend service was not running or does not exist.
)

echo.
echo Stopping Frontend service...
net stop "DonationApp-Frontend"
if %errorLevel% equ 0 (
    echo [OK] Frontend service stopped.
) else (
    echo [INFO] Frontend service was not running or does not exist.
)

echo.
echo ========================================
echo Services Stopped
echo ========================================
echo.
pause
