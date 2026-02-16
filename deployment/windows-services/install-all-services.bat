@echo off
echo ========================================
echo Donation App - Service Installation
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

echo This will install both Backend and Frontend services.
echo.
pause

echo.
echo [1/2] Installing Backend Service...
echo ========================================
call "%~dp0install-backend-service.bat"
if %errorLevel% neq 0 (
    echo ERROR: Failed to install backend service
    pause
    exit /b 1
)

echo.
echo [2/2] Installing Frontend Service...
echo ========================================
call "%~dp0install-frontend-service.bat"
if %errorLevel% neq 0 (
    echo ERROR: Failed to install frontend service
    pause
    exit /b 1
)

echo.
echo ========================================
echo Installation Complete!
echo ========================================
echo.
echo Services installed:
echo   - DonationApp-Backend (Port 5000)
echo   - DonationApp-Frontend (Port 3000)
echo.
echo To start services:
echo   net start DonationApp-Backend
echo   net start DonationApp-Frontend
echo.
echo Or use: start-services.bat
echo.
echo To stop services:
echo   net stop DonationApp-Backend
echo   net stop DonationApp-Frontend
echo.
echo Or use: stop-services.bat
echo.
echo Services are configured to start automatically on Windows boot.
echo.
pause
