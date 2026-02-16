@echo off
echo ========================================
echo Starting Donation App Services
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

echo Starting Backend service...
net start "DonationApp-Backend"
if %errorLevel% equ 0 (
    echo [OK] Backend service started successfully!
) else (
    echo [ERROR] Failed to start Backend service.
    echo Check logs at: logs\backend-error.log
)

echo.
echo Starting Frontend service...
net start "DonationApp-Frontend"
if %errorLevel% equ 0 (
    echo [OK] Frontend service started successfully!
) else (
    echo [ERROR] Failed to start Frontend service.
    echo Check logs at: logs\frontend-error.log
)

echo.
echo ========================================
echo Service Status
echo ========================================
echo.
sc query "DonationApp-Backend" | findstr /C:"STATE"
sc query "DonationApp-Frontend" | findstr /C:"STATE"
echo.
echo Services should be accessible at:
echo   Frontend: http://localhost:3000
echo   Backend API: http://localhost:5000/api/health
echo.
pause
