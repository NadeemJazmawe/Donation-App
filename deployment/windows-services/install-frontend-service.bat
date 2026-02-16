@echo off
echo ========================================
echo Installing Donation App Frontend Service
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

REM Get the directory where this script is located
set SCRIPT_DIR=%~dp0
set APP_DIR=%SCRIPT_DIR%..\..\

REM Check if NSSM exists
set NSSM=%APP_DIR%nssm\nssm-2.24\win64\nssm.exe
if not exist "%NSSM%" (
    echo ERROR: NSSM not found at %NSSM%
    echo Please download NSSM from https://nssm.cc/download
    echo Extract it to: %APP_DIR%nssm\nssm-2.24\win64\
    pause
    exit /b 1
)

REM Check if Node.js exists
set NODE=%APP_DIR%nodejs\node.exe
if not exist "%NODE%" (
    echo ERROR: Node.js not found at %NODE%
    echo Please download portable Node.js and extract it to: %APP_DIR%nodejs\
    pause
    exit /b 1
)

REM Check if frontend server.js exists
set SERVER_SCRIPT=%APP_DIR%frontend\server.js
if not exist "%SERVER_SCRIPT%" (
    echo ERROR: Frontend server.js not found at %SERVER_SCRIPT%
    echo Please make sure frontend/server.js exists
    pause
    exit /b 1
)

echo Installing service...
echo Service Name: DonationApp-Frontend
echo Node.js: %NODE%
echo Script: %SERVER_SCRIPT%
echo Working Directory: %APP_DIR%frontend
echo.

REM Stop service if it already exists
sc query "DonationApp-Frontend" >nul 2>&1
if %errorLevel% equ 0 (
    echo Service already exists. Stopping it first...
    net stop "DonationApp-Frontend" >nul 2>&1
    timeout /t 2 >nul
)

REM Install the service
"%NSSM%" install "DonationApp-Frontend" "%NODE%" "%SERVER_SCRIPT%"

if %errorLevel% neq 0 (
    echo ERROR: Failed to install service
    pause
    exit /b 1
)

REM Set working directory
"%NSSM%" set "DonationApp-Frontend" AppDirectory "%APP_DIR%frontend"

REM Set environment variables
"%NSSM%" set "DonationApp-Frontend" AppEnvironmentExtra "NODE_ENV=production^&PORT=3000"

REM Set description
"%NSSM%" set "DonationApp-Frontend" Description "Donation App Frontend Web Server"

REM Set startup type to Automatic
"%NSSM%" set "DonationApp-Frontend" Start SERVICE_AUTO_START

REM Create logs directory if it doesn't exist
if not exist "%APP_DIR%logs" mkdir "%APP_DIR%logs"

REM Set output files
"%NSSM%" set "DonationApp-Frontend" AppStdout "%APP_DIR%logs\frontend-out.log"
"%NSSM%" set "DonationApp-Frontend" AppStderr "%APP_DIR%logs\frontend-error.log"

REM Set log rotation
"%NSSM%" set "DonationApp-Frontend" AppRotateFiles 1
"%NSSM%" set "DonationApp-Frontend" AppRotateOnline 1
"%NSSM%" set "DonationApp-Frontend" AppRotateSeconds 86400
"%NSSM%" set "DonationApp-Frontend" AppRotateBytes 10485760

echo.
echo ========================================
echo Frontend service installed successfully!
echo ========================================
echo.
echo Service Name: DonationApp-Frontend
echo Port: 3000
echo.
echo To start the service:
echo   net start DonationApp-Frontend
echo.
echo To view service status:
echo   sc query DonationApp-Frontend
echo.
pause
