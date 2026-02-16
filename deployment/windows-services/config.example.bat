@echo off
REM ========================================
REM Configuration Template for Deployment
REM ========================================
REM
REM Copy this file to config.bat and modify the paths
REM according to your installation.
REM
REM This file is optional - the scripts will auto-detect
REM paths based on the script location.
REM

REM Application root directory
REM Default: Auto-detected from script location
REM Example: set APP_DIR=C:\DonationApp
set APP_DIR=

REM Node.js executable path
REM Default: %APP_DIR%\nodejs\node.exe
REM Example: set NODE=C:\DonationApp\nodejs\node.exe
set NODE=

REM NSSM executable path
REM Default: %APP_DIR%\nssm\nssm-2.24\win64\nssm.exe
REM Example: set NSSM=C:\DonationApp\nssm\nssm-2.24\win64\nssm.exe
set NSSM=

REM Backend port (default: 5000)
set BACKEND_PORT=5000

REM Frontend port (default: 3000)
set FRONTEND_PORT=3000

REM MongoDB connection string
REM This should be set in backend\.env file
REM Example: mongodb://localhost:27017/donation-app
set MONGODB_URI=

REM ========================================
REM Service Names
REM ========================================
set BACKEND_SERVICE_NAME=DonationApp-Backend
set FRONTEND_SERVICE_NAME=DonationApp-Frontend

REM ========================================
REM Logging
REM ========================================
REM Logs directory (relative to APP_DIR)
set LOGS_DIR=logs

REM ========================================
REM Notes
REM ========================================
REM - If APP_DIR is empty, scripts will auto-detect
REM - Node.js should be portable version extracted to nodejs\ folder
REM - NSSM should be extracted to nssm\nssm-2.24\win64\ folder
REM - MongoDB must be installed separately
REM - All paths can use forward slashes or backslashes
