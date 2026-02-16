@echo off
echo ========================================
echo Preparing Donation App for Deployment
echo ========================================
echo.

REM Get the directory where this script is located
set SCRIPT_DIR=%~dp0
set APP_DIR=%SCRIPT_DIR%..\..\

echo Application Directory: %APP_DIR%
echo.

REM Check if Node.js exists
set NODE=%APP_DIR%nodejs\node.exe
if not exist "%NODE%" (
    echo ERROR: Node.js not found at %NODE%
    echo.
    echo Please download portable Node.js:
    echo   1. Go to https://nodejs.org/dist/v20.11.0/
    echo   2. Download: node-v20.11.0-win-x64.zip
    echo   3. Extract to: %APP_DIR%nodejs\
    echo.
    pause
    exit /b 1
)

echo [OK] Node.js found: %NODE%
echo.

REM Check if npm exists
set NPM=%APP_DIR%nodejs\npm.cmd
if not exist "%NPM%" (
    echo ERROR: npm not found at %NPM%
    pause
    exit /b 1
)

echo [OK] npm found: %NPM%
echo.

REM Step 1: Install root dependencies
echo [1/5] Installing root dependencies...
cd /d "%APP_DIR%"
call "%NPM%" install
if %errorLevel% neq 0 (
    echo ERROR: Failed to install root dependencies
    pause
    exit /b 1
)
echo [OK] Root dependencies installed
echo.

REM Step 2: Install backend dependencies
echo [2/5] Installing backend dependencies...
cd /d "%APP_DIR%backend"
call "%NPM%" install
if %errorLevel% neq 0 (
    echo ERROR: Failed to install backend dependencies
    pause
    exit /b 1
)
echo [OK] Backend dependencies installed
echo.

REM Step 3: Install frontend dependencies
echo [3/5] Installing frontend dependencies...
cd /d "%APP_DIR%frontend"
call "%NPM%" install
if %errorLevel% neq 0 (
    echo ERROR: Failed to install frontend dependencies
    pause
    exit /b 1
)
echo [OK] Frontend dependencies installed
echo.

REM Step 4: Install express in frontend (for production server)
echo [4/5] Installing express in frontend...
cd /d "%APP_DIR%frontend"
call "%NPM%" install express
if %errorLevel% neq 0 (
    echo ERROR: Failed to install express
    pause
    exit /b 1
)
echo [OK] Express installed
echo.

REM Step 5: Build frontend for production
echo [5/5] Building frontend for production...
cd /d "%APP_DIR%frontend"
call "%NPM%" run build
if %errorLevel% neq 0 (
    echo ERROR: Failed to build frontend
    pause
    exit /b 1
)
echo [OK] Frontend built successfully
echo.

REM Check if build directory exists
if not exist "%APP_DIR%frontend\build" (
    echo ERROR: Build directory not found after build
    pause
    exit /b 1
)

REM Check if frontend server.js exists
if not exist "%APP_DIR%frontend\server.js" (
    echo ERROR: frontend\server.js not found
    echo Please make sure frontend\server.js exists
    pause
    exit /b 1
)

REM Create logs directory
if not exist "%APP_DIR%logs" mkdir "%APP_DIR%logs"
echo [OK] Logs directory created
echo.

REM Check backend .env file
if not exist "%APP_DIR%backend\.env" (
    echo WARNING: backend\.env file not found
    echo Creating from env.example...
    if exist "%APP_DIR%backend\env.example" (
        copy "%APP_DIR%backend\env.example" "%APP_DIR%backend\.env" >nul
        echo [OK] Created backend\.env from env.example
        echo Please edit backend\.env with your MongoDB connection string
    ) else (
        echo ERROR: env.example not found
    )
    echo.
) else (
    echo [OK] Backend .env file exists
    echo.
)

echo ========================================
echo Preparation Complete!
echo ========================================
echo.
echo Next steps:
echo   1. Make sure MongoDB is installed and running
echo   2. Edit backend\.env with your MongoDB connection string
echo   3. Download NSSM from https://nssm.cc/download
echo   4. Extract NSSM to: %APP_DIR%nssm\nssm-2.24\win64\
echo   5. Run install-all-services.bat as Administrator
echo.
pause
