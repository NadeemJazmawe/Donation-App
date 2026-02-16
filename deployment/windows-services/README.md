# Donation App - Windows Service Deployment

Complete guide for deploying the Donation App as Windows Services using NSSM (Non-Sucking Service Manager).

## Prerequisites

**Only MongoDB is required as an external dependency!**

1. **MongoDB** (Community Server)
   - Download from: https://www.mongodb.com/try/download/community
   - Install MongoDB Community Server
   - Ensure MongoDB service is running (check Windows Services)

2. **Portable Node.js** (bundled with deployment)
   - Download from: https://nodejs.org/dist/v20.11.0/
   - Download: `node-v20.11.0-win-x64.zip` (or latest LTS)
   - Extract to: `nodejs\` folder in application directory

3. **NSSM** (Non-Sucking Service Manager)
   - Download from: https://nssm.cc/download
   - Download: `nssm-2.24.zip`
   - Extract to: `nssm\nssm-2.24\win64\` folder

## Quick Start

### Step 1: Prepare the Application

1. Copy the entire application folder to your target Windows machine (e.g., `C:\DonationApp\`)

2. Download and extract portable Node.js:
   ```
   C:\DonationApp\nodejs\
   ├── node.exe
   ├── npm.cmd
   └── ...
   ```

3. Download and extract NSSM:
   ```
   C:\DonationApp\nssm\nssm-2.24\win64\
   └── nssm.exe
   ```

4. Run the preparation script:
   ```cmd
   cd C:\DonationApp\deployment\windows-services
   prepare-deployment.bat
   ```

   This will:
   - Install all dependencies
   - Build the frontend for production
   - Install Express in frontend
   - Create necessary directories
   - Create backend `.env` file if it doesn't exist

### Step 2: Configure MongoDB

1. Edit `backend\.env`:
   ```env
   PORT=5000
   MONGODB_URI=mongodb://localhost:27017/donation-app
   ```

   For remote MongoDB:
   ```env
   MONGODB_URI=mongodb://username:password@host:port/database-name
   ```

2. Ensure MongoDB is running:
   ```cmd
   sc query MongoDB
   ```

### Step 3: Install Services

**Run as Administrator:**

```cmd
cd C:\DonationApp\deployment\windows-services
install-all-services.bat
```

This installs both:
- **DonationApp-Backend** (Port 5000)
- **DonationApp-Frontend** (Port 3000)

### Step 4: Start Services

```cmd
start-services.bat
```

Or manually:
```cmd
net start DonationApp-Backend
net start DonationApp-Frontend
```

### Step 5: Verify Installation

1. Open browser: http://localhost:3000
2. Check backend API: http://localhost:5000/api/health
3. Check Windows Services (`services.msc`):
   - DonationApp-Backend (Running, Automatic)
   - DonationApp-Frontend (Running, Automatic)

## Service Management

### Start Services
```cmd
start-services.bat
```

### Stop Services
```cmd
stop-services.bat
```

### Restart Services
```cmd
stop-services.bat
start-services.bat
```

### Uninstall Services
```cmd
uninstall-services.bat
```

### Using Windows Services GUI
1. Press `Win + R`, type `services.msc`
2. Find "DonationApp-Backend" and "DonationApp-Frontend"
3. Right-click → Start/Stop/Restart

### Using Command Line
```cmd
REM Start
net start DonationApp-Backend
net start DonationApp-Frontend

REM Stop
net stop DonationApp-Backend
net stop DonationApp-Frontend

REM Status
sc query DonationApp-Backend
sc query DonationApp-Frontend
```

## Logs

Service logs are located in:
```
C:\DonationApp\logs\
├── backend-out.log
├── backend-error.log
├── frontend-out.log
└── frontend-error.log
```

View logs:
```cmd
type C:\DonationApp\logs\backend-out.log
type C:\DonationApp\logs\frontend-out.log
```

## Troubleshooting

### Service Won't Start

1. **Check logs:**
   ```cmd
   type C:\DonationApp\logs\backend-error.log
   type C:\DonationApp\logs\frontend-error.log
   ```

2. **Verify MongoDB is running:**
   ```cmd
   sc query MongoDB
   ```

3. **Check service status:**
   ```cmd
   sc query DonationApp-Backend
   sc query DonationApp-Frontend
   ```

4. **Verify Node.js path:**
   ```cmd
   C:\DonationApp\nodejs\node.exe --version
   ```

5. **Verify NSSM path:**
   ```cmd
   C:\DonationApp\nssm\nssm-2.24\win64\nssm.exe
   ```

### Port Already in Use

```cmd
REM Check what's using the ports
netstat -ano | findstr :5000
netstat -ano | findstr :3000

REM Stop the process or change ports in:
REM - backend\.env (PORT=5000)
REM - NSSM service settings (for frontend PORT=3000)
```

### MongoDB Connection Error

1. Check `backend\.env` MongoDB URI
2. Verify MongoDB service is running
3. Test connection:
   ```cmd
   mongosh
   # or
   mongo
   ```

### Node.js Not Found

- Verify `C:\DonationApp\nodejs\node.exe` exists
- Check NSSM service settings:
  ```cmd
  C:\DonationApp\nssm\nssm-2.24\win64\nssm.exe edit DonationApp-Backend
  ```

### Frontend Not Loading

1. Verify frontend is built:
   ```cmd
   dir C:\DonationApp\frontend\build
   ```

2. Check if `frontend\server.js` exists

3. Verify Express is installed:
   ```cmd
   dir C:\DonationApp\frontend\node_modules\express
   ```

## Service Configuration

### View Service Details
```cmd
C:\DonationApp\nssm\nssm-2.24\win64\nssm.exe status DonationApp-Backend
C:\DonationApp\nssm\nssm-2.24\win64\nssm.exe status DonationApp-Frontend
```

### Edit Service Settings
```cmd
C:\DonationApp\nssm\nssm-2.24\win64\nssm.exe edit DonationApp-Backend
C:\DonationApp\nssm\nssm-2.24\win64\nssm.exe edit DonationApp-Frontend
```

### Service Properties
- **Startup Type:** Automatic (starts on Windows boot)
- **Log Rotation:** Enabled (daily, 10MB max)
- **Working Directory:** Set automatically
- **Environment Variables:** NODE_ENV=production

## Updating the Application

1. Stop services:
   ```cmd
   stop-services.bat
   ```

2. Update application files

3. Rebuild frontend:
   ```cmd
   cd C:\DonationApp\frontend
   ..\nodejs\npm.cmd run build
   ```

4. Start services:
   ```cmd
   start-services.bat
   ```

## Uninstallation

1. Stop services:
   ```cmd
   stop-services.bat
   ```

2. Uninstall services:
   ```cmd
   uninstall-services.bat
   ```

3. (Optional) Delete application folder

## File Structure

```
C:\DonationApp\
├── nodejs\                    # Portable Node.js
│   ├── node.exe
│   ├── npm.cmd
│   └── ...
├── nssm\                      # NSSM tool
│   └── nssm-2.24\
│       └── win64\
│           └── nssm.exe
├── backend\
│   ├── server.js
│   ├── .env
│   ├── package.json
│   └── ...
├── frontend\
│   ├── server.js
│   ├── build\                 # Built React app
│   ├── package.json
│   └── ...
├── deployment\
│   └── windows-services\
│       ├── install-all-services.bat
│       ├── install-backend-service.bat
│       ├── install-frontend-service.bat
│       ├── uninstall-services.bat
│       ├── start-services.bat
│       ├── stop-services.bat
│       ├── prepare-deployment.bat
│       └── README.md
└── logs\                      # Service logs
    ├── backend-out.log
    ├── backend-error.log
    ├── frontend-out.log
    └── frontend-error.log
```

## Support

For issues or questions:
1. Check the logs in `logs\` directory
2. Verify all prerequisites are installed
3. Ensure services are running as Administrator
4. Review the troubleshooting section above

## Notes

- Services start automatically when Windows boots
- Logs are rotated daily (10MB max per file)
- Only MongoDB needs to be installed separately
- Node.js is portable (no system installation required)
- All scripts must be run as Administrator
