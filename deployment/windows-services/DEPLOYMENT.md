# Donation App - Detailed Deployment Guide

Complete step-by-step deployment process for Windows Services.

## Pre-Deployment Checklist

Before starting deployment, ensure you have:

- [ ] Windows machine (Windows 7 or later)
- [ ] Administrator access
- [ ] MongoDB installed and running
- [ ] Application source code
- [ ] Internet connection (for downloading Node.js and NSSM)

## Deployment Process

### Phase 1: Preparation

#### 1.1 Download Required Components

**Node.js (Portable):**
1. Visit: https://nodejs.org/dist/v20.11.0/
2. Download: `node-v20.11.0-win-x64.zip`
3. Extract to: `C:\DonationApp\nodejs\`
4. Verify: `C:\DonationApp\nodejs\node.exe --version`

**NSSM:**
1. Visit: https://nssm.cc/download
2. Download: `nssm-2.24.zip`
3. Extract to: `C:\DonationApp\nssm\nssm-2.24\win64\`
4. Verify: `C:\DonationApp\nssm\nssm-2.24\win64\nssm.exe`

**MongoDB:**
1. Download from: https://www.mongodb.com/try/download/community
2. Install MongoDB Community Server
3. Verify service is running: `sc query MongoDB`

#### 1.2 Copy Application Files

1. Copy entire application folder to target location (e.g., `C:\DonationApp\`)
2. Ensure folder structure is maintained:
   ```
   C:\DonationApp\
   ├── backend\
   ├── frontend\
   ├── deployment\
   └── ...
   ```

#### 1.3 Prepare Application

Run the preparation script:

```cmd
cd C:\DonationApp\deployment\windows-services
prepare-deployment.bat
```

**What this does:**
- Installs root dependencies
- Installs backend dependencies
- Installs frontend dependencies
- Installs Express in frontend (for production server)
- Builds frontend for production
- Creates logs directory
- Creates backend `.env` file if missing

**Expected output:**
```
[OK] Node.js found
[OK] npm found
[1/5] Installing root dependencies...
[OK] Root dependencies installed
[2/5] Installing backend dependencies...
[OK] Backend dependencies installed
[3/5] Installing frontend dependencies...
[OK] Frontend dependencies installed
[4/5] Installing express in frontend...
[OK] Express installed
[5/5] Building frontend for production...
[OK] Frontend built successfully
```

### Phase 2: Configuration

#### 2.1 Configure Backend

Edit `C:\DonationApp\backend\.env`:

```env
PORT=5000
MONGODB_URI=mongodb://localhost:27017/donation-app
```

**For remote MongoDB:**
```env
MONGODB_URI=mongodb://username:password@host:port/database-name
```

**For MongoDB Atlas:**
```env
MONGODB_URI=mongodb+srv://username:password@cluster.mongodb.net/donation-app
```

#### 2.2 Verify MongoDB Connection

Test MongoDB connection:

```cmd
mongosh
# or
mongo
```

If connection fails, check:
- MongoDB service is running
- Connection string is correct
- Firewall allows MongoDB port (27017)

#### 2.3 Verify Frontend Build

Check that frontend was built successfully:

```cmd
dir C:\DonationApp\frontend\build
```

Should see:
- `index.html`
- `static\` folder
- Other build artifacts

### Phase 3: Service Installation

#### 3.1 Install Services

**Important: Run as Administrator!**

1. Right-click `install-all-services.bat`
2. Select "Run as administrator"
3. Follow the prompts

**What happens:**
- Checks for administrator privileges
- Verifies Node.js and NSSM paths
- Installs backend service
- Installs frontend service
- Configures automatic startup
- Sets up logging

**Expected output:**
```
Installing Backend Service...
[OK] Service installed
Installing Frontend Service...
[OK] Service installed
Installation Complete!
```

#### 3.2 Verify Service Installation

Check services are installed:

```cmd
sc query DonationApp-Backend
sc query DonationApp-Frontend
```

Should show:
- STATE: STOPPED (or RUNNING)
- Service exists in Windows Services

### Phase 4: Start Services

#### 4.1 Start Services

**Run as Administrator:**

```cmd
cd C:\DonationApp\deployment\windows-services
start-services.bat
```

Or manually:
```cmd
net start DonationApp-Backend
net start DonationApp-Frontend
```

#### 4.2 Verify Services Are Running

```cmd
sc query DonationApp-Backend
sc query DonationApp-Frontend
```

Should show: `STATE: RUNNING`

#### 4.3 Test Application

1. **Frontend:** Open browser: http://localhost:3000
2. **Backend API:** http://localhost:5000/api/health

Expected responses:
- Frontend: Application loads
- Backend: `{"status":"OK"}`

### Phase 5: Post-Deployment Verification

#### 5.1 Check Logs

```cmd
type C:\DonationApp\logs\backend-out.log
type C:\DonationApp\logs\frontend-out.log
```

Look for:
- "Server is running on port 5000" (backend)
- "Frontend server running on port 3000" (frontend)
- "Connected to MongoDB" (backend)

#### 5.2 Check Windows Services

1. Open `services.msc`
2. Find:
   - DonationApp-Backend
   - DonationApp-Frontend
3. Verify:
   - Status: Running
   - Startup Type: Automatic
   - Log On: Local System

#### 5.3 Test Full Application Flow

1. Open http://localhost:3000
2. Add a new person
3. Verify data is saved
4. Check MongoDB:
   ```cmd
   mongosh
   use donation-app
   db.persons.find()
   ```

## Maintenance

### Updating the Application

1. **Stop services:**
   ```cmd
   stop-services.bat
   ```

2. **Update files:**
   - Replace application files
   - Update dependencies if needed

3. **Rebuild frontend:**
   ```cmd
   cd C:\DonationApp\frontend
   ..\nodejs\npm.cmd run build
   ```

4. **Start services:**
   ```cmd
   start-services.bat
   ```

### Viewing Logs

**Real-time log viewing:**
```cmd
REM Backend
powershell Get-Content C:\DonationApp\logs\backend-out.log -Wait -Tail 50

REM Frontend
powershell Get-Content C:\DonationApp\logs\frontend-out.log -Wait -Tail 50
```

**View error logs:**
```cmd
type C:\DonationApp\logs\backend-error.log
type C:\DonationApp\logs\frontend-error.log
```

### Restarting Services

**Restart both:**
```cmd
stop-services.bat
start-services.bat
```

**Restart individual:**
```cmd
net stop DonationApp-Backend
net start DonationApp-Backend
```

### Changing Ports

**Backend port:**
1. Edit `backend\.env`: `PORT=5001`
2. Restart service

**Frontend port:**
1. Edit NSSM service settings:
   ```cmd
   C:\DonationApp\nssm\nssm-2.24\win64\nssm.exe edit DonationApp-Frontend
   ```
2. Change AppEnvironmentExtra: `NODE_ENV=production&PORT=3001`
3. Restart service

## Troubleshooting

### Service Fails to Start

1. **Check logs:**
   ```cmd
   type C:\DonationApp\logs\backend-error.log
   ```

2. **Common issues:**
   - MongoDB not running
   - Port already in use
   - Missing dependencies
   - Incorrect paths

3. **Verify service configuration:**
   ```cmd
   C:\DonationApp\nssm\nssm-2.24\win64\nssm.exe status DonationApp-Backend
   ```

### Application Not Accessible

1. **Check services are running:**
   ```cmd
   sc query DonationApp-Backend
   sc query DonationApp-Frontend
   ```

2. **Check ports:**
   ```cmd
   netstat -ano | findstr :5000
   netstat -ano | findstr :3000
   ```

3. **Check firewall:**
   - Windows Firewall may block ports
   - Add exceptions for ports 3000 and 5000

### MongoDB Connection Issues

1. **Verify MongoDB is running:**
   ```cmd
   sc query MongoDB
   ```

2. **Test connection:**
   ```cmd
   mongosh
   ```

3. **Check connection string in `.env`**

4. **Check MongoDB logs:**
   ```cmd
   type "C:\Program Files\MongoDB\Server\*\log\mongod.log"
   ```

## Uninstallation

### Complete Removal

1. **Stop services:**
   ```cmd
   stop-services.bat
   ```

2. **Uninstall services:**
   ```cmd
   uninstall-services.bat
   ```

3. **Verify removal:**
   ```cmd
   sc query DonationApp-Backend
   sc query DonationApp-Frontend
   ```
   Should show: "The specified service does not exist"

4. **Delete application folder (optional):**
   ```cmd
   rmdir /s C:\DonationApp
   ```

## Security Considerations

1. **MongoDB Authentication:**
   - Enable authentication in MongoDB
   - Use strong passwords
   - Update connection string in `.env`

2. **Firewall:**
   - Only expose necessary ports
   - Use firewall rules to restrict access

3. **Service Account:**
   - Services run as Local System (default)
   - Consider using dedicated service account for production

4. **Environment Variables:**
   - Keep `.env` files secure
   - Don't commit `.env` to version control

## Performance Tuning

### Log Rotation

Logs are automatically rotated:
- Daily rotation
- 10MB max file size
- Old logs are archived

### Service Recovery

Configure service recovery in NSSM:
```cmd
C:\DonationApp\nssm\nssm-2.24\win64\nssm.exe edit DonationApp-Backend
```

Set:
- First failure: Restart service
- Second failure: Restart service
- Subsequent failures: Restart service

## Backup and Recovery

### Backup MongoDB

```cmd
mongodump --db donation-app --out C:\Backup\mongodb\
```

### Backup Application

1. Stop services
2. Copy entire `C:\DonationApp\` folder
3. Start services

### Recovery

1. Restore MongoDB:
   ```cmd
   mongorestore --db donation-app C:\Backup\mongodb\donation-app
   ```

2. Restore application files
3. Restart services

## Support and Maintenance Schedule

### Daily
- Check service status
- Review error logs

### Weekly
- Review application logs
- Check disk space
- Verify backups

### Monthly
- Update dependencies (if needed)
- Review security
- Performance monitoring

## Additional Resources

- NSSM Documentation: https://nssm.cc/usage
- MongoDB Documentation: https://docs.mongodb.com/
- Node.js Documentation: https://nodejs.org/docs/
