# Required Dependencies

To run this application, you need the following software installed on your PC:

## 1. Node.js and npm (Required)

**What it is:** Node.js is a JavaScript runtime that allows you to run JavaScript on your computer. npm (Node Package Manager) comes bundled with Node.js and is used to install project dependencies.

**Minimum Version:** Node.js v14.0.0 or higher  
**Recommended:** Node.js v18.x.x or v20.x.x (LTS)  
**Tested with:** Node.js v20.19.6 ✅

### Installation:

**Windows:**
1. Download the installer from: https://nodejs.org/
2. Choose the LTS (Long Term Support) version
3. Run the installer and follow the setup wizard
4. Make sure to check "Add to PATH" during installation

**Verify Installation:**
Open PowerShell or Command Prompt and run:
```bash
node --version
npm --version
```
You should see version numbers for both commands.

---

## 2. MongoDB (Required)

**What it is:** MongoDB is a NoSQL database that stores your application data locally.

**Version:** MongoDB Community Server 6.0 or higher

### Installation:

**Windows:**
1. Download MongoDB Community Server from: https://www.mongodb.com/try/download/community
2. Run the installer (.msi file)
3. Choose "Complete" installation
4. **Important:** Check "Install MongoDB as a Service" and "Run service as Network Service user"
5. Check "Install MongoDB Compass" (optional GUI tool)
6. Complete the installation

**After Installation:**
- MongoDB should start automatically as a Windows service
- If not, you can start it manually:
  - Open Services (Win + R, type `services.msc`)
  - Find "MongoDB" service
  - Right-click and select "Start"

**Verify Installation:**
Open PowerShell or Command Prompt and run:
```bash
mongod --version
```
You should see MongoDB version information.

**Alternative - MongoDB via Chocolatey (if you have Chocolatey):**
```bash
choco install mongodb
```

---

## 3. Modern Web Browser (Required)

**What it is:** You need a web browser to view and interact with the application.

**Supported Browsers:**
- Google Chrome (recommended)
- Microsoft Edge
- Mozilla Firefox
- Safari (Mac)

The application will run at `http://localhost:3000` in your browser.

---

## 4. Git (Optional but Recommended)

**What it is:** Version control system. Useful for managing your code.

**Installation:**
- Download from: https://git-scm.com/download/win
- Or install via Chocolatey: `choco install git`

---

## Quick Verification Checklist

Before running the application, verify you have everything:

```bash
# Check Node.js
node --version
# Should show: v14.x.x or higher

# Check npm
npm --version
# Should show: 6.x.x or higher

# Check MongoDB
mongod --version
# Should show: db version v6.x.x or higher

# Check if MongoDB is running
# Open Services (services.msc) and verify "MongoDB" service is running
```

---

## Installation Order

1. **First:** Install Node.js (includes npm)
2. **Second:** Install MongoDB
3. **Third:** Verify both are working
4. **Fourth:** Run the application setup script

---

## Troubleshooting

### MongoDB not starting?
- Check if the MongoDB service is running in Windows Services
- Make sure port 27017 is not blocked by firewall
- Try starting MongoDB manually: `net start MongoDB`

### Node.js not found?
- Restart your terminal/PowerShell after installation
- Make sure Node.js was added to PATH during installation
- Reinstall Node.js if needed

### Port already in use?
- If port 3000 is in use, the React app will automatically try the next available port
- If port 5000 is in use, change it in `backend/.env` file

---

## Next Steps

Once all dependencies are installed:

1. Navigate to the project directory
2. Run the setup script: `setup.bat` (Windows) or `./setup.sh` (Linux/Mac)
3. Start the application: `npm run dev`
4. Open your browser to `http://localhost:3000`
