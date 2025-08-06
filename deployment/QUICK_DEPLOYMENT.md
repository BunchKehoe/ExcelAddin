# Quick Production Deployment Guide

## Problem
When trying to build with `npm install --production` followed by `npm run build`, you get:
```
Der Befehl "custom-functions-metadata" ist entweder falsch geschrieben oder konnte nicht gefunden werden.
```

## Solution
Use the proper build-then-deploy approach instead of trying to build on the production server with limited dependencies.

## Recommended Deployment Process

### Step 1: Build with Full Dependencies
```bash
# On your development machine or CI/CD server
npm install                # Install ALL dependencies (including devDependencies)
npm run build:staging      # Build for staging/production
```

### Step 2: Deploy Built Assets  
```bash
# Copy only the dist/ folder to your production server
scp -r dist/* user@server:/path/to/deployment/
# OR use your preferred deployment method
```

### Step 3: Configure Web Server
Point your nginx or IIS to serve files from the deployed `dist/` directory.

## Alternative: Use Deployment Script

The included PowerShell deployment script handles everything:
```powershell
.\deployment\scripts\deploy-windows.ps1
```

This script:
1. Installs all dependencies
2. Builds the application  
3. Configures nginx
4. Sets up Windows services

## Why This Approach?

- ✅ **Proper separation:** Build tools stay on build environment
- ✅ **Faster deployments:** Only deploy compiled assets
- ✅ **Security:** Production server doesn't need dev tools
- ✅ **Reliability:** No dependency conflicts on production
- ✅ **Performance:** Smaller production footprint

## Files You Need on Production Server
```
dist/
├── taskpane.html           # Main Excel add-in interface
├── commands.html           # Custom functions page  
├── functions.json          # Excel custom functions metadata
├── manifest.xml            # Excel add-in manifest
├── *.js                    # Compiled JavaScript bundles
└── assets/                 # Icons and static files
```

## Production Server Requirements
- Web server (nginx, IIS, Apache)
- SSL certificate
- No Node.js or npm required!

For detailed deployment instructions, see `deployment/WINDOWS_DEPLOYMENT.md`.