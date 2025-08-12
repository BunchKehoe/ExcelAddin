# Simple IIS Deployment - Quick Start

## What this does

This creates the simplest possible IIS deployment:

1. **Frontend**: React build served as static files at `/excellence/`
2. **Backend**: Python Flask app served via wfastcgi at `/excellence/backend/`

## Quick Install

```powershell
# 1. Install dependencies and build
npm install
npm run build:staging

# 2. Deploy everything
.\deployment\simple\deploy-all.ps1 -Force

# 3. Test it works
.\deployment\simple\test-deployment.ps1
```

## URLs After Deployment

- **Frontend**: `https://server:9443/excellence/`
- **Backend API**: `https://server:9443/excellence/backend/api/health`
- **Excel Add-in Manifest**: `https://server:9443/excellence/manifest.xml`

## If something fails

1. **Frontend not loading**: Check IIS Manager, verify `/excellence/` application exists
2. **Backend API not working**: Check Event Logs, verify Python/wfastcgi installation
3. **Certificate errors**: Run `npm run cert:fix` for local development

## Directory Structure After Deployment

```
C:\inetpub\wwwroot\excellence\
├── index.html              (Frontend files)
├── assets/
├── manifest.xml
└── backend/                (Backend application)
    ├── app.py
    ├── wsgi_app.py
    ├── web.config
    └── src/
```

## Requirements

- IIS with FastCGI module
- Python 3.8+
- Node.js/npm (for building)