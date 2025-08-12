# Simple IIS Deployment - Fresh Start

⚠️ **This completely replaces all previous deployment work and starts fresh.**

## What This Is

The simplest possible IIS deployment that actually works:
- **Frontend**: Static React files served directly by IIS  
- **Backend**: Python Flask app served via wfastcgi
- **No complex routing**: Backend is a separate IIS application at `/backend/`

## Architecture

```
IIS Default Web Site
└── excellence/                    (Frontend - Static files)
    ├── index.html
    ├── assets/
    ├── manifest.xml
    └── backend/                   (Backend - Python app)
        ├── app.py
        ├── wsgi_app.py
        └── web.config
```

## URLs After Deployment

- **Frontend**: `https://server:9443/excellence/`  
- **Backend Health**: `https://server:9443/excellence/backend/api/health`
- **Excel Manifest**: `https://server:9443/excellence/manifest.xml`

## Quick Deploy

```powershell
# 1. Install and build
npm install
npm run build:staging

# 2. Deploy everything (use -Force to overwrite)
.\deployment\simple\deploy-all.ps1 -Force

# 3. Test it works
.\deployment\simple\test-deployment.ps1
```

## Files

- `deploy-frontend.ps1` - Deploy React build to IIS
- `deploy-backend.ps1` - Setup Python backend with wfastcgi 
- `deploy-all.ps1` - Deploy both frontend and backend
- `test-deployment.ps1` - Test that everything works
- `uninstall.ps1` - Remove everything
- `QUICKSTART.md` - Quick installation guide

## Requirements

- Windows Server with IIS
- FastCGI module installed in IIS
- Python 3.8+ installed
- Node.js/npm (for building)

## If Problems

1. Check `QUICKSTART.md` for common issues
2. Use `uninstall.ps1` to clean up and start over
3. Check Windows Event Logs for detailed errors