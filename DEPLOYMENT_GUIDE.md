# Excel Add-in - Deployment Guide

## NEW DEPLOYMENT SYSTEM

**This deployment system has been completely redesigned with a clean separation of concerns.**

⚠️ **BREAKING CHANGE**: The previous IIS-hosted deployment has been replaced with a modern service-based architecture.

## Quick Deployment

**Prerequisites:** Windows 10/Server with IIS, Node.js, Python 3.8+, PM2, and NSSM installed.

### 1. Initial Deployment (First Time)
```powershell
# Run as Administrator from the deployment folder
.\deploy-all.ps1
```

### 2. Update Existing Deployment
```powershell
# Run as Administrator to update services
.\update-all.ps1
```

### 3. Test Installation
```powershell
# Comprehensive testing suite
.\test-deployment.ps1
```

**Your application will be available at:** `https://server-vs81t.intranet.local:9443`

## New Architecture

```
Excel Client → IIS Proxy (Port 9443) → Backend Service (Port 5000)
                                   ↘ Frontend Service (Port 3000)
```

**Key Changes:**
- **IIS**: Reverse proxy ONLY (no hosting)
- **Backend**: Python Flask via NSSM Windows Service (Port 5000)
- **Frontend**: React app via PM2 Service (Port 3000)
- **SSL**: Handled at IIS proxy level
- **Simplified Scripts**: Only 5 deployment scripts total

### Service Details
- **Backend Service**: `ExcelAddin-Backend` (NSSM)
- **Frontend Service**: `exceladdin-frontend` (PM2)
- **IIS Site**: `ExcelAddin` (Reverse proxy on port 9443)

### URL Structure
- Frontend: `https://server-vs81t.intranet.local:9443/excellence/` → `http://localhost:3000/`
- API: `https://server-vs81t.intranet.local:9443/api/` → `http://localhost:5000/api/`
- Health check: `https://server-vs81t.intranet.local:9443/api/health`

## Development & Building

### Local Development
```bash
npm install
npm start
# Runs on https://localhost:3000
```

### Production Build
```bash
npm run build:staging  # For staging server
npm run build:prod     # For production server
```

## Troubleshooting

### Common Issues

**HTTP 500.19 Error (Configuration Error)**
- Cause: Conflicting web.config settings
- Solution: The web.config is designed to avoid common conflicts with IIS defaults

**Site not accessible**  
- Check IIS site is started: `Get-Website -Name "ExcelAddin"`
- Check firewall allows port 9443
- Verify SSL certificates exist in `C:\Cert\`

**API calls fail**
- Ensure Flask backend is running on port 5000
- Check backend service: `Get-Service -Name "ExcelAddin*"`

### Scripts Reference

| Script | Purpose | Usage |
|--------|---------|--------|
| `deploy-to-existing-iis.ps1` | Create IIS site and app pool | One-time setup |
| `build-and-deploy-iis.ps1` | Build React app and deploy | After code changes |
| `test-iis-simple.ps1` | Verify deployment works | Testing |
| `setup-backend-service.ps1` | Install Flask as Windows service | Backend setup |
| `add-firewall-rule.ps1` | Open firewall for port 9443 | If needed |

## Manual Configuration

If you prefer to configure manually instead of using scripts:

### IIS Site Configuration
- **Site Name:** ExcelAddin
- **Physical Path:** `C:\inetpub\wwwroot\ExcelAddin`  
- **Binding:** HTTPS port 9443 with server-vs81t certificate
- **Application Pool:** .NET v4.0 Integrated

### Required IIS Modules  
- URL Rewrite Module
- Application Request Routing (ARR)

### Directory Structure
```
C:\inetpub\wwwroot\ExcelAddin\
├── web.config              # IIS configuration
└── excellence\             # React app files
    ├── taskpane.html
    ├── commands.html  
    ├── *.js, *.css
    └── assets\
        └── manifest.xml
```

## PowerShell Testing Commands

### For PowerShell 6.0+
```powershell
# Test frontend
Invoke-WebRequest -Uri "https://server-vs81t.intranet.local:9443/excellence/" -SkipCertificateCheck

# Test API  
Invoke-WebRequest -Uri "https://server-vs81t.intranet.local:9443/excellence/api/health" -SkipCertificateCheck
```

### For Windows PowerShell 5.1
```powershell
# Disable certificate validation temporarily
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}

# Test frontend
Invoke-WebRequest -Uri "https://server-vs81t.intranet.local:9443/excellence/" -UseBasicParsing

# Test API
Invoke-WebRequest -Uri "https://server-vs81t.intranet.local:9443/excellence/api/health" -UseBasicParsing

# Re-enable certificate validation
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
```