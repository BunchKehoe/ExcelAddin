# Excel Add-in - IIS Deployment Guide

## Quick Setup for Existing IIS Servers

**Prerequisites:** Windows Server with IIS, URL Rewrite Module, and Application Request Routing (ARR) installed.

### 1. Deploy to IIS (One-time Setup)
```powershell
# Run as Administrator
.\deployment\scripts\deploy-to-existing-iis.ps1
```

### 2. Build and Deploy Application  
```powershell
# Build React app and deploy to IIS
.\deployment\scripts\build-and-deploy-iis.ps1
```

### 3. Test Installation
```powershell
# Verify everything works
.\deployment\scripts\test-iis-simple.ps1
```

**That's it!** Your application will be available at: `https://server-vs81t.intranet.local:9443/excellence/`

## How It Works

### Architecture
```
Excel Add-in → IIS (Port 9443) → Flask Backend (Port 5000)
```

- **IIS** serves React frontend files and proxies API calls
- **Flask** backend handles API requests on port 5000  
- **SSL** certificates from `C:\Cert\server-vs81t.(crt|key)`
- **Files** deployed to `C:\inetpub\wwwroot\ExcelAddin\excellence\`

### URL Structure
- Frontend: `https://server-vs81t.intranet.local:9443/excellence/`
- API: `https://server-vs81t.intranet.local:9443/excellence/api/*` → proxied to `http://localhost:5000/api/*`
- Health check: `https://server-vs81t.intranet.local:9443/health`

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