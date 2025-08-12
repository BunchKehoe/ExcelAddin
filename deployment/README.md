# ExcelAddin Deployment Guide

This deployment system hosts the ExcelAddin service with a clean separation of concerns:
- **IIS**: Reverse proxy only (HTTPS termination and request forwarding)
- **Backend**: Python Flask API hosted via NSSM service
- **Frontend**: React application hosted via NSSM service (Windows-compatible, no PM2)

## Server Configuration
- **Service Name**: ExcelAddin
- **Public Address**: https://server-vs81t.intranet.local:9443
- **Target Environment**: Windows 10/11
- **Certificate Path**: C:\Cert\ (for SSL certificates)

## Architecture Overview

```
Excel Client → IIS Proxy (Port 9443) → Backend Service (Port 5000)
                                   ↘ Frontend Service (Port 3000)
```

## Prerequisites

1. **Node.js and npm** installed
2. **Python 3.8+** with pip installed
3. **NSSM** downloaded and available in PATH
4. **IIS** with URL Rewrite module installed
5. **SSL Certificate** placed in C:\Cert\ directory (*.pfx or *.p12 format)

## SSL Certificate Configuration

Place your SSL certificate file in `C:\Cert\`:
- Supported formats: `.pfx` or `.p12`
- The deployment scripts will automatically import and bind the certificate
- If no certificate is found in C:\Cert\, the scripts will search the Windows certificate store
- Manual binding instructions will be provided if automatic binding fails

## Deployment Scripts

### Initial Deployment
```powershell
.\deploy-all.ps1
```
Performs complete first-time deployment including service installation and IIS configuration.

### Updates After First Run
```powershell
.\update-all.ps1
```
Updates services without reinstalling or reconfiguring infrastructure.

### Individual Service Deployment
```powershell
.\deploy-backend.ps1             # Deploy Python backend via NSSM
.\deploy-frontend.ps1            # Deploy React frontend via NSSM
.\deploy-frontend.ps1 -ConfigureIIS  # Deploy frontend and configure IIS
```

**Note**: Both individual deployment scripts automatically stop and overwrite existing services without requiring the `-Force` flag. This makes re-deployments seamless.

### Testing
```powershell
.\test-deployment.ps1
```
Runs comprehensive tests to verify all services are running and accessible.

## Service Details

### Backend Service (NSSM)
- **Service Name**: ExcelAddin-Backend
- **Port**: 5000
- **Environment**: staging
- **Health Check**: http://localhost:5000/api/health

### Frontend Service (NSSM)
- **Service Name**: ExcelAddin-Frontend  
- **Port**: 3000
- **Health Check**: http://localhost:3000

### IIS Configuration
- **Site Name**: ExcelAddin
- **Port**: 9443 (HTTPS)
- **Certificate**: server-vs81t.intranet.local
- **Purpose**: Reverse proxy to backend and frontend services

## Troubleshooting

### Service Status Check
```powershell
# Check NSSM services
Get-Service -Name ExcelAddin-Backend
Get-Service -Name ExcelAddin-Frontend

# Check IIS site
Get-IISSite -Name ExcelAddin
```

### Port Conflict Resolution
If the frontend service fails to start due to port conflicts:

```powershell
# Automatically resolve port conflicts (with confirmation)
.\scripts\kill-port-3000.ps1

# Force kill without confirmation  
.\scripts\kill-port-3000.ps1 -Force
```

The deployment scripts automatically detect and resolve port conflicts, but you can use the above script for manual resolution.

### Log Locations
- **Backend Logs**: C:\Logs\ExcelAddin\backend-stdout.log, C:\Logs\ExcelAddin\backend-stderr.log
- **Frontend Logs**: C:\Logs\ExcelAddin\frontend-stdout.log, C:\Logs\ExcelAddin\frontend-stderr.log
- **IIS Logs**: Default IIS log location

### Common Issues
1. **Port Conflicts**: Use `.\scripts\kill-port-3000.ps1` to resolve port 3000 conflicts
2. **404 Errors**: Ensure `dist` directory exists and contains `index.html` (run `npm run build:staging`)
3. **SSL Certificate**: Verify certificate is properly configured in IIS
4. **Firewall**: Ensure Windows Firewall allows traffic on port 9443
5. **Service Dependencies**: Backend and Frontend must be running before IIS can proxy requests

### Enhanced Diagnostics
The deployment scripts now provide detailed diagnostics for common issues:
- Automatic port conflict detection and resolution
- Static file directory validation
- Service log analysis for startup failures
- HTTP response testing with detailed error reporting

## Files Overview

- `deploy-all.ps1` - Complete initial deployment
- `update-all.ps1` - Update existing deployment
- `deploy-backend.ps1` - Deploy backend service only
- `deploy-frontend.ps1` - Deploy frontend service only
- `test-deployment.ps1` - Comprehensive testing suite
- `config/` - Configuration files for services
- `scripts/` - Helper scripts and utilities