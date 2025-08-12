# ExcelAddin Deployment Guide

This deployment system hosts the ExcelAddin service with a clean separation of concerns:
- **IIS**: Reverse proxy only (HTTPS termination and request forwarding)
- **Backend**: Python Flask API hosted via NSSM service
- **Frontend**: React application hosted via PM2 service

## Server Configuration
- **Service Name**: ExcelAddin
- **Public Address**: https://server-vs81t.intranet.local:9443
- **Target Environment**: Windows 10

## Architecture Overview

```
Excel Client → IIS Proxy (Port 9443) → Backend Service (Port 5000)
                                   ↘ Frontend Service (Port 3000)
```

## Prerequisites

1. **Node.js and npm** installed
2. **Python 3.8+** with pip installed
3. **PM2** globally installed: `npm install -g pm2`
4. **NSSM** downloaded and available in PATH
5. **IIS** with URL Rewrite module installed
6. **SSL Certificate** configured for the domain

## Deployment Scripts

### Initial Deployment
```powershell
.\deploy-all.ps1
```
Performs complete first-time deployment including service installation and configuration.

### Updates After First Run
```powershell
.\update-all.ps1
```
Updates services without reinstalling or reconfiguring infrastructure.

### Individual Service Deployment
```powershell
.\deploy-backend.ps1   # Deploy Python backend via NSSM
.\deploy-frontend.ps1  # Deploy React frontend via PM2
```

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

### Frontend Service (PM2)
- **Application Name**: exceladdin-frontend
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
# Check NSSM service
nssm status ExcelAddin-Backend

# Check PM2 service
pm2 status exceladdin-frontend

# Check IIS site
Get-IISSite -Name ExcelAddin
```

### Log Locations
- **Backend Logs**: Check NSSM service logs
- **Frontend Logs**: `pm2 logs exceladdin-frontend`
- **IIS Logs**: Default IIS log location

### Common Issues
1. **Port Conflicts**: Ensure ports 3000 and 5000 are available
2. **SSL Certificate**: Verify certificate is properly configured in IIS
3. **Firewall**: Ensure Windows Firewall allows traffic on port 9443
4. **Service Dependencies**: Backend and Frontend must be running before IIS can proxy requests

## Files Overview

- `deploy-all.ps1` - Complete initial deployment
- `update-all.ps1` - Update existing deployment
- `deploy-backend.ps1` - Deploy backend service only
- `deploy-frontend.ps1` - Deploy frontend service only
- `test-deployment.ps1` - Comprehensive testing suite
- `config/` - Configuration files for services
- `scripts/` - Helper scripts and utilities