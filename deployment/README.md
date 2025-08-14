# ExcelAddin Deployment Guide (Vite + Task Scheduler)

This guide covers the deployment of the Excel Add-in using the new Vite-based build system with Windows Task Scheduler for frontend and NSSM for backend services on Windows Server 10.

## Overview

The Excel Add-in now uses:
- **Build System**: Vite (replaced webpack)
- **Frontend Hosting**: Express server via Windows Task Scheduler
- **Backend Hosting**: Python Flask via NSSM  
- **Reverse Proxy**: IIS on port 9443
- **Target Platform**: Windows Server 10

## Quick Deployment

### Prerequisites
- Windows Server 10
- Node.js 18+ 
- Python 3.8+
- NSSM (Non-Sucking Service Manager) - for backend only
- IIS with URL Rewrite module
- Windows Task Scheduler (built-in)

### 1. Complete Deployment (Recommended)
```powershell
# Run as Administrator
cd deployment
.\deploy-all.ps1 -Environment staging
```

### 2. Individual Service Deployment
```powershell
# Backend only
.\deploy-backend.ps1

# Frontend only  
.\deploy-frontend.ps1 -Environment staging

# IIS configuration
.\configure-iis.ps1
```

### 3. Testing and Troubleshooting
```powershell
# Comprehensive diagnostics
.\debug-integration.ps1 -Detailed

# Fix issues automatically
.\debug-integration.ps1 -FixIssues
```

## Architecture

```
Excel Client → IIS (9443) → Frontend Service (3000)
                        ↘ Backend Service (5000)
```

### Service Details
- **Backend Service**: `ExcelAddin-Backend` (NSSM)
  - Port: 5000
  - Technology: Python Flask
  - Endpoints: `/api/*`

- **Frontend Task**: `ExcelAddin-Frontend` (Task Scheduler)
  - Port: 3000  
  - Technology: Express.js serving Vite build
  - Endpoints: `/excellence/*`, static files

- **IIS Proxy**: `ExcelAddin` site
  - Port: 9443 (HTTPS)
  - Routes to appropriate services

## Build System Changes

### Vite vs Webpack Benefits
- **Bundle Size**: 346KB vs 836KB (58% reduction)
- **Dependencies**: 156 vs 577 packages (73% reduction)
- **Build Time**: 3.4s vs 14.6s (77% faster)
- **Development**: Better HTTPS support, faster HMR

### Build Commands
```bash
# Development
npm run dev

# Staging build
npm run build:staging

# Production build  
npm run build:prod

# Type checking
npm run lint
```

## Excel Add-in Compatibility

### Required Files (Maintained)
- `taskpane.html` - Main Excel interface
- `commands.html` - Custom functions
- `functions.json` - Custom functions manifest
- `manifest*.xml` - Add-in manifests
- `assets/` - Icons and static files

### URL Structure  
- **Development**: `https://localhost:3000/`
- **Staging**: `https://server:9443/excellence/`
- **Production**: `https://server:9443/excellence/`

### Excel Endpoints Tested
- `/excellence/taskpane.html` - Main taskpane
- `/excellence/commands.html` - Custom functions
- `/functions.json` - Functions manifest
- `/manifest.xml` - Add-in manifest  
- `/api/health` - Backend health check

## Service/Task Management

### Frontend Task Commands
```powershell
# Using the management script
.\manage-frontend-task.ps1 -Action status
.\manage-frontend-task.ps1 -Action start
.\manage-frontend-task.ps1 -Action stop  
.\manage-frontend-task.ps1 -Action restart

# Direct Task Scheduler commands
schtasks /run /tn ExcelAddin-Frontend      # Start
schtasks /end /tn ExcelAddin-Frontend      # Stop
schtasks /query /tn ExcelAddin-Frontend /fo LIST  # Status
```

### Backend Service Commands
```powershell
# Standard Windows service commands
Start-Service ExcelAddin-Backend
Stop-Service ExcelAddin-Backend
Get-Service ExcelAddin-Backend

# View logs
Get-Content C:\Logs\ExcelAddin\*-stderr.log -Tail 20
```

### Service/Task Configuration
- **Backend Service (NSSM)**: Auto-start, automatic restart on failure, logging to C:\Logs\ExcelAddin\
- **Frontend Task (Scheduler)**: Runs continuously, restarts on failure, auto-start on boot
- **Environment**: Production settings via environment variables and task configuration

## Troubleshooting

### Common Issues

1. **Task/Service won't start**
   ```powershell
   # Check prerequisites
   node --version  # Should be 18+
   python --version  # Should be 3.8+
   nssm --version   # Should exist (backend only)
   
   # Check task status
   .\manage-frontend-task.ps1 -Action status -Detailed
   
   # Check logs
   Get-Content C:\Logs\ExcelAddin\*-stderr.log -Tail 50
   ```

2. **Port conflicts**
   ```powershell
   # Check what's using ports
   Get-NetTCPConnection -LocalPort 3000,5000,9443 | 
     ForEach-Object { Get-Process -Id $_.OwningProcess }
   ```

3. **Build failures**
   ```powershell
   # Clean build
   npm run clean
   npm install
   npm run build:staging
   ```

4. **Excel Add-in not loading**
   - Verify manifest URLs point to correct server
   - Check HTTPS certificates are valid
   - Ensure Office.js is accessible
   - Test taskpane.html loads correctly

### Debug Script
The debug script tests all critical endpoints:
```powershell
.\debug-integration.ps1 -Detailed -FixIssues
```

This will:
- Check backend service status and frontend task status
- Test port connectivity  
- Verify Excel-specific endpoints
- Validate Office.js references
- Test IIS proxy routing
- Analyze log files
- Provide recommendations

## Development Workflow

### Local Development
```bash
# Start dev server with HTTPS
npm run dev

# In another terminal, start backend
cd backend
python run.py
```

### Testing Changes
```bash
# Build and test
npm run build:staging
npm run preview

# Deploy to staging
.\deployment\deploy-all.ps1 -Environment staging
```

### Production Deployment
```bash
# Build production
npm run build:prod

# Deploy to production  
.\deployment\deploy-all.ps1 -Environment production
```

## Security Considerations

- HTTPS required for Excel Add-ins
- CORS headers configured for Excel domains
- Services run with minimal privileges
- Log files have appropriate permissions
- IIS configured with security headers

## Performance

### Optimizations Applied
- Code splitting for faster loading
- Asset optimization and compression
- Efficient chunk loading
- Browser caching headers
- Minimal JavaScript bundles

### Monitoring
- Health check endpoints for both services
- Comprehensive logging
- Performance metrics in debug script
- Service restart on failure

## Support

### Useful Commands
```powershell
# Frontend task management
.\manage-frontend-task.ps1 -Action status
schtasks /query /tn ExcelAddin-Frontend /fo LIST

# Backend service management
Get-Service ExcelAddin-Backend | Format-Table
Restart-Service ExcelAddin-Backend

# IIS management  
Get-Website ExcelAddin
Start-Website ExcelAddin

# Testing
Invoke-WebRequest http://localhost:3000/health
Invoke-WebRequest http://localhost:5000/api/health
```

### Log Locations
- Frontend: Console output via Task Scheduler (check Windows Event Viewer)
- Backend: `C:\Logs\ExcelAddin\backend-stderr.log`
- IIS: Windows Event Logs (System, Application)

This deployment system provides a robust, maintainable foundation for the Excel Add-in with modern build tooling, reliable Windows Task Scheduler frontend hosting, and comprehensive monitoring.