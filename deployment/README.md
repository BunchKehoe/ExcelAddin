# ExcelAddin Deployment Guide 

This guide covers the deployment of the Excel Add-in using Vite build system with service-based architecture on Windows Server.

## Overview

**Deployment Architecture**:
- **Build System**: Vite for fast, optimized builds
- **Frontend Service**: Express server via PM2/node-windows Windows service  
- **Backend Service**: Python Flask via NSSM Windows service
- **Reverse Proxy**: IIS on port 9443 with SSL termination
- **Target Platform**: Windows Server 2016+

## Prerequisites

### System Requirements
- **Windows Server 2016+** (Windows 10 also supported)
- **Node.js 18+** 
- **Python 3.8+**
- **IIS** with URL Rewrite module and FastCGI support
- **NSSM** (Non-Sucking Service Manager) - automatically installed
- **Enterprise SSL certificates** for production

### Network Requirements
- **Port 9443**: HTTPS public access for Excel clients
- **Port 3000**: Internal frontend service (localhost only)
- **Port 5000**: Internal backend service (localhost only)
- **Firewall**: Allow inbound HTTPS on port 9443

## Deployment Scripts

The deployment system provides four streamlined PowerShell scripts:

| Script | Purpose | When to Use |
|--------|---------|-------------|
| **deploy-backend.ps1** | Deploy Python Flask backend service | Initial deployment or backend updates |
| **deploy-frontend.ps1** | Deploy React frontend service | Initial deployment or frontend updates |
| **deploy-iis.ps1** | Configure IIS reverse proxy | Initial deployment or IIS configuration changes |
| **troubleshooting.ps1** | Comprehensive diagnostics and fixes | When issues occur or for health checks |

### Quick Start Deployment

**1. Complete Deployment (Recommended)**
```powershell
# Run as Administrator in PowerShell
cd deployment

# Deploy backend service
.\deploy-backend.ps1 -Environment staging

# Deploy frontend service  
.\deploy-frontend.ps1 -Environment staging

# Configure IIS reverse proxy
.\deploy-iis.ps1

# Verify deployment
.\troubleshooting.ps1 -TestAll
```

**2. Environment-Specific Deployment**
```powershell
# For staging environment
.\deploy-backend.ps1 -Environment staging
.\deploy-frontend.ps1 -Environment staging

# For production environment  
.\deploy-backend.ps1 -Environment production
.\deploy-frontend.ps1 -Environment production
```

## Architecture

```
Excel Client → IIS (9443) → Frontend Service (3000)
                        ↘ Backend Service (5000)
```

**Service Details**:
- **Backend Service**: `ExcelAddin-Backend` (NSSM managed)
  - Port: 5000 (internal only)
  - Technology: Python Flask
  - Endpoints: `/api/*`
  - Process: `python app.py`

- **Frontend Service**: `ExcelAddin Frontend` (PM2/node-windows managed)
  - Port: 3000 (internal only)
  - Technology: Express.js serving Vite build
  - Endpoints: `/excellence/*`, static files
  - Process: `node server.js`

- **IIS Reverse Proxy**: 
  - Port: 9443 (HTTPS public)
  - SSL termination and routing
  - Routes requests to appropriate internal services
  - Enterprise certificate support

## Detailed Deployment Procedures

### Backend Service Deployment (deploy-backend.ps1)

**What it does**:
1. Installs Python dependencies via Poetry
2. Creates/updates NSSM Windows service
3. Configures service to start automatically
4. Sets up logging and error handling
5. Applies environment-specific configuration

**Options**:
```powershell
# Basic deployment
.\deploy-backend.ps1

# Specify environment 
.\deploy-backend.ps1 -Environment production

# Force restart service
.\deploy-backend.ps1 -Force

# Enable verbose logging
.\deploy-backend.ps1 -Verbose
```

**Service Configuration**:
- **Service Name**: `ExcelAddin-Backend`
- **Start Type**: Automatic (delayed start)
- **Recovery**: Restart on failure (3 attempts)
- **Logging**: Windows Event Log + file logging
- **Working Directory**: `C:\inetpub\wwwroot\ExcelAddin\backend`

### Frontend Service Deployment (deploy-frontend.ps1)

**What it does**:
1. Builds optimized Vite bundle for target environment
2. Creates Express.js server for static file serving
3. Installs/updates Windows service via node-windows
4. Configures service for automatic startup
5. Sets up request logging and error handling

**Options**:
```powershell
# Basic deployment
.\deploy-frontend.ps1

# Specify environment
.\deploy-frontend.ps1 -Environment staging

# Skip build step (use existing dist/)
.\deploy-frontend.ps1 -SkipBuild

# Enable development mode logging
.\deploy-frontend.ps1 -Debug
```

**Service Configuration**:
- **Service Name**: `ExcelAddin Frontend`  
- **Start Type**: Automatic
- **Recovery**: Restart on failure
- **Logging**: Console + file logging
- **Working Directory**: `C:\inetpub\wwwroot\ExcelAddin`

### IIS Proxy Configuration (deploy-iis.ps1)

**What it does**:
1. Creates IIS application pool for ExcelAddin
2. Configures website on port 9443 with SSL
3. Sets up URL rewrite rules for frontend/backend routing
4. Applies security headers and performance optimizations
5. Configures SSL certificate binding

**Key IIS Configuration**:
```xml
<!-- URL Rewrite Rules -->
<rule name="API Proxy" stopProcessing="true">
    <match url="^excellence/api/(.*)" />
    <action type="Rewrite" url="http://localhost:5000/api/{R:1}" />
</rule>

<rule name="Frontend Proxy" stopProcessing="true">  
    <match url="^excellence/(.*)" />
    <action type="Rewrite" url="http://localhost:3000/{R:1}" />
</rule>
```

**SSL Configuration**:
- **Port**: 9443 (HTTPS)
- **Certificate**: Enterprise CA certificate
- **Security**: TLS 1.2+ enforced
- **Headers**: HSTS, CSP, X-Frame-Options configured

## Environment Configuration

### Build Configuration

**Environment Variables**:
```bash
# Staging
VITE_API_BASE_URL=https://server-vs81t.intranet.local:9443/excellence/api
VITE_ENVIRONMENT=staging

# Production  
VITE_API_BASE_URL=https://server-vs84.intranet.local:9443/excellence/api
VITE_ENVIRONMENT=production
```

**Build Commands**:
```bash
# Development build (with source maps)
npm run build:dev

# Staging build (optimized)
npm run build:staging

# Production build (fully optimized)  
npm run build:prod

# Development server with hot reload
npm run dev
```

**Vite Performance Benefits**:
- **Bundle Size**: 346KB (vs 836KB with webpack - 58% reduction)
- **Dependencies**: 156 packages (vs 577 - 73% reduction)
- **Build Time**: 3.4s (vs 14.6s - 77% faster)
- **Development**: Better HTTPS support, faster HMR

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

## Service Management

### IIS Proxy Management
The IIS proxy can be deployed and managed independently:

```powershell
# Deploy IIS proxy only
.\deploy-iis-proxy.ps1

# Deploy with custom settings
.\deploy-iis-proxy.ps1 -SiteName "MyExcelProxy" -Port 8443 -Force

# Management commands
Start-Website -Name "ExcelAddin-Proxy"
Stop-Website -Name "ExcelAddin-Proxy" 
Get-Website -Name "ExcelAddin-Proxy" | Select Name, State, PhysicalPath

# Check proxy status
Invoke-WebRequest https://localhost:9443/ -UseBasicParsing
```

Key features of the IIS proxy deployment:
- **Independent deployment**: Can be deployed separately from frontend/backend services
- **Advanced URL rewriting**: Handles all Excel Add-in endpoints with proper forwarding
- **CORS support**: Configured for Excel Add-in compatibility
- **HTTPS support**: Automatic SSL certificate detection and binding
- **Health monitoring**: Built-in status page and service health checks
- **Comprehensive logging**: Detailed deployment and runtime logging

### Frontend Service Commands
```powershell
# Standard Windows service commands
Start-Service "ExcelAddin Frontend"
Stop-Service "ExcelAddin Frontend"
Get-Service "ExcelAddin Frontend"
Restart-Service "ExcelAddin Frontend"

# Advanced management using node-windows
node service.cjs start
node service.cjs stop
node service.cjs restart
node service.cjs install    # Reinstall service
node service.cjs uninstall  # Remove service
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

### Service Configuration
- **Backend Service (NSSM)**: Auto-start, automatic restart on failure, environment-aware configuration (development/staging/production), logging to C:\Logs\ExcelAddin\
- **Frontend Service (node-windows)**: Auto-start Windows service, automatic restart on failure, environment-aware configuration, logs to Windows Event Log
- **Environment Support**: Both services now support environment-specific configuration files (.env.development, .env.staging, .env.production)

## Troubleshooting

### Common Issues

1. **Service won't start**
   ```powershell
   # Check prerequisites
   node --version  # Should be 18+
   python --version  # Should be 3.8+
   nssm --version   # Should exist (backend only)
   
   # Check service status
   Get-Service "ExcelAddin Frontend"
   Get-Service "ExcelAddin-Backend"
   
   # Check logs
   Get-WinEvent -LogName Application -Source "ExcelAddin Frontend" -MaxEvents 20
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
- Check backend service status and frontend service status
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
# Frontend service management
Get-Service "ExcelAddin Frontend" | Format-Table
Start-Service "ExcelAddin Frontend"
Stop-Service "ExcelAddin Frontend"
Restart-Service "ExcelAddin Frontend"

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
- Frontend: Windows Event Viewer (Application Log, Source: "ExcelAddin Frontend")
- Backend: `C:\Logs\ExcelAddin\backend-stderr.log`
- IIS: Windows Event Logs (System, Application)

This deployment system provides a robust, maintainable foundation for the Excel Add-in with modern build tooling, reliable node-windows service hosting, and comprehensive monitoring.