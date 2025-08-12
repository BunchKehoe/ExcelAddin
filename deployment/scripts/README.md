# IIS Deployment Scripts

Essential scripts for deploying the Excel Add-in to Windows IIS.

## Scripts

### deploy-to-existing-iis.ps1
**Purpose:** One-time setup of IIS site and application pool  
**Usage:** `.\deploy-to-existing-iis.ps1`  
**Requirements:** Run as Administrator on server with IIS already installed

### build-and-deploy-iis.ps1  
**Purpose:** Build React app and deploy to IIS directory  
**Usage:** `.\build-and-deploy-iis.ps1`  
**Requirements:** Node.js/npm installed, IIS site already configured

### test-iis-simple.ps1
**Purpose:** Test that IIS deployment is working  
**Usage:** `.\test-iis-simple.ps1`  
**Requirements:** None (read-only testing)

### setup-backend-iis.ps1
**Purpose:** Deploy and configure Flask backend as IIS application using FastCGI  
**Usage:** `.\setup-backend-iis.ps1 -SiteName "ExcelAddin"`  
**Requirements:** Run as Administrator, Python installed, IIS site already configured
**Creates:** IIS application at `/excellence/backend/` with FastCGI Python integration

### test-backend-integration.ps1
**Purpose:** Test that backend IIS integration is working correctly  
**Usage:** `.\test-backend-integration.ps1 -SiteName "ExcelAddin"`  
**Requirements:** Backend should be already set up with setup-backend-iis.ps1

### setup-backend-service.ps1 (DEPRECATED)
**Purpose:** ~~Install Flask backend as Windows service~~ **DEPRECATED - Use setup-backend-iis.ps1 instead**  
**Usage:** ~~`.\setup-backend-service.ps1`~~ **Redirects to setup-backend-iis.ps1**  
**Requirements:** ~~Run as Administrator, NSSM available~~ **No longer required**

### add-firewall-rule.ps1
**Purpose:** Open Windows Firewall for port 9443  
**Usage:** `.\add-firewall-rule.ps1`  
**Requirements:** Run as Administrator

## Quick Deployment

For existing IIS servers:

```powershell
# 1. Initial IIS setup (run once)
.\deploy-to-existing-iis.ps1

# 2. Complete deployment with frontend and backend
.\build-and-deploy-iis.ps1 -SiteName "ExcelAddin"
# Note: This automatically runs setup-backend-iis.ps1 as part of the process

# 3. Test the deployment
.\test-backend-integration.ps1 -SiteName "ExcelAddin"
.\test-iis-simple.ps1
```

## Architecture

**New Unified IIS Architecture:**
- Frontend: IIS serves React app at `/excellence/`
- Backend: IIS serves Flask app at `/excellence/backend/` via FastCGI
- API Routing: `/excellence/api/*` routes to `/excellence/backend/api/*`
- **No separate services** - everything runs in IIS

**Benefits:**
- Single server setup (no NSSM, no nginx, no port 5000 service)
- Simplified deployment and management
- Standard IIS application pool management
- Better security and performance integration