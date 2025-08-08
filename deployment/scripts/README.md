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

### setup-backend-service.ps1
**Purpose:** Install Flask backend as Windows service  
**Usage:** `.\setup-backend-service.ps1`  
**Requirements:** Run as Administrator, NSSM available

### add-firewall-rule.ps1
**Purpose:** Open Windows Firewall for port 9443  
**Usage:** `.\add-firewall-rule.ps1`  
**Requirements:** Run as Administrator

## Quick Deployment

For existing IIS servers:

```powershell
# 1. Setup (run once)
.\deploy-to-existing-iis.ps1

# 2. Deploy app
.\build-and-deploy-iis.ps1

# 3. Test
.\test-iis-simple.ps1
```