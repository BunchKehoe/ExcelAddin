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
**Purpose:** Install Flask backend directly in IIS using FastCGI  
**Usage:** `.\setup-backend-iis.ps1`  
**Requirements:** Run as Administrator, IIS with FastCGI support

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
# 1. Setup (run once)
.\deploy-to-existing-iis.ps1

# 2. Setup backend in IIS
.\setup-backend-iis.ps1

# 3. Deploy app  
.\build-and-deploy-iis.ps1

# 4. Test
.\test-iis-simple.ps1
```