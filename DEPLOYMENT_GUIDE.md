# PrimeExcelence Excel Add-in - Deployment Guide

## Table of Contents
1. [Deployment Architecture](#deployment-architecture)
2. [Local Development Setup](#local-development-setup)
3. [Windows Server Production Deployment](#windows-server-production-deployment)
4. [Step-by-Step Deployment Process](#step-by-step-deployment-process)
5. [Configuration Files](#configuration-files)
6. [SSL Certificate Management](#ssl-certificate-management)
7. [Service Management](#service-management)

## Deployment Architecture

### Local Development Architecture
```
Excel Desktop/Online
        │
        └── https://localhost:3000 (Development Server)
                    │
                    └── http://localhost:5000 (Backend API)
```

### Windows Server Production Architecture
```
Excel Desktop/Online
        │
        └── https://server-vs81t.intranet.local:9443/excellence
                    │
                    ▼
        ┌─────────────────────────────────────────────┐
        │             Windows IIS (Port 9443)        │
        │                                             │
        │  • SSL with server-vs81t certificates      │
        │  • Static File Serving (/excellence/)      │
        │    - taskpane.html, commands.html          │
        │    - JavaScript bundles, CSS               │
        │    - Images, manifest.xml                  │
        │  • API Proxy (/excellence/api/ → :5000)    │
        │  • URL Rewrite + ARR for routing          │
        │  • Native Windows Service                  │
        └─────────────────────────────────────────────┘
                    │
                    └── Backend Flask App (Port 5000)
                              │
                              └── Windows Service (via NSSM)
```

**How Frontend Hosting Works:**
- IIS serves all frontend files (HTML, JS, CSS, images) as static content
- Files are served from `C:\inetpub\wwwroot\ExcelAddin\excellence\` directory  
- URL `https://server-vs81t.intranet.local:9443/excellence/` loads `excellence/taskpane.html`
- Excel add-in loads in taskpane, makes API calls to `/excellence/api/` endpoints
- IIS URL Rewrite with ARR proxies API calls to backend Flask app running on port 5000

## Frontend Deployment Overview

The Excel add-in frontend is a **React-based single-page application** that gets compiled into static files and served by IIS. Here's how it works:

### 1. Frontend Build Process
```bash
# Development builds (for testing)
npm run build:dev          # Development build with source maps
npm run start              # Development server on https://localhost:3000

# Production builds (for deployment)  
npm run build:staging      # Production build configured for staging server
npm run build:prod         # Production build configured for production server
```

### 2. Frontend Architecture
- **Entry Points**: `taskpane.tsx` (main interface), `commands.ts` (ribbon commands)
- **Build Output**: Static HTML, JavaScript bundles, CSS, assets
- **Hosting**: IIS serves files from `C:\inetpub\wwwroot\ExcelAddin\excellence\`
- **Public Path**: All resources served under `/excellence/` subpath

### 3. Integration with Excel
- **Manifest File**: `manifest.xml` defines add-in metadata and entry points
- **Taskpane**: Main interface loads at `https://server:9443/excellence/taskpane.html`  
- **Commands**: Ribbon commands load at `https://server:9443/excellence/commands.html`
- **API Communication**: Frontend makes AJAX calls to backend API

### 4. IIS Configuration for Frontend
The IIS server is configured to:
- Serve static files from the application directory
- Handle URL routing for single-page app behavior using URL Rewrite module
- Proxy API requests to the backend Flask application using Application Request Routing (ARR)
- Provide SSL termination for secure HTTPS access
- Native Windows service integration without third-party tools

## Local Development Setup

### 1. Prerequisites Installation

```bash
# Install Node.js (v16+)
# Download from https://nodejs.org/

# Verify installation
node --version
npm --version

# Install Python (3.8+)
# Download from https://python.org/

# Install Office development certificates
npm install -g office-addin-dev-certs
office-addin-dev-certs install
```

### 2. Repository Setup

```bash
# Clone repository
git clone <repository-url>
cd ExcelAddin

# Install frontend dependencies
npm install

# Install backend dependencies  
cd backend
pip install -r requirements.txt
cd ..
```

### 3. Start Development Services

```bash
# Terminal 1 - Start backend
cd backend
python run.py

# Terminal 2 - Start frontend
npm start
```

### 4. Load Add-in in Excel

1. Open Microsoft Excel
2. Go to **Developer** tab (enable in Options if needed)
3. Click **Add-ins** → **My Add-ins**  
4. Choose **Upload My Add-in**
5. Select `manifest.xml` from project root
6. Click **PrimeExcelence** button in Home tab

## Windows Server Production Deployment  

### Prerequisites for Windows Server

1. **Windows Server 2016+** with Administrator access
2. **nginx for Windows** - Download from http://nginx.org/en/download.html
3. **Python 3.8+** - Download from https://python.org/downloads/windows/
4. **NSSM** (Non-Sucking Service Manager) - Download from https://nssm.cc/
5. **PowerShell 5.1+** (usually pre-installed)
6. **SSL Certificates** from your organization's Certificate Authority

### Installation Steps

#### 1. Install Core Components

```powershell
# Create application directory
mkdir C:\inetpub\wwwroot\ExcelAddin
cd C:\inetpub\wwwroot\ExcelAddin

# Extract nginx to C:\nginx
# Extract NSSM to C:\Tools\nssm (add to PATH)
# Install Python with "Add to PATH" option checked
```

#### 2. Deploy Application Files

```bash
# On development machine - Build production files
npm install
npm run build:staging

# Copy files to server:
# - Copy dist/* to C:\inetpub\wwwroot\ExcelAddin\
# - Copy backend/* to C:\inetpub\wwwroot\ExcelAddin\backend\
# - Copy deployment/* to C:\inetpub\wwwroot\ExcelAddin\deployment\
```

#### 3. Install Backend Dependencies

```powershell
# On Windows server
cd C:\inetpub\wwwroot\ExcelAddin\backend
pip install -r requirements.txt
```

#### 4. Build and Deploy Frontend

The frontend is a React-based Excel add-in that gets built into static files and served by nginx.

```bash
# On development machine - Build frontend for staging
npm install
npm run build:staging

# This creates a 'dist' directory containing:
# - taskpane.html (main add-in interface)
# - commands.html (ribbon commands)
# - JavaScript bundles with content hashing
# - functions.json (custom functions metadata)
# - assets/ (images, icons, etc.)
# - manifest.xml (copied from manifest-staging.xml)
```

```powershell
# Copy built frontend files to Windows server
# Copy contents of dist/* to C:\inetpub\wwwroot\ExcelAddin\dist\

# Create the dist directory on the server
New-Item -ItemType Directory -Path "C:\inetpub\wwwroot\ExcelAddin\dist" -Force

# Your server directory should look like:
# C:\inetpub\wwwroot\ExcelAddin\
# ├── dist/                    # Frontend files served by nginx
# │   ├── taskpane.html
# │   ├── commands.html
# │   ├── taskpane.[hash].js
# │   ├── commands.[hash].js
# │   ├── vendors.[hash].js
# │   ├── functions.json
# │   ├── assets/
# │   └── manifest.xml
# └── backend/                 # Backend API served by Flask
#     ├── app.py
#     ├── requirements.txt
#     └── ...
```

## Step-by-Step Deployment Process

### Phase 1: SSL Certificate Setup

#### Option A: Using Existing Company Certificates

```powershell
# Copy certificates to C:\Cert\
# Required files:
# - server.crt (or server.pfx)
# - server.key (if using .crt)  
# - cacert.pem (company root CA)

# If using .pfx file, extract key and certificate:
cd C:\inetpub\wwwroot\ExcelAddin\deployment\scripts
.\extract-pfx.ps1 -PfxPath "C:\Cert\server.pfx" -OutputPath "C:\Cert"
```

#### Option B: Handle Encrypted Private Keys

```powershell
# If private key is encrypted (prompts for password)
cd C:\inetpub\wwwroot\ExcelAddin\deployment\scripts
.\handle-encrypted-key.ps1 -KeyPath "C:\Cert\server.key" -OutputPath "C:\Cert"
```

### Phase 2: Configure IIS

#### Option A: New IIS Installation

```powershell
# Run as Administrator
cd C:\inetpub\wwwroot\ExcelAddin\deployment\scripts

# Install and configure IIS with required modules
.\setup-iis.ps1 -Force

# This script will:
# - Install IIS with URL Rewrite and ARR modules
# - Create ExcelAddin website on port 9443
# - Configure SSL with server-vs81t certificates
# - Set up web.config for static files and API proxying
# - Configure Windows Firewall
```

#### Option B: Existing IIS Server (Recommended)

If you already have IIS installed and running, use the lightweight deployment script:

```powershell
# Run as Administrator
cd C:\inetpub\wwwroot\ExcelAddin\deployment\scripts

# Deploy to existing IIS server
.\deploy-to-existing-iis.ps1

# Or with custom parameters:
.\deploy-to-existing-iis.ps1 -Force -SiteName "MyExcelApp" -Port 8443

# This script will:
# - Create application pool and website for Excel Add-in
# - Configure SSL if certificates are available
# - Set up web.config and permissions
# - Skip IIS feature installation (works with existing IIS)
```

**Use Option B if you get errors about IIS already running or features already installed.**

### Phase 3: Set Up Backend Service

```powershell
# Run as Administrator
cd C:\inetpub\wwwroot\ExcelAddin\deployment\scripts

# Install and configure backend service
.\setup-backend-service.ps1 -Force

# Verify service installation
Get-Service ExcelAddinBackend
```

### Phase 4: Start Services

```powershell
# Start backend service first
Start-Service ExcelAddinBackend

# Wait 10 seconds for backend to initialize
Start-Sleep -Seconds 10

# Start IIS website (IIS service W3SVC starts automatically)
Start-Website -Name "ExcelAddin"

# Verify services are running
Get-Service ExcelAddinBackend, W3SVC
Get-Website -Name "ExcelAddin"
```

### Phase 5: Validate Deployment

```powershell
# Test backend API directly
Invoke-WebRequest -Uri "http://localhost:5000/api/health" -UseBasicParsing

# Test through IIS proxy - Choose method based on your PowerShell version
# Check PowerShell version first
$PSVersionTable.PSVersion

# For PowerShell 6.0+ (PowerShell Core)
if ($PSVersionTable.PSVersion.Major -ge 6) {
    Invoke-WebRequest -Uri "https://localhost:9443/excellence/api/health" -SkipCertificateCheck -UseBasicParsing
    Invoke-WebRequest -Uri "https://localhost:9443/excellence/taskpane.html" -SkipCertificateCheck -UseBasicParsing
} else {
    # For Windows PowerShell 5.1 and earlier - use .NET classes to ignore SSL
    Write-Host "Using Windows PowerShell - ignoring SSL certificates for testing"
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    
    try {
        Invoke-WebRequest -Uri "https://localhost:9443/excellence/api/health" -UseBasicParsing
        Invoke-WebRequest -Uri "https://localhost:9443/excellence/taskpane.html" -UseBasicParsing
        Write-Host "✅ HTTPS endpoints are accessible"
    } catch {
        Write-Host "❌ HTTPS connection failed: $($_.Exception.Message)"
        Write-Host "💡 Check troubleshooting section below"
    }
    
    # Reset certificate validation for security
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
}

# Alternative: Use curl.exe directly (if Git for Windows is installed)
if (Get-Command curl.exe -ErrorAction SilentlyContinue) {
    Write-Host "Testing with curl.exe..."
    curl.exe -k -s -o /dev/null -w "%{http_code}" https://localhost:9443/excellence/api/health
    curl.exe -k -s -o /dev/null -w "%{http_code}" https://localhost:9443/excellence/taskpane.html
}
```

### Phase 7: Deploy to Excel Users

1. **Update manifest file**:
   - Copy `manifest-staging.xml` to network location or SharePoint
   - Or email to users for manual installation

2. **User installation**:
   - Excel → Developer → Add-ins → My Add-ins → Upload
   - Select the manifest-staging.xml file
   - Add-in appears in Home tab

## Configuration Files

### IIS Configuration (`deployment/iis/web.config`)

**IIS replaces nginx** - This configuration provides the same functionality with native Windows IIS:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<configuration>
  <system.webServer>
    <defaultDocument>
      <files>
        <add value="taskpane.html" />
      </files>
    </defaultDocument>

    <!-- URL Rewrite rules for API proxy and SPA routing -->
    <rewrite>
      <rules>
        <!-- API Proxy Rule: Forward /excellence/api/ to Flask backend -->
        <rule name="API Proxy" stopProcessing="true">
          <match url="^excellence/api/(.*)$" />
          <action type="Rewrite" url="http://127.0.0.1:5000/api/{R:1}" />
        </rule>

        <!-- Health check endpoint -->
        <rule name="Health Check" stopProcessing="true">
          <match url="^health$" />
          <action type="CustomResponse" statusCode="200" statusReason="OK" />
        </rule>

        <!-- SPA fallback: serve taskpane.html for /excellence/ routes -->
        <rule name="SPA Fallback">
          <match url="^excellence/.*$" />
          <conditions>
            <add input="{REQUEST_FILENAME}" matchType="IsFile" negate="true" />
            <add input="{REQUEST_FILENAME}" matchType="IsDirectory" negate="true" />
          </conditions>
          <action type="Rewrite" url="/excellence/taskpane.html" />
        </rule>
      </rules>
    </rewrite>

    <!-- CORS headers for Excel add-in -->
    <httpProtocol>
      <customHeaders>
        <add name="Access-Control-Allow-Origin" value="*" />
        <add name="Access-Control-Allow-Methods" value="GET, POST, PUT, DELETE, OPTIONS" />
        <add name="Access-Control-Allow-Headers" value="Content-Type, Authorization" />
      </customHeaders>
    </httpProtocol>
  </system.webServer>
</configuration>
```

**Key Features of the IIS Configuration:**
- **Native Windows integration** - No third-party software required
- **URL Rewrite module** - Handles API proxying and SPA routing
- **Application Request Routing** - Proxies API calls to Flask backend
- **SSL termination** - Configured through IIS Manager or PowerShell
- **Same functionality** - Replaces nginx with equivalent IIS features

## Migrating from nginx to IIS

If you previously used nginx and want to switch to IIS (e.g., due to antivirus blocking nginx):

```powershell
# Run as Administrator
cd C:\inetpub\wwwroot\ExcelAddin\deployment\scripts

# Migrate from nginx to IIS (stops nginx, sets up IIS)
.\migrate-nginx-to-iis.ps1 -Force

# Or set up IIS from scratch
.\setup-iis.ps1 -Force

# Test IIS configuration
.\test-iis-simple.ps1
```

The migration script will:
- Stop and optionally remove nginx service
- Install IIS with required modules (URL Rewrite, ARR)
- Create ExcelAddin website with SSL on port 9443
- Configure web.config for static files and API proxying
- Update firewall rules
- Keep the same URLs and functionality

### Backend Environment (`.env.production`)

```bash
FLASK_ENV=production
FLASK_DEBUG=False
API_BASE_URL=https://server-vs81t.intranet.local:9443/excellence/api
CORS_ORIGINS=https://server-vs81t.intranet.local:9443
DATABASE_CONFIG=database.cfg
```

### Excel Manifest (`manifest-staging.xml`)

```xml
<SourceLocation DefaultValue="https://server-vs81t.intranet.local:9443/excellence/taskpane.html"/>
<SupportUrl DefaultValue="https://server-vs81t.intranet.local:9443/excellence/"/>
<AppDomains>
  <AppDomain>https://server-vs81t.intranet.local:9443</AppDomain>
</AppDomains>
```

## IIS Installation & Testing

### IIS Quick Start

#### For New IIS Installation:
```bash
# Apply library upgrades (if not done)
cd backend
pip install -r requirements.txt --upgrade

# Build frontend
npm run build:staging

# Set up IIS (run as Administrator)
.\deployment\scripts\setup-iis.ps1 -Force

# Test IIS configuration
.\deployment\scripts\test-iis-simple.ps1
```

#### For Existing IIS Server (Recommended):
```bash
# Apply library upgrades (if not done)
cd backend
pip install -r requirements.txt --upgrade

# Build frontend
npm run build:staging

# Deploy to existing IIS (run as Administrator)
.\deployment\scripts\deploy-to-existing-iis.ps1

# Build and deploy application
.\deployment\scripts\build-and-deploy-iis.ps1

# Fix any configuration issues (if needed)
.\deployment\scripts\fix-iis-config.ps1

# Test configuration
.\deployment\scripts\test-iis-simple.ps1
```

# Test IIS configuration
.\deployment\scripts\test-iis-simple.ps1
```

### Manual IIS Module Installation

If the setup script cannot automatically install modules, download and install these manually:

1. **URL Rewrite Module**: https://www.iis.net/downloads/microsoft/url-rewrite
2. **Application Request Routing (ARR)**: https://www.iis.net/downloads/microsoft/application-request-routing

After installing, run the setup script again: `.\setup-iis.ps1 -Force`

## SSL Certificate Management

### Supported Certificate Formats

1. **Separate Files**: `server.crt` + `server.key`
2. **PFX/P12 Files**: `server.pfx` (requires extraction)
3. **Root CA**: `cacert.pem` for certificate chain validation

### Certificate Installation Process

```powershell
# Method 1: Copy existing certificates
copy \\fileserver\certificates\*.* C:\Cert\

# Method 2: Extract from PFX
.\deployment\scripts\extract-pfx.ps1 -PfxPath "C:\Cert\server.pfx"

# Method 3: Convert encrypted keys
.\deployment\scripts\handle-encrypted-key.ps1 -KeyPath "C:\Cert\server.key"
```

### Certificate Validation

```powershell
# Test certificate validity
openssl x509 -in C:\Cert\server.crt -text -noout

# Test private key match
openssl rsa -in C:\Cert\server.key -check

# Test SSL configuration with IIS
Get-Website -Name "ExcelAddin" | Select-Object Name, State, Bindings
Get-WebBinding -Name "ExcelAddin"
```

## Testing and Debugging

### IIS Testing Scripts

**Test IIS Configuration**
```powershell
# Test IIS setup and configuration
.\deployment\scripts\test-iis-simple.ps1

# This script checks:
# - IIS service is running
# - ExcelAddin website exists and is running  
# - Port 9443 is listening
# - Physical path and files exist
# - SSL certificate is configured
# - Health check endpoint responds
# - Frontend taskpane.html is accessible
```

**Test Backend Standalone (Outside NSSM)**
```powershell
# Run Flask backend directly for debugging
.\deployment\scripts\test-backend-standalone.ps1

# This allows you to:
# - Debug backend issues without NSSM service wrapper
# - See Python error messages directly
# - Test backend API endpoints manually
# - Validate Flask configuration
```

**Enhanced Connectivity Diagnostics**
```powershell
# Comprehensive connectivity testing (updated for IIS)
.\deployment\scripts\diagnose-connectivity.ps1 -DomainName "server-vs81t.intranet.local:9443" -Detailed

# Tests both IIS and nginx configurations:
# - Correct server name (server-vs81t)
# - Port 9443 connectivity
# - SSL certificate validation
# - Frontend file deployment
# - Firewall rules
```

### Development Server vs IIS Comparison

| Aspect | npm start (Dev Server) | IIS Production |
|--------|------------------------|----------------|
| **URL** | `https://localhost:3000` | `https://server-vs81t.intranet.local:9443/excellence` |
| **SSL** | Self-signed development cert | Production server-vs81t certificates |
| **Hot Reload** | ✅ Automatic code reload | ❌ Must rebuild and deploy |
| **API Proxy** | Built into webpack dev server | IIS URL Rewrite + ARR to port 5000 |
| **Purpose** | Development and testing | Production deployment |
| **Debugging** | Easy to debug frontend issues | Use test scripts for debugging |
| **Management** | Command line (npm) | IIS Manager or PowerShell |

### Troubleshooting Common Issues

**HTTP Error 500.19 - Configuration Section Locked:**
This is the most common issue when migrating from nginx to IIS or updating web.config.

**Error symptoms:**
- `Config Error: This configuration section cannot be used at this path`
- Error Code `0x80070021` 
- References to `<handlers>` section in web.config

**Quick Fix:**
```powershell
# Run as Administrator
.\deployment\scripts\fix-iis-config.ps1
```

This script automatically:
- Backs up your current web.config
- Deploys the latest compatible configuration
- Removes problematic `<handlers>` section
- Tests the configuration

**When IIS doesn't work but npm start does:**
1. Run `.\deployment\scripts\test-iis-simple.ps1` to validate setup
2. Check that IIS has URL Rewrite and ARR modules installed
3. Verify SSL certificate is configured in IIS Manager 
4. Verify frontend files exist in `C:\inetpub\wwwroot\ExcelAddin\dist\`
5. Test backend separately with `.\deployment\scripts\test-backend-standalone.ps1`
6. Check web.config syntax and IIS logs in Event Viewer

**Common IIS Setup Issues:**
- **Missing modules**: Install URL Rewrite and Application Request Routing
- **SSL configuration**: Use IIS Manager to bind server-vs81t certificate
- **Permissions**: Ensure IIS_IUSRS has read access to physical directory
- **ARR not enabled**: Run `Set-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' -Filter 'system.webServer/proxy' -Name 'enabled' -Value 'true'`

## Service Management

### Essential PowerShell Scripts (in deployment/scripts/)

| Script | Purpose | Usage |
|--------|---------|-------|
| `setup-backend-service.ps1` | Install/configure backend Windows service | Run once during deployment |
| `setup-iis.ps1` | **Install/configure IIS with Excel Add-in site** | **For new IIS installations** |
| `deploy-to-existing-iis.ps1` | **Deploy Excel Add-in to existing IIS server** | **For existing IIS installations (recommended)** |
| `migrate-nginx-to-iis.ps1` | **Migrate from nginx to IIS** | **Use to switch from nginx to IIS** |
| `test-iis-simple.ps1` | **Test IIS configuration and connectivity** | **Use to validate IIS setup** |
| `diagnose-backend-service.ps1` | Troubleshoot backend service issues | Use when service fails to start |
| `diagnose-connectivity.ps1` | **Comprehensive connectivity diagnostics** | **Use when IIS is running but website not accessible** |
| `test-backend-standalone.ps1` | **Run Flask backend outside NSSM for debugging** | **Use when backend issues occur** |
| `add-firewall-rule.ps1` | **Automatically add Windows Firewall rule for port 9443** | **Run as Administrator** |
| `handle-encrypted-key.ps1` | Convert encrypted SSL keys | Use if SSL keys require passwords |
| `extract-pfx.ps1` | Extract certificates from PFX files | Use with PFX certificate files |

### Service Management Commands

```powershell
# Check service status
Get-Service ExcelAddinBackend, W3SVC
Get-Website -Name "ExcelAddin"

# Start services (in order)
Start-Service ExcelAddinBackend
Start-Website -Name "ExcelAddin"  # IIS service W3SVC starts automatically

# Stop services  
Stop-Website -Name "ExcelAddin"
Stop-Service ExcelAddinBackend

# Restart services
Restart-Service ExcelAddinBackend
Stop-Website -Name "ExcelAddin"; Start-Website -Name "ExcelAddin"

# View backend service logs (NSSM creates these)
Get-Content C:\inetpub\wwwroot\ExcelAddin\backend\logs\service.log -Tail 50

# View IIS logs
Get-Content C:\inetpub\logs\LogFiles\W3SVC*\*.log | Select-Object -Last 50

# IIS Management
inetmgr  # Opens IIS Manager GUI (run as Administrator)
```

### IIS Management Commands

```powershell
# Website management
Get-Website                           # List all websites
Get-Website -Name "ExcelAddin"       # Get specific website info
Start-Website -Name "ExcelAddin"     # Start website
Stop-Website -Name "ExcelAddin"      # Stop website
Restart-Website -Name "ExcelAddin"   # Restart website

# Application pool management
Get-IISAppPool -Name "ExcelAddinAppPool"     # Get app pool info
Start-WebAppPool -Name "ExcelAddinAppPool"   # Start app pool  
Stop-WebAppPool -Name "ExcelAddinAppPool"    # Stop app pool
Restart-WebAppPool -Name "ExcelAddinAppPool" # Restart app pool

# SSL certificate management
Get-WebBinding -Name "ExcelAddin"            # View SSL binding
Get-ChildItem Cert:\LocalMachine\My\        # List installed certificates
```
```

### Service Configuration with NSSM

Backend service configuration:
```powershell
# View current configuration
nssm dump ExcelAddinBackend

# Common configuration fixes
nssm set ExcelAddinBackend Application "C:\Python39\python.exe"
nssm set ExcelAddinBackend AppDirectory "C:\inetpub\wwwroot\ExcelAddin\backend"
nssm set ExcelAddinBackend AppParameters "service_wrapper.py"
```

nginx service configuration:
```powershell
# View current configuration
nssm dump nginx

# Ensure correct settings
nssm set nginx Application "C:\nginx\nginx.exe"
nssm set nginx AppDirectory "C:\nginx"
nssm set nginx AppParameters "-g \"daemon off;\""
```

## Subpath Deployment Configuration

To deploy at `https://server-vs81t.intranet.local:9443/excellence` instead of root:

### 1. nginx Configuration
- Set location blocks for `/excellence/` and `/excellence/api/`
- Configure proper proxying and static file serving

### 2. Frontend Configuration  
- Update `webpack.prod.config.js`: set `publicPath: '/excellence/'`
- Rebuild: `npm run build:staging`

### 3. Backend Configuration
- Update CORS_ORIGINS in `.env.production`
- Restart backend service

### 4. Manifest Configuration
- Update all URLs in `manifest-staging.xml` to include `/excellence` path
- Redistribute manifest to users

This deployment guide covers both local development and Windows Server production scenarios. For troubleshooting deployment issues, refer to the [Troubleshooting Guide](TROUBLESHOOTING_GUIDE.md).

## Troubleshooting Common Issues

### PowerShell curl Command Issues

**Problem**: `curl -k https://localhost:9443/excellence/api/health` fails with "A parameter cannot be found that matches parameter name 'k'."

**Cause**: In Windows PowerShell, `curl` is an alias for `Invoke-WebRequest`, which doesn't support the `-k` parameter.

**Solutions**:

```powershell
# Method 1: Use PowerShell 6.0+ syntax (if you have PowerShell Core installed)
Invoke-WebRequest -Uri "https://localhost:9443/excellence/api/health" -SkipCertificateCheck -UseBasicParsing

# Method 2: For Windows PowerShell 5.1 - ignore SSL certificates programmatically
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
Invoke-WebRequest -Uri "https://localhost:9443/excellence/api/health" -UseBasicParsing
# Reset for security after testing
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null

# Method 3: Use the actual curl.exe if installed (Git for Windows includes it)
curl.exe -k https://localhost:9443/excellence/api/health

# Method 4: Remove the curl alias to use real curl (advanced users only)
Remove-Item alias:curl -Force
curl -k https://localhost:9443/excellence/api/health
```

**Check PowerShell Version**: Run `$PSVersionTable.PSVersion` to determine which method to use.

### nginx Running but Website Not Accessible

**Problem**: nginx service is running but `Invoke-WebRequest` fails with "Unable to connect to the remote server"

**Cause**: Multiple possible issues with nginx configuration, SSL setup, or file deployment.

**Diagnosis Steps**:

```powershell
# 1. Verify nginx is actually listening on port 9443
netstat -an | findstr :9443
# Should show: TCP    0.0.0.0:9443    0.0.0.0:0    LISTENING

# 2. Check nginx error logs for specific issues
Get-Content "C:\nginx\logs\error.log" -Tail 20

# 3. Test if nginx is responding at all (without SSL)
# First, check if nginx has HTTP redirect configured
netstat -an | findstr :80
curl.exe -I http://localhost/excellence/ 2>&1

# 4. Verify SSL certificate files exist and are readable
Test-Path "C:\Cert\server.crt"
Test-Path "C:\Cert\server.key"
# Check certificate is not expired
if (Get-Command openssl -ErrorAction SilentlyContinue) {
    openssl x509 -in "C:\Cert\server.crt" -noout -dates
}

# 5. Verify frontend files are deployed correctly  
Test-Path "C:\inetpub\wwwroot\ExcelAddin\dist\taskpane.html"
Get-ChildItem "C:\inetpub\wwwroot\ExcelAddin\dist\" | Select-Object Name, Length

# 6. Test nginx configuration syntax
C:\nginx\nginx.exe -t

# 7. Check Windows Firewall is not blocking port 9443
New-NetFirewallRule -DisplayName "nginx HTTPS" -Direction Inbound -Protocol TCP -LocalPort 9443 -Action Allow -ErrorAction SilentlyContinue
```

**Common Fixes**:

```powershell
# Fix 1: Restart nginx service properly
Stop-Service nginx -Force
Start-Sleep -Seconds 5
Start-Service nginx

# Fix 2: Check nginx is using correct configuration file
C:\nginx\nginx.exe -t -c C:\nginx\conf\nginx.conf

# Fix 3: Verify nginx service is pointing to correct executable
nssm get nginx Application
nssm get nginx AppDirectory
# Should be: Application=C:\nginx\nginx.exe, AppDirectory=C:\nginx

# Fix 4: If SSL issues, temporarily test without HTTPS
# Edit nginx configuration to temporarily add HTTP listener on port 8080:
# server {
#     listen 8080;
#     server_name _;
#     location /excellence/ {
#         alias C:/inetpub/wwwroot/ExcelAddin/dist/;
#         index taskpane.html;
#     }
# }
# Then test: Invoke-WebRequest -Uri "http://localhost:8080/excellence/taskpane.html"
```

**Comprehensive Diagnostic Tool**: Use the automated connectivity diagnostic script to identify issues:
```powershell
cd C:\inetpub\wwwroot\ExcelAddin
.\deployment\scripts\diagnose-connectivity.ps1 -DomainName "localhost:9443" -Detailed
```

This script will automatically check:
- nginx service status and port listening
- SSL certificate files and validity  
- Frontend file deployment
- Connectivity tests with proper PowerShell version handling
- nginx configuration syntax
- Windows Firewall settings
- Recent error logs

### Frontend Not Loading

**Problem**: Excel add-in loads but shows blank taskpane or errors.

**Troubleshooting Steps**:

1. **Verify frontend files are deployed**:
```powershell
# Check that these files exist in the dist directory:
Test-Path "C:\inetpub\wwwroot\ExcelAddin\dist\taskpane.html"
Test-Path "C:\inetpub\wwwroot\ExcelAddin\dist\commands.html"  
Get-ChildItem "C:\inetpub\wwwroot\ExcelAddin\dist\" -Filter "*.js"
```

2. **Test nginx static file serving**:
```powershell
# For PowerShell 6.0+
if ($PSVersionTable.PSVersion.Major -ge 6) {
    Invoke-WebRequest -Uri "https://localhost:9443/excellence/taskpane.html" -SkipCertificateCheck
    Invoke-WebRequest -Uri "https://localhost:9443/excellence/manifest.xml" -SkipCertificateCheck
} else {
    # For Windows PowerShell 5.1
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri "https://localhost:9443/excellence/taskpane.html" -UseBasicParsing
    Invoke-WebRequest -Uri "https://localhost:9443/excellence/manifest.xml" -UseBasicParsing
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
}
```

3. **Check browser developer tools** (F12 in Excel):
   - Look for 404 errors on JavaScript/CSS files
   - Check console for JavaScript errors
   - Verify API calls are reaching backend

4. **Verify manifest file URLs** match your server configuration:
```xml
<!-- In manifest.xml, ensure URLs match your deployment -->
<SourceLocation DefaultValue="https://server-vs81t.intranet.local:9443/excellence/taskpane.html"/>
```