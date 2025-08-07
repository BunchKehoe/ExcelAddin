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
        └── https://server01.intranet.local:8443/excellence
                    │
                    ▼
        ┌─────────────────────────────────────────────┐
        │              nginx (Port 8443)              │
        │                                             │
        │  • SSL Termination                         │
        │  • Static File Serving (/excellence/)      │
        │    - taskpane.html, commands.html          │
        │    - JavaScript bundles, CSS               │
        │    - Images, manifest.xml                  │
        │  • API Proxy (/excellence/api/ → :5000)    │
        │  • Windows Service (via NSSM)              │
        └─────────────────────────────────────────────┘
                    │
                    └── Backend Flask App (Port 5000)
                              │
                              └── Windows Service (via NSSM)
```

**How Frontend Hosting Works:**
- nginx serves all frontend files (HTML, JS, CSS, images) as static content
- Files are served from `C:\inetpub\wwwroot\ExcelAddin\dist\` directory
- URL `https://server01.intranet.local:8443/excellence/` loads `dist/taskpane.html`
- Excel add-in loads in taskpane, makes API calls to `/excellence/api/` endpoints
- nginx proxies API calls to backend Flask app running on port 5000

## Frontend Deployment Overview

The Excel add-in frontend is a **React-based single-page application** that gets compiled into static files and served by nginx. Here's how it works:

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
- **Hosting**: nginx serves files from `C:\inetpub\wwwroot\ExcelAddin\dist\`
- **Public Path**: All resources served under `/excellence/` subpath

### 3. Integration with Excel
- **Manifest File**: `manifest.xml` defines add-in metadata and entry points
- **Taskpane**: Main interface loads at `https://server:8443/excellence/taskpane.html`  
- **Commands**: Ribbon commands load at `https://server:8443/excellence/commands.html`
- **API Communication**: Frontend makes AJAX calls to backend API

### 4. nginx Configuration for Frontend
The nginx server is configured to:
- Serve static files from the application directory
- Handle URL routing for single-page app behavior
- Proxy API requests to the backend Flask application
- Provide SSL termination for secure HTTPS access

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

### Phase 2: Configure nginx

```powershell
# Copy nginx configuration
cd C:\inetpub\wwwroot\ExcelAddin\deployment

# Copy nginx.conf.windows.template to C:\nginx\conf\nginx.conf
copy nginx\nginx.conf.windows.template C:\nginx\conf\nginx.conf

# Copy Excel add-in configuration  
copy nginx\excel-addin.conf C:\nginx\conf\excel-addin.conf

# Edit C:\nginx\conf\excel-addin.conf:
# - Update server_name to your domain
# - Verify certificate paths in C:\Cert\
# - Ensure subpath is /excellence/ if needed
```

### Phase 3: Set Up Backend Service

```powershell
# Run as Administrator
cd C:\inetpub\wwwroot\ExcelAddin\deployment\scripts

# Install and configure backend service
.\setup-backend-service.ps1 -Force

# Verify service installation
Get-Service ExcelAddinBackend
```

### Phase 4: Set Up nginx Service

```powershell  
# Run as Administrator  
cd C:\inetpub\wwwroot\ExcelAddin\deployment\scripts

# Install and configure nginx service
.\setup-nginx-service.ps1 -Force

# Verify service installation
Get-Service nginx
```

### Phase 5: Start Services

```powershell
# Start backend service first
Start-Service ExcelAddinBackend

# Wait 10 seconds for backend to initialize
Start-Sleep -Seconds 10

# Start nginx service
Start-Service nginx

# Verify both services are running
Get-Service ExcelAddinBackend, nginx
```

### Phase 6: Validate Deployment

```powershell
# Test backend API directly
Invoke-WebRequest -Uri "http://localhost:5000/api/health" -UseBasicParsing

# Test through nginx proxy (skipping certificate check for self-signed certs)
Invoke-WebRequest -Uri "https://localhost:8443/excellence/api/health" -SkipCertificateCheck -UseBasicParsing

# Test main application
Invoke-WebRequest -Uri "https://localhost:8443/excellence/" -SkipCertificateCheck -UseBasicParsing

# Alternative: Use curl.exe directly (if installed) instead of PowerShell's curl alias
# curl.exe -k https://localhost:8443/excellence/api/health
# curl.exe -k https://localhost:8443/excellence/
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

### nginx Configuration (`deployment/nginx/excel-addin.conf`)

Key settings for subpath deployment:

```nginx
server {
    listen 8443 ssl;
    server_name server01.intranet.local;
    
    # SSL Configuration
    ssl_certificate C:/Cert/server.crt;
    ssl_private_key C:/Cert/server.key;
    
    # Serve app at /excellence/
    location /excellence/ {
        alias C:/inetpub/wwwroot/ExcelAddin/;
        try_files $uri $uri/ @fallback;
    }
    
    # Proxy API calls to backend  
    location /excellence/api/ {
        proxy_pass http://127.0.0.1:5000/api/;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
}
```

### Backend Environment (`.env.production`)

```bash
FLASK_ENV=production
FLASK_DEBUG=False
API_BASE_URL=https://server01.intranet.local:8443/excellence/api
CORS_ORIGINS=https://server01.intranet.local:8443
DATABASE_CONFIG=database.cfg
```

### Excel Manifest (`manifest-staging.xml`)

```xml
<SourceLocation DefaultValue="https://server01.intranet.local:8443/excellence/taskpane.html"/>
<SupportUrl DefaultValue="https://server01.intranet.local:8443/excellence/"/>
<AppDomains>
  <AppDomain>https://server01.intranet.local:8443</AppDomain>
</AppDomains>
```

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

# Test SSL configuration  
.\deployment\scripts\validate-nginx-config.ps1
```

## Service Management

### Essential PowerShell Scripts (in deployment/scripts/)

| Script | Purpose | Usage |
|--------|---------|-------|
| `setup-backend-service.ps1` | Install/configure backend Windows service | Run once during deployment |
| `setup-nginx-service.ps1` | Install/configure nginx Windows service | Run once during deployment |
| `diagnose-backend-service.ps1` | Troubleshoot backend service issues | Use when service fails to start |
| `validate-nginx-config.ps1` | Test nginx configuration | Use before starting nginx |
| `handle-encrypted-key.ps1` | Convert encrypted SSL keys | Use if SSL keys require passwords |
| `extract-pfx.ps1` | Extract certificates from PFX files | Use with PFX certificate files |

### Service Management Commands

```powershell
# Check service status
Get-Service ExcelAddinBackend, nginx

# Start services (in order)
Start-Service ExcelAddinBackend
Start-Service nginx

# Stop services  
Stop-Service nginx
Stop-Service ExcelAddinBackend

# Restart services
Restart-Service ExcelAddinBackend
Restart-Service nginx

# View service logs (NSSM creates these)
Get-Content C:\nginx\logs\service.log -Tail 50
Get-Content C:\inetpub\wwwroot\ExcelAddin\backend\logs\service.log -Tail 50
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

To deploy at `https://server01.intranet.local:8443/excellence` instead of root:

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

**Problem**: `curl -k https://localhost:8443/excellence/api/health` fails with "A parameter cannot be found that matches parameter name 'k'."

**Cause**: In Windows PowerShell, `curl` is an alias for `Invoke-WebRequest`, which doesn't support the `-k` parameter.

**Solutions**:

```powershell
# Option 1: Use Invoke-WebRequest with proper PowerShell syntax
Invoke-WebRequest -Uri "https://localhost:8443/excellence/api/health" -SkipCertificateCheck -UseBasicParsing

# Option 2: Use the actual curl.exe if installed (Git for Windows includes it)
curl.exe -k https://localhost:8443/excellence/api/health

# Option 3: Remove the curl alias to use real curl (advanced users)
Remove-Item alias:curl
curl -k https://localhost:8443/excellence/api/health
```

**Health Check Script**: Use the provided health check script for comprehensive testing:
```powershell
cd C:\inetpub\wwwroot\ExcelAddin
.\deployment\monitoring\health-check.ps1 -DomainName "localhost:8443" -Detailed
```

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
Invoke-WebRequest -Uri "https://localhost:8443/excellence/taskpane.html" -SkipCertificateCheck
Invoke-WebRequest -Uri "https://localhost:8443/excellence/manifest.xml" -SkipCertificateCheck  
```

3. **Check browser developer tools** (F12 in Excel):
   - Look for 404 errors on JavaScript/CSS files
   - Check console for JavaScript errors
   - Verify API calls are reaching backend

4. **Verify manifest file URLs** match your server configuration:
```xml
<!-- In manifest.xml, ensure URLs match your deployment -->
<SourceLocation DefaultValue="https://server01.intranet.local:8443/excellence/taskpane.html"/>
```