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
        │  • API Proxy (/excellence/api/ → :5000)    │
        │  • Windows Service (via NSSM)              │
        └─────────────────────────────────────────────┘
                    │
                    └── Backend Flask App (Port 5000)
                              │
                              └── Windows Service (via NSSM)
```

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
curl http://localhost:5000/api/health

# Test through nginx proxy
curl -k https://localhost:8443/excellence/api/health

# Test main application
curl -k https://localhost:8443/excellence/
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