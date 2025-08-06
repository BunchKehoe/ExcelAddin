# Excel Add-in Windows Server Deployment Guide

This guide provides step-by-step instructions for deploying the Excel Add-in to a Windows Server environment with nginx as a reverse proxy and Windows services for the backend.

## Table of Contents

- [Prerequisites](#prerequisites)
- [Server Preparation](#server-preparation)
- [SSL Certificate Setup](#ssl-certificate-setup)
- [Deployment Process](#deployment-process)
- [Service Management](#service-management)
- [Monitoring and Maintenance](#monitoring-and-maintenance)
- [Troubleshooting](#troubleshooting)
- [User Distribution](#user-distribution)

## Prerequisites

### Software Requirements

1. **Windows Server 2016 or later**
2. **Python 3.8 or later**
   - Download from [python.org](https://www.python.org/downloads/)
   - Ensure Python is added to PATH during installation
3. **Node.js 16 or later** (for building the frontend)
   - Download from [nodejs.org](https://nodejs.org/)
4. **nginx for Windows**
   - Download from [nginx.org](http://nginx.org/en/download.html)
   - Extract to `C:\nginx`
5. **NSSM (Non-Sucking Service Manager)**
   - Download from [nssm.cc](https://nssm.cc/download)
   - Extract to a location in PATH or `C:\Program Files\NSSM`

### Network Requirements

- **Port 443 (HTTPS)** - Primary application access
- **Port 80 (HTTP)** - Redirect to HTTPS
- **Port 5000** - Backend API (internal, not exposed externally)
- **DNS record** pointing your domain to the server

### Permissions

- **Local Administrator rights** on the Windows Server
- **Ability to create Windows Services**
- **Access to DNS management** for your domain

## Server Preparation

### 1. Create Directory Structure

```powershell
# Create main application directories
New-Item -ItemType Directory -Path "C:\inetpub\wwwroot\ExcelAddin" -Force
New-Item -ItemType Directory -Path "C:\Logs\ExcelAddin" -Force
New-Item -ItemType Directory -Path "C:\Logs\nginx" -Force
New-Item -ItemType Directory -Path "C:\ssl\certs" -Force
New-Item -ItemType Directory -Path "C:\ssl\private" -Force
```

### 2. Install Python Dependencies

```powershell
# Verify Python installation
python --version
pip --version

# Install required Python packages globally
pip install flask flask-cors python-dotenv requests sqlalchemy pyodbc configparser
```

### 3. Install nginx

1. Download nginx for Windows
2. Extract to `C:\nginx`
3. Test nginx installation:
   ```cmd
   cd C:\nginx
   nginx -t
   ```

### 4. Install NSSM

1. Download NSSM
2. Extract to `C:\Program Files\NSSM` or add to PATH
3. Verify installation:
   ```cmd
   nssm version
   ```

## SSL Certificate Setup

Choose one of the following methods:

### Option 1: Commercial Certificate (Recommended for Production)

1. **Generate Certificate Signing Request (CSR)**:
   ```powershell
   # Create CSR using certreq
   # Edit the subject line with your organization details
   certreq -new -f @"
   [NewRequest]
   Subject="CN=your-staging-domain.com,O=Your Organization,L=City,S=State,C=US"
   KeyLength=2048
   Exportable=TRUE
   MachineKeySet=TRUE
   RequestType=PKCS10
   "@
   ```

2. **Submit CSR to Certificate Authority**
3. **Install the certificate** in Windows Certificate Store
4. **Export certificate files** for nginx

### Option 2: Let's Encrypt (Free, Automated)

1. **Install win-acme**:
   - Download from [win-acme.com](https://www.win-acme.com/)
   - Run the executable and follow prompts
   - Choose nginx integration

### Option 3: Self-Signed (Development/Testing Only)

```powershell
# Generate self-signed certificate
$cert = New-SelfSignedCertificate -DnsName "your-staging-domain.com" -CertStoreLocation "cert:\LocalMachine\My" -KeyLength 2048

# Export certificate files
$certPath = "C:\ssl\certs\excel-addin.crt"
$keyPath = "C:\ssl\private\excel-addin.key"

# Export certificate
[System.IO.File]::WriteAllBytes($certPath, $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))
```

## Deployment Process

### 1. Build and Prepare Application

```powershell
# Clone repository (if not already done)
git clone <your-repo-url> C:\temp\ExcelAddin
cd C:\temp\ExcelAddin

# Install frontend dependencies and build
npm install
npm run build

# Copy built files to deployment directory
Copy-Item -Path "dist\*" -Destination "C:\inetpub\wwwroot\ExcelAddin\dist\" -Recurse -Force
Copy-Item -Path "backend\*" -Destination "C:\inetpub\wwwroot\ExcelAddin\backend\" -Recurse -Force
Copy-Item -Path "manifest-staging.xml" -Destination "C:\inetpub\wwwroot\ExcelAddin\" -Force
```

### 2. Configure Backend Environment

```powershell
# Navigate to backend directory
cd C:\inetpub\wwwroot\ExcelAddin\backend

# Create production environment file
Copy-Item -Path ".env.production" -Destination ".env" -Force

# Edit .env file with your specific settings
# Update the following variables:
# - CORS_ORIGINS=https://your-staging-domain.com
# - NIFI_ENDPOINT (if applicable)
# - Domain-specific configurations
```

### 3. Configure nginx

```powershell
# Copy nginx configuration
Copy-Item -Path "C:\temp\ExcelAddin\deployment\nginx\excel-addin.conf" -Destination "C:\nginx\conf\conf.d\" -Force

# Update configuration with your domain
$configPath = "C:\nginx\conf\conf.d\excel-addin.conf"
(Get-Content $configPath) -replace 'your-staging-domain.com', 'your-actual-domain.com' | Set-Content $configPath

# Update SSL certificate paths
(Get-Content $configPath) -replace 'C:/ssl/certs/excel-addin.crt', 'C:\ssl\certs\excel-addin.crt' | Set-Content $configPath
(Get-Content $configPath) -replace 'C:/ssl/private/excel-addin.key', 'C:\ssl\private\excel-addin.key' | Set-Content $configPath

# Test nginx configuration
cd C:\nginx
nginx -t
```

### 4. Install Backend Windows Service

```powershell
# Install service using NSSM
nssm install ExcelAddinBackend python "C:\inetpub\wwwroot\ExcelAddin\backend\service_wrapper.py"

# Configure service
nssm set ExcelAddinBackend AppDirectory "C:\inetpub\wwwroot\ExcelAddin\backend"
nssm set ExcelAddinBackend DisplayName "Excel Add-in Backend Service"
nssm set ExcelAddinBackend Description "Python Flask backend service for Excel Add-in"
nssm set ExcelAddinBackend Start SERVICE_AUTO_START
nssm set ExcelAddinBackend AppStdout "C:\Logs\ExcelAddin\service-stdout.log"
nssm set ExcelAddinBackend AppStderr "C:\Logs\ExcelAddin\service-stderr.log"

# Set environment variables
nssm set ExcelAddinBackend AppEnvironmentExtra FLASK_ENV=production DEBUG=false HOST=127.0.0.1 PORT=5000 PYTHONPATH=C:\inetpub\wwwroot\ExcelAddin\backend

# Start the service
Start-Service ExcelAddinBackend
```

### 5. Start nginx

```powershell
# Start nginx
cd C:\nginx
Start-Process nginx -WindowStyle Hidden

# Verify nginx is running
Get-Process nginx
```

### 6. Update Manifest File

```powershell
# Edit manifest-staging.xml
$manifestPath = "C:\inetpub\wwwroot\ExcelAddin\manifest-staging.xml"

# Update URLs with your domain
(Get-Content $manifestPath) -replace 'your-staging-domain.com', 'your-actual-domain.com' | Set-Content $manifestPath

# Generate a new GUID for the staging environment
$newGuid = [System.Guid]::NewGuid().ToString()
(Get-Content $manifestPath) -replace 'a1b2c3d4-e5f6-7890-abcd-ef1234567890', $newGuid | Set-Content $manifestPath
```

## Service Management

### Backend Service Commands

```powershell
# Start service
Start-Service ExcelAddinBackend

# Stop service
Stop-Service ExcelAddinBackend

# Restart service
Restart-Service ExcelAddinBackend

# Check service status
Get-Service ExcelAddinBackend

# View service logs
Get-Content "C:\Logs\ExcelAddin\service-stdout.log" -Tail 50 -Wait
```

### nginx Commands

```powershell
# Start nginx
cd C:\nginx
.\nginx.exe

# Stop nginx
.\nginx.exe -s quit

# Reload configuration
.\nginx.exe -s reload

# Test configuration
.\nginx.exe -t

# Check nginx processes
Get-Process nginx
```

## Monitoring and Maintenance

### 1. Set Up Health Check Monitoring

```powershell
# Copy health check script
Copy-Item -Path "C:\temp\ExcelAddin\deployment\monitoring\health-check.ps1" -Destination "C:\Scripts\" -Force

# Test health check
PowerShell -ExecutionPolicy Bypass -File "C:\Scripts\health-check.ps1" -DomainName "your-actual-domain.com"

# Create scheduled task for health monitoring
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-ExecutionPolicy Bypass -File C:\Scripts\health-check.ps1 -DomainName your-actual-domain.com -SendAlerts"
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes 15) -RepetitionDuration (New-TimeSpan -Days 365)
$principal = New-ScheduledTaskPrincipal -UserID "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
Register-ScheduledTask -TaskName "Excel-Addin-Health-Check" -Action $action -Trigger $trigger -Principal $principal
```

### 2. Log Rotation

Set up log rotation to prevent log files from growing too large:

```powershell
# Create log cleanup script
@"
Get-ChildItem "C:\Logs\ExcelAddin\*.log" | Where-Object {$_.LastWriteTime -lt (Get-Date).AddDays(-30)} | Remove-Item -Force
Get-ChildItem "C:\Logs\nginx\*.log" | Where-Object {$_.LastWriteTime -lt (Get-Date).AddDays(-30)} | Remove-Item -Force
"@ | Set-Content "C:\Scripts\cleanup-logs.ps1"

# Schedule log cleanup
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-ExecutionPolicy Bypass -File C:\Scripts\cleanup-logs.ps1"
$trigger = New-ScheduledTaskTrigger -Daily -At "2:00AM"
$principal = New-ScheduledTaskPrincipal -UserID "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
Register-ScheduledTask -TaskName "Excel-Addin-Log-Cleanup" -Action $action -Trigger $trigger -Principal $principal
```

## Troubleshooting

### Common Issues

1. **Service won't start**
   - Check Python installation and PATH
   - Verify all dependencies are installed
   - Check service logs in `C:\Logs\ExcelAddin\`

2. **nginx SSL errors**
   - Verify certificate file paths
   - Check certificate file permissions
   - Test certificate with `openssl` if available

3. **Application not accessible**
   - Check Windows Firewall settings
   - Verify DNS configuration
   - Test with local IP address first

4. **CORS errors in Excel**
   - Verify CORS_ORIGINS in backend .env file
   - Check that domain matches exactly
   - Ensure HTTPS is working properly

### Diagnostic Commands

```powershell
# Check all services
Get-Service | Where-Object {$_.Name -like "*Excel*"}

# Check ports
netstat -an | findstr ":443"
netstat -an | findstr ":5000"

# Test backend directly
Invoke-WebRequest -Uri "http://127.0.0.1:5000/api/health" -UseBasicParsing

# Test frontend
Invoke-WebRequest -Uri "https://your-domain.com/health" -UseBasicParsing

# Check certificates
Get-ChildItem -Path "Cert:\LocalMachine\My" | Where-Object {$_.Subject -like "*your-domain*"}
```

## User Distribution

### For Development Team (Staging Environment)

1. **Distribute the manifest file**:
   - Share `manifest-staging.xml` with developers
   - Include installation instructions

2. **Excel Installation Instructions**:
   ```
   1. Open Excel
   2. Go to Insert > Add-ins > Upload My Add-in
   3. Select the manifest-staging.xml file
   4. The add-in will appear in the Home tab
   ```

3. **Access Control**:
   - Limit staging environment access to developer IP addresses
   - Use VPN if necessary
   - Consider basic authentication for additional security

### For Production Distribution

1. **SharePoint App Catalog** (Recommended for organizations)
2. **Microsoft AppSource** (For public distribution)
3. **Network Share** (For internal distribution)

## Security Considerations

### Network Security
- Configure Windows Firewall to only allow necessary ports
- Use HTTPS everywhere
- Implement IP restrictions if possible

### Application Security
- Regular security updates for all components
- Monitor logs for suspicious activity
- Implement proper authentication if handling sensitive data

### Certificate Management
- Set up certificate renewal monitoring
- Use strong encryption (TLS 1.2+)
- Regular certificate expiration checks

---

This deployment guide provides a comprehensive approach to running the Excel Add-in on Windows Server. Adjust the domain names, paths, and configurations according to your specific environment.