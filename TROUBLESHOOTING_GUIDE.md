# PrimeExcelence Excel Add-in - Troubleshooting Guide

## Table of Contents
1. [Troubleshooting Tools and Scripts](#troubleshooting-tools-and-scripts)
2. [Common Issues and Solutions](#common-issues-and-solutions)  
3. [Service-Specific Troubleshooting](#service-specific-troubleshooting)
4. [IIS Backend Issues](#iis-backend-issues)
5. [Backend Service Issues](#backend-service-issues)
6. [Excel Add-in Issues](#excel-add-in-issues)
7. [SSL Certificate Issues](#ssl-certificate-issues)
8. [Performance Issues](#performance-issues)
9. [Diagnostic Commands](#diagnostic-commands)

## Troubleshooting Tools and Scripts

### Core Diagnostic Scripts (deployment/scripts/)

| Script | When to Use | Purpose |
|--------|-------------|---------|
| `diagnose-backend-service.ps1` | Backend service won't start or crashes | Comprehensive backend service diagnostics and automatic fixes |
| `setup-backend-iis.ps1` | Setting up backend in IIS | Configures IIS FastCGI for Python backend hosting |  
| `handle-encrypted-key.ps1` | SSL key passwords required | Converts encrypted SSL keys to unencrypted format |
| `extract-pfx.ps1` | Using PFX certificate files | Extracts certificate and key from PFX files |

### Usage Examples

```powershell
# Diagnose backend service issues  
.\deployment\scripts\diagnose-backend-service.ps1 -FixCommonIssues

# Setup backend in IIS
.\deployment\scripts\setup-backend-iis.ps1

# Fix SSL key password prompts
.\deployment\scripts\handle-encrypted-key.ps1 -KeyPath "C:\Cert\server.key"

# Extract certificates from PFX
.\deployment\scripts\extract-pfx.ps1 -PfxPath "C:\Cert\server.pfx"
```

## Common Issues and Solutions

### 1. Certificate Errors in Excel Add-in

**Symptoms**: Excel shows "The content is blocked because it isn't signed by a valid security certificate"

**Quick Fix**:
```bash
npm run cert:install
```

**Detailed Solution**: See [CERTIFICATE_GUIDE.md](CERTIFICATE_GUIDE.md) for comprehensive certificate management instructions.

### 2. "Service won't start" (No error messages)

**Symptoms**: Service fails to start with no logs or error messages

**Diagnosis**:
```powershell
# Run comprehensive diagnostics
.\deployment\scripts\diagnose-backend-service.ps1 -FixCommonIssues

# Check IIS Application Configuration
Get-WebApplication -Name "backend" -Site "Default Web Site"
```

**Common Causes & Solutions**:
- **Wrong Python path**: IIS FastCGI using incorrect Python executable
  ```powershell
  # Fix: Update web.config with correct Python path
  # Edit backend/web.config and update the FastCGI application path
  ```
- **Missing dependencies**: Python packages not installed
  ```powershell
  cd C:\inetpub\wwwroot\ExcelAddin\backend
  poetry install
  ```
- **WSGI application errors**: Check IIS logs and application event viewer
  ```powershell
  # Check IIS logs
  Get-ChildItem "C:\inetpub\logs\LogFiles\W3SVC1" | Sort-Object LastWriteTime -Descending | Select-Object -First 5
  ```

### 2. "IIS Application Pool crashes immediately"

**Symptoms**: IIS application pool for backend stops unexpectedly

**Solution**: Check IIS logs and Python WSGI configuration
```powershell
# Check IIS application pool status
Get-WebAppPoolState -Name "DefaultAppPool"

# Check IIS logs for errors
Get-ChildItem "C:\inetpub\logs\LogFiles\W3SVC1" | Sort-Object LastWriteTime -Descending | Select-Object -First 1 | Get-Content -Tail 50
```

### 3. "Python WSGI application errors"

**Symptoms**: HTTP 500 errors when accessing backend API

**Solution**: Check Python environment and dependencies
```powershell
# Check Python path in web.config
Get-Content "C:\inetpub\wwwroot\ExcelAddin\backend\web.config" | Select-String "python.exe"

# Test WSGI app directly
cd "C:\inetpub\wwwroot\ExcelAddin\backend"
python wsgi_app.py
```

### 4. "npm build fails - custom-functions-metadata not found"

**Symptoms**: Build fails after `npm install --production`

**Root Cause**: Build tools are devDependencies, not available in production install

**Solution**: Use proper build-then-deploy approach
```bash
# On build machine (with full dependencies)
npm install  
npm run build:staging

# Deploy only the dist/ folder to production server
# No npm install needed on production server
```

### 5. "Bundle size warnings" 

**Symptoms**: Webpack warnings about large bundle sizes (>244 KiB)

**Solution**: Already optimized - 65% reduction achieved
- Bundle reduced from 983 KiB to 338 KiB
- Uses lazy loading for page components
- Optimized chart libraries and dependencies
- Remaining size is acceptable for feature-rich Excel add-in

## Service-Specific Troubleshooting

### Backend Service Diagnostics

#### Manual Service Testing
```powershell
# Test Python environment manually
cd C:\inetpub\wwwroot\ExcelAddin\backend
poetry shell
python service_wrapper.py

# Test Flask app directly
poetry shell
python run.py

# Test API endpoint
curl http://localhost:5000/api/health
```

#### Service Debugging Tools
```batch
# Use debug batch file for interactive testing
cd C:\inetpub\wwwroot\ExcelAddin\backend
debug-service.bat

# View detailed service logs
Get-Content logs\service.log -Tail 50 -Wait
```

#### Common Backend Service Fixes

**Issue**: Service installed but won't start
```powershell
# 1. Verify Python path in IIS FastCGI configuration
where python
# Update backend/web.config with correct Python path

# 2. Check IIS application physical path
Get-WebApplication -Name "backend" -Site "Default Web Site"

# 3. Verify WSGI entry point exists
Test-Path "C:\inetpub\wwwroot\ExcelAddin\backend\service_wrapper.py"

# 4. Test manual startup
cd C:\inetpub\wwwroot\ExcelAddin\backend
poetry shell  
python service_wrapper.py
```

**Issue**: Service starts but API unreachable
```powershell
# Check if port 5000 is in use
netstat -an | findstr :5000

# Test backend directly
curl http://localhost:5000/api/health

# Check firewall settings
New-NetFirewallRule -DisplayName "Flask API" -Direction Inbound -Protocol TCP -LocalPort 5000 -Action Allow
```

### nginx Service Diagnostics  

#### Configuration Validation
```powershell
# Test nginx configuration syntax
C:\nginx\nginx.exe -t

# Test with specific config file
C:\nginx\nginx.exe -t -c C:\nginx\conf\nginx.conf

# Use validation script
.\deployment\scripts\validate-nginx-config.ps1
```

#### Common nginx Service Fixes

**Issue**: HTTP/2 deprecation warnings
```nginx
# Old (deprecated)
listen 8443 ssl http2;

# New (correct)  
listen 8443 ssl;
http2 on;
```

**Issue**: SSL stapling warnings for company certificates
```nginx
# Disable SSL stapling for internal certificates
# ssl_stapling off;
# ssl_stapling_verify off;
```

## IIS Backend Issues

### Python FastCGI Configuration Problems

#### 1. FastCGI Module Issues
```powershell
# Verify FastCGI module is installed
Get-WindowsFeature -Name IIS-CGI

# Install if needed
Enable-WindowsOptionalFeature -Online -FeatureName IIS-CGI
```

#### 2. Python Path Configuration
```powershell  
# Check current Python installation
where python

# Update web.config with correct path
# Edit backend/web.config FastCGI application fullPath
```

#### 3. WSGI Handler Configuration  
```xml
<!-- Verify web.config has correct WSGI handler -->
<add name="WSGI_HANDLER" value="wsgi_app.application" />
```

### IIS Log Analysis
```powershell
# Check IIS logs for errors
Get-ChildItem "C:\inetpub\logs\LogFiles\W3SVC1" | Sort-Object LastWriteTime -Descending | Select-Object -First 1 | Get-Content -Tail 50

# Check IIS Failed Request Tracing (if enabled)
Get-ChildItem "C:\inetpub\logs\FailedReqLogFiles" | Sort-Object LastWriteTime -Descending

# Monitor IIS logs in real-time
Get-Content "C:\inetpub\logs\LogFiles\W3SVC1\*.log" -Wait -Tail 10
```

## Backend Service Issues

### Python Environment Issues

#### 0. Poetry Setup Issues
```powershell
# Install Poetry if not installed
(Invoke-WebRequest -Uri https://install.python-poetry.org -UseBasicParsing).Content | python -

# Verify Poetry installation
poetry --version

# If Poetry is not in PATH, add it manually or restart terminal
```

#### 1. Python Path Problems
```powershell
# Find Python executable
where python
Get-Command python

# Update web.config FastCGI configuration with correct path
$pythonPath = (Get-Command python).Source
# Edit backend/web.config and update fullPath attribute
```

#### 2. Dependency Issues
```powershell
# Verify all packages installed
cd C:\inetpub\wwwroot\ExcelAddin\backend
poetry show

# Reinstall dependencies
poetry install --no-dev

# Check for import errors in Poetry shell
poetry shell
python -c "import flask; print('Flask OK')"
```

#### 3. IIS Integration Issues
```powershell
# Check IIS application pool status
Get-WebAppPoolState -Name "DefaultAppPool"

# Restart application pool if needed
Restart-WebAppPool -Name "DefaultAppPool"

# Check IIS application configuration
Get-WebApplication -Name "backend" -Site "Default Web Site"
```

### Flask Application Issues

#### 1. Environment Configuration
```bash
# Check .env.production file
FLASK_ENV=production
FLASK_DEBUG=False
API_BASE_URL=https://server01.intranet.local:8443/excellence/api
CORS_ORIGINS=https://server01.intranet.local:8443
```

#### 2. CORS Configuration Problems
```python
# Verify CORS settings in app.py
CORS(app, origins=[
    'https://server01.intranet.local:8443',
    'https://excel.office.com',
    'https://excel-online.microsoft.com'
])
```

## Excel Add-in Issues

### 1. Add-in Won't Load

**Symptoms**: PrimeExcelence button doesn't appear in Excel

**Diagnosis**:
```powershell
# Test manifest URL accessibility
curl -k https://server01.intranet.local:8443/excellence/taskpane.html

# Check if domain is in Excel's trusted domains
# Excel → File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs
```

**Solutions**:
- Verify manifest XML syntax
- Ensure all URLs use HTTPS  
- Check that SSL certificate is trusted
- Confirm add-in is properly uploaded/installed

### 2. SSL Certificate Trust Issues

**Symptoms**: "Add-in error" or "Cannot load add-in"

**Solutions**:
```powershell
# Install company root CA certificate
certlm.msc  # Open certificate manager
# Import cacert.pem to Trusted Root Certification Authorities

# Or via PowerShell
Import-Certificate -FilePath "C:\Cert\cacert.pem" -CertStoreLocation Cert:\LocalMachine\Root
```

### 3. CORS Errors in Excel

**Symptoms**: API calls fail with CORS errors in Excel's developer console

**Solutions**:
```bash
# Update backend CORS configuration
CORS_ORIGINS=https://server01.intranet.local:8443,https://excel.office.com

# Restart backend service  
Restart-Service ExcelAddinBackend
```

## SSL Certificate Issues

### Certificate Format Problems

#### 1. PFX Certificate Usage
```powershell
# Extract certificate and key from PFX
.\deployment\scripts\extract-pfx.ps1 -PfxPath "C:\Cert\server.pfx" -OutputPath "C:\Cert"

# Results in:
# C:\Cert\server.crt
# C:\Cert\server.key  
```

#### 2. Encrypted Private Key Issues
```powershell
# Check if key is encrypted
openssl rsa -in C:\Cert\server.key -check
# If prompts for password, key is encrypted

# Convert to unencrypted
.\deployment\scripts\handle-encrypted-key.ps1 -KeyPath "C:\Cert\server.key"
```

#### 3. Certificate Chain Issues
```powershell
# Verify certificate chain
openssl verify -CAfile C:\Cert\cacert.pem C:\Cert\server.crt

# Test SSL connection
openssl s_client -connect server01.intranet.local:8443 -servername server01.intranet.local
```

### Certificate Installation Issues

#### 1. Windows Certificate Store
```powershell
# Import root CA to system store
Import-Certificate -FilePath "C:\Cert\cacert.pem" -CertStoreLocation Cert:\LocalMachine\Root

# Verify installation
Get-ChildItem Cert:\LocalMachine\Root | Where-Object {$_.Subject -like "*Your Company*"}
```

#### 2. nginx Certificate Configuration
```nginx
# Ensure correct paths and permissions
ssl_certificate C:/Cert/server.crt;
ssl_private_key C:/Cert/server.key;
ssl_trusted_certificate C:/Cert/cacert.pem;
```

## Performance Issues

### 1. Slow Add-in Loading

**Diagnosis**:
```powershell
# Test response times
Measure-Command { curl -k https://server01.intranet.local:8443/excellence/taskpane.html }

# Check bundle sizes
Get-ChildItem C:\inetpub\wwwroot\ExcelAddin\*.js | Select-Object Name, Length
```

**Optimizations Applied**:
- Lazy loading for page components (10-20 KB each)
- Bundle splitting for vendor libraries
- Custom lightweight charts (replaced 400+ KB recharts)
- Total bundle reduction: 65% (983 KB → 338 KB)

### 2. API Response Times

**Diagnosis**:
```powershell  
# Test API response times
Measure-Command { curl http://localhost:5000/api/health }
Measure-Command { curl -k https://server01.intranet.local:8443/excellence/api/health }
```

**Solutions**:
- Enable nginx caching for static assets
- Optimize database queries in Flask app
- Consider API response caching

## Diagnostic Commands

### Service Status Commands
```powershell
# Check IIS status
Get-Service W3SVC, WAS

# View IIS configuration
Get-WebApplication -Name "backend" -Site "Default Web Site"
Get-WebAppPoolState -Name "DefaultAppPool"

# View IIS logs  
Get-ChildItem "C:\inetpub\logs\LogFiles\W3SVC1" | Sort-Object LastWriteTime -Descending | Select-Object -First 1 | Get-Content -Tail 50
```

### Network Diagnostic Commands
```powershell
# Test IIS web server
Test-NetConnection -ComputerName localhost -Port 80   # HTTP
Test-NetConnection -ComputerName localhost -Port 443  # HTTPS

# Check IIS bindings
Get-WebBinding -Name "Default Web Site"

# Test SSL (if configured)  
openssl s_client -connect server01.intranet.local:443
```

### Application Diagnostic Commands
```powershell
# Test backend API through IIS
curl http://localhost/backend/api/health

# Test through IIS with SSL
curl -k https://localhost/excellence/api/health

# Test frontend application
curl -k https://localhost/excellence/

# Check IIS configuration
Get-WebConfiguration -Filter "system.webServer/fastCgi"
```

### Certificate Diagnostic Commands
```powershell
# View certificate details
openssl x509 -in C:\Cert\server.crt -text -noout

# Test private key
openssl rsa -in C:\Cert\server.key -check

# Verify certificate-key match
openssl x509 -in C:\Cert\server.crt -noout -modulus | openssl md5
openssl rsa -in C:\Cert\server.key -noout -modulus | openssl md5
# Results should match
```

### Log File Locations
```
Backend Logs:
  C:\inetpub\logs\LogFiles\W3SVC1\*.log           # IIS access logs
  C:\inetpub\logs\FailedReqLogFiles\*\*.xml       # IIS failed request logs

Frontend Logs:
  C:\inetpub\logs\LogFiles\W3SVC1\*.log           # IIS access logs
  Browser Developer Console                        # Client-side errors

nginx Logs:  
  C:\nginx\logs\access.log
  C:\nginx\logs\error.log

Windows Event Logs:
  Event Viewer → Windows Logs → Application
  Look for "ExcelAddinBackend" and "nginx" sources
```

This troubleshooting guide covers the most common issues and their solutions. For additional support, use the provided diagnostic scripts which can automatically detect and fix many common configuration problems.