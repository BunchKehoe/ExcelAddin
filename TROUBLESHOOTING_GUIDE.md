# PrimeExcelence Excel Add-in - Troubleshooting Guide

## Table of Contents
1. [Troubleshooting Tools and Scripts](#troubleshooting-tools-and-scripts)
2. [Common Issues and Solutions](#common-issues-and-solutions)  
3. [Service-Specific Troubleshooting](#service-specific-troubleshooting)
4. [nginx Issues](#nginx-issues)
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
| `validate-nginx-config.ps1` | nginx fails to start or configuration errors | Tests nginx configuration syntax and common issues |
| `handle-encrypted-key.ps1` | nginx prompts for SSL key passwords | Converts encrypted SSL keys to unencrypted format |
| `extract-pfx.ps1` | Using PFX certificate files | Extracts certificate and key from PFX files |

### Usage Examples

```powershell
# Diagnose backend service issues
.\deployment\scripts\diagnose-backend-service.ps1 -FixCommonIssues

# Validate nginx configuration  
.\deployment\scripts\validate-nginx-config.ps1

# Fix SSL key password prompts
.\deployment\scripts\handle-encrypted-key.ps1 -KeyPath "C:\Cert\server.key"

# Extract certificates from PFX
.\deployment\scripts\extract-pfx.ps1 -PfxPath "C:\Cert\server.pfx"
```

## Common Issues and Solutions

### 1. "Service won't start" (No error messages)

**Symptoms**: Service fails to start with no logs or error messages

**Diagnosis**:
```powershell
# Run comprehensive diagnostics
.\deployment\scripts\diagnose-backend-service.ps1 -FixCommonIssues

# Check NSSM configuration
nssm dump ExcelAddinBackend
```

**Common Causes & Solutions**:
- **Wrong Python path**: NSSM using "python" instead of full path
  ```powershell
  # Fix: Update NSSM with correct Python path
  nssm set ExcelAddinBackend Application "C:\Python39\python.exe"
  ```
- **Missing dependencies**: Python packages not installed
  ```powershell
  cd C:\inetpub\wwwroot\ExcelAddin\backend
  poetry install
  ```
- **Wrong working directory**: Service can't find required files
  ```powershell
  nssm set ExcelAddinBackend AppDirectory "C:\inetpub\wwwroot\ExcelAddin\backend"
  ```

### 2. "nginx process closes immediately"

**Symptoms**: nginx terminates with alert: `the event "ngx_master_*" was not signaled for 5s`

**Solution**: Use Windows-optimized nginx configuration
```powershell
# Copy Windows-specific nginx config
copy deployment\nginx\nginx.conf.windows.template C:\nginx\conf\nginx.conf

# Key Windows optimizations:
# - daemon off;
# - worker_processes 1;  
# - use select;
```

### 3. "nginx requests password on startup"

**Symptoms**: nginx CLI prompts for SSL private key password

**Solution**: Convert encrypted key to unencrypted
```powershell
.\deployment\scripts\handle-encrypted-key.ps1 -KeyPath "C:\Cert\server.key" -OutputPath "C:\Cert"
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
# 1. Verify Python path
where python
nssm set ExcelAddinBackend Application "C:\Users\AppData\Local\Programs\Python\Python39\python.exe"

# 2. Check working directory
nssm set ExcelAddinBackend AppDirectory "C:\inetpub\wwwroot\ExcelAddin\backend"

# 3. Verify service wrapper exists
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

**Issue**: nginx won't run as Windows service
```powershell
# Ensure daemon off in nginx.conf
daemon off;

# Configure NSSM properly  
nssm set nginx Application "C:\nginx\nginx.exe"
nssm set nginx AppParameters "-g \"daemon off;\""
```

## nginx Issues

### Windows-Specific nginx Problems

#### 1. Process Management Issues
```nginx
# nginx.conf optimizations for Windows
daemon off;                    # Required for NSSM
worker_processes 1;           # Single worker for Windows
events {
    use select;               # Windows-compatible event method
    worker_connections 1024;
}
```

#### 2. Path and File Issues
```nginx
# Use forward slashes in paths (even on Windows)
ssl_certificate C:/Cert/server.crt;
ssl_private_key C:/Cert/server.key;

# Use alias instead of root for subpaths
location /excellence/ {
    alias C:/inetpub/wwwroot/ExcelAddin/;
}
```

#### 3. Service Configuration Issues
```powershell
# Correct NSSM configuration for nginx
nssm install nginx "C:\nginx\nginx.exe"
nssm set nginx AppDirectory "C:\nginx"  
nssm set nginx AppParameters "-g \"daemon off;\""
nssm set nginx DisplayName "nginx Web Server"
nssm set nginx Description "nginx HTTP server"
```

### nginx Log Analysis
```powershell
# Check nginx error logs
Get-Content C:\nginx\logs\error.log -Tail 50

# Check nginx access logs  
Get-Content C:\nginx\logs\access.log -Tail 50

# Monitor logs in real-time
Get-Content C:\nginx\logs\error.log -Wait -Tail 10
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

# Set correct path in NSSM
$pythonPath = (Get-Command python).Source
nssm set ExcelAddinBackend Application $pythonPath
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

#### 3. Port Conflicts
```powershell
# Check what's using port 5000
netstat -ano | findstr :5000

# Kill processes using the port
taskkill /PID <process_id> /F
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
# Check all services
Get-Service ExcelAddinBackend, nginx

# View service configuration
nssm dump ExcelAddinBackend
nssm dump nginx

# View service logs
Get-Content C:\inetpub\wwwroot\ExcelAddin\backend\logs\service.log -Tail 50
Get-Content C:\nginx\logs\error.log -Tail 50
```

### Network Diagnostic Commands
```powershell
# Test ports
Test-NetConnection -ComputerName localhost -Port 5000  # Backend
Test-NetConnection -ComputerName localhost -Port 8443  # nginx

# Check listening ports
netstat -an | findstr :5000
netstat -an | findstr :8443

# Test SSL
openssl s_client -connect server01.intranet.local:8443
```

### Application Diagnostic Commands
```powershell
# Test backend API
curl http://localhost:5000/api/health

# Test through nginx proxy  
curl -k https://localhost:8443/excellence/api/health

# Test frontend application
curl -k https://localhost:8443/excellence/

# Validate nginx configuration
C:\nginx\nginx.exe -t
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
Backend Service Logs:
  C:\inetpub\wwwroot\ExcelAddin\backend\logs\service.log
  C:\inetpub\wwwroot\ExcelAddin\backend\logs\error.log

nginx Logs:  
  C:\nginx\logs\access.log
  C:\nginx\logs\error.log

Windows Event Logs:
  Event Viewer → Windows Logs → Application
  Look for "ExcelAddinBackend" and "nginx" sources
```

This troubleshooting guide covers the most common issues and their solutions. For additional support, use the provided diagnostic scripts which can automatically detect and fix many common configuration problems.