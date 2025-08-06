# Quick Fix Guide for nginx Windows Issues

This guide provides immediate solutions to the three critical issues reported with nginx on Windows Server.

## Issue 1: Process Closes with Master Process Alert

**Symptom:** nginx process closes with only: `[alert] the event "ngx_master_*" was not signaled for 5s`

### Immediate Fix:
```powershell
# Step 1: Copy the Windows-optimized nginx configuration
Copy-Item "deployment\nginx\nginx.conf.windows.template" "C:\nginx\conf\nginx.conf" -Force

# Step 2: Test the configuration
C:\nginx\nginx.exe -t

# Step 3: Try running nginx directly (should not close immediately)
cd C:\nginx
.\nginx.exe
```

If this fixes the immediate closing issue, proceed to set up as a Windows service.

## Issue 2: Password Prompt on Startup

**Symptom:** nginx prompts for password when starting

### Immediate Fix:
```powershell
# Check if your key is encrypted
.\deployment\scripts\handle-encrypted-key.ps1 -KeyPath "C:\Cert\server.key" -TestOnly

# If encrypted, convert to unencrypted format
.\deployment\scripts\handle-encrypted-key.ps1 -KeyPath "C:\Cert\server.key"

# Update nginx configuration to use the unencrypted key
# Edit C:\nginx\conf\excel-addin.conf and change:
# ssl_certificate_key C:/Cert/server-unencrypted.key;
```

### Alternative - Using OpenSSL directly:
```bash
# Convert encrypted key to unencrypted
openssl rsa -in "C:\Cert\server.key" -out "C:\Cert\server-unencrypted.key"
```

## Issue 3: Running nginx with NSSM

**Claim:** "nginx cannot be run with NSSM"  
**Truth:** nginx CAN be run with NSSM using proper configuration.

### Setup nginx as Windows Service:
```powershell
# Automated setup (recommended)
.\deployment\scripts\setup-nginx-service.ps1 -NginxPath "C:\nginx" -Force

# This script:
# 1. Installs nginx as a Windows service using NSSM
# 2. Configures proper service settings for Windows compatibility
# 3. Sets up logging and recovery options
# 4. Starts the service automatically
```

### Manual NSSM Setup (if automated script fails):
```cmd
# Install NSSM service
nssm install nginx "C:\nginx\nginx.exe"
nssm set nginx AppDirectory "C:\nginx"
nssm set nginx DisplayName "nginx Web Server"

# CRITICAL: Configure for Windows service compatibility
nssm set nginx AppNoConsole 1
nssm set nginx AppPriority NORMAL_PRIORITY_CLASS

# Start the service
net start nginx
```

## Complete Solution Workflow

Run these commands in order to solve all three issues:

```powershell
# 1. Fix configuration for Windows stability
Write-Host "Fixing nginx configuration..." -ForegroundColor Yellow
Copy-Item "deployment\nginx\nginx.conf.windows.template" "C:\nginx\conf\nginx.conf" -Force
C:\nginx\nginx.exe -t

# 2. Handle encrypted key if present
Write-Host "Checking SSL key encryption..." -ForegroundColor Yellow
if (Test-Path "C:\Cert\server.key") {
    .\deployment\scripts\handle-encrypted-key.ps1 -KeyPath "C:\Cert\server.key" -TestOnly
    $encrypted = $LASTEXITCODE -ne 0
    if ($encrypted) {
        Write-Host "Converting encrypted key..." -ForegroundColor Yellow
        .\deployment\scripts\handle-encrypted-key.ps1 -KeyPath "C:\Cert\server.key"
    }
}

# 3. Set up as Windows service with NSSM
Write-Host "Setting up Windows service..." -ForegroundColor Yellow
.\deployment\scripts\setup-nginx-service.ps1 -NginxPath "C:\nginx" -Force

# 4. Verify everything is working
Write-Host "Verifying setup..." -ForegroundColor Yellow
Start-Sleep 5
Get-Service nginx
Invoke-WebRequest "https://server01.intranet.local:8443/excellence/health" -SkipCertificateCheck
```

## Verification Steps

After applying the fixes:

1. **Check service status:**
   ```powershell
   Get-Service nginx
   # Should show "Running"
   ```

2. **Test nginx response:**
   ```powershell
   # Test health endpoint
   Invoke-WebRequest "https://server01.intranet.local:8443/excellence/health" -SkipCertificateCheck
   ```

3. **Check logs for errors:**
   ```powershell
   # Service logs
   Get-Content "C:\Logs\nginx\service-stdout.log" -Tail 10
   
   # nginx error log
   Get-Content "C:\nginx\logs\error.log" -Tail 10
   ```

4. **Verify no password prompts:**
   ```powershell
   # Restart service - should not prompt for passwords
   Restart-Service nginx
   ```

## Expected Results

After implementing these fixes:
- ✅ nginx process should run continuously without closing
- ✅ No password prompts during service startup
- ✅ nginx runs as a Windows service with automatic startup
- ✅ Service can be managed with standard Windows service commands
- ✅ Logs are properly written to designated directories

## Troubleshooting

If issues persist:

1. **Check Windows Event Viewer:**
   ```
   Windows Logs → Application
   Windows Logs → System
   ```

2. **Review all log files:**
   - `C:\nginx\logs\error.log`
   - `C:\Logs\nginx\service-stdout.log` 
   - `C:\Logs\nginx\service-stderr.log`

3. **Test nginx configuration:**
   ```cmd
   cd C:\nginx
   nginx.exe -t
   ```

4. **Verify file permissions:**
   ```powershell
   # Ensure nginx can read certificate files
   icacls "C:\Cert\*.crt" /grant "Everyone:(R)"
   icacls "C:\Cert\*-unencrypted.key" /grant "Everyone:(R)"
   ```

These solutions address the core Windows-specific issues with nginx deployment and provide a stable, production-ready configuration.