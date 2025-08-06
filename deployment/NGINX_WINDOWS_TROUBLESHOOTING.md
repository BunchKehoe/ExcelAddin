# nginx Windows Troubleshooting Guide

This guide addresses common nginx warnings and issues specific to Windows Server deployments.

## CRITICAL ISSUES AND SOLUTIONS

### Issue 1: Process Closes with Master Process Alert

**Problem:**
nginx process closes immediately with only the warning:
```
[alert] the event "ngx_master_*" was not signaled for 5s
```

**Root Causes:**
- Windows-specific process synchronization issues
- nginx daemon mode conflicts with Windows service management
- Worker process configuration incompatible with Windows

**Solutions:**

#### Solution A: Use Windows-Optimized Configuration (Recommended)
1. Replace your nginx.conf with the optimized template:
   ```powershell
   Copy-Item "deployment\nginx\nginx.conf.windows.template" "C:\nginx\conf\nginx.conf"
   ```

2. Key optimizations in the template:
   ```nginx
   # Single worker for Windows stability
   worker_processes 1;
   
   # Daemon off for service compatibility
   daemon off;
   
   events {
       use select;           # Windows-compatible event method
       worker_connections 512;  # Reduced for stability
       accept_mutex_delay 100ms;
       accept_mutex on;
   }
   ```

#### Solution B: Run nginx as a Windows Service with NSSM
1. Install NSSM (Non-Sucking Service Manager):
   ```powershell
   # Download from https://nssm.cc/
   # Or use: winget install NSSM.NSSM
   ```

2. Use the service setup script:
   ```powershell
   .\deployment\scripts\setup-nginx-service.ps1 -NginxPath "C:\nginx"
   ```

3. This automatically configures nginx for service operation with proper Windows compatibility.

### Issue 2: Password Prompt on Startup

**Problem:**
nginx requests password in CLI on startup for encrypted private key files.

**Root Cause:**
The SSL private key file (server.key) is encrypted with a password.

**Solutions:**

#### Solution A: Convert to Unencrypted Key (Recommended)
1. Use the provided script to convert the key:
   ```powershell
   .\deployment\scripts\handle-encrypted-key.ps1 -KeyPath "C:\Cert\server.key"
   ```

2. Update nginx configuration to use the unencrypted key:
   ```nginx
   ssl_certificate_key C:/Cert/server-unencrypted.key;
   ```

#### Solution B: Extract Unencrypted Key from PFX
If you have a .pfx file:
1. Extract unencrypted certificate and key:
   ```powershell
   .\deployment\scripts\extract-pfx.ps1 -PfxPath "C:\Cert\server.pfx" -OutputDir "C:\Cert"
   ```

2. Use OpenSSL to ensure key is unencrypted:
   ```bash
   openssl rsa -in "C:\Cert\server.key" -out "C:\Cert\server-unencrypted.key"
   ```

#### Solution C: Generate New Unencrypted Key
Create a new private key without password:
```bash
openssl genrsa -out "C:\Cert\server-new.key" 2048
```

### Issue 3: nginx with NSSM Service Management

**Problem:**
Some sources claim nginx cannot be run with NSSM (Non-Sucking Service Manager).

**Truth:** nginx CAN be run with NSSM using proper configuration.

**Solution:**

#### Automated Setup
Use the provided service setup script:
```powershell
.\deployment\scripts\setup-nginx-service.ps1 -NginxPath "C:\nginx" -Force
```

#### Manual NSSM Configuration
1. Install the service:
   ```cmd
   nssm install nginx "C:\nginx\nginx.exe"
   nssm set nginx AppDirectory "C:\nginx"
   nssm set nginx AppParameters "-c C:\nginx\conf\nginx.conf"
   ```

2. Configure for Windows service compatibility:
   ```cmd
   nssm set nginx AppPriority NORMAL_PRIORITY_CLASS
   nssm set nginx AppNoConsole 1
   nssm set nginx AppStopMethodConsole 10000
   nssm set nginx AppStopMethodWindow 10000
   ```

3. **CRITICAL:** Ensure nginx.conf has `daemon off;` setting:
   ```nginx
   # This is essential for NSSM service operation
   daemon off;
   ```

4. Start the service:
   ```powershell
   Start-Service nginx
   ```

## COMPREHENSIVE SOLUTION WORKFLOW

### Step 1: Fix Configuration Issues
```powershell
# Copy optimized configuration
Copy-Item "deployment\nginx\nginx.conf.windows.template" "C:\nginx\conf\nginx.conf"

# Test configuration
C:\nginx\nginx.exe -t
```

### Step 2: Handle Encrypted Keys
```powershell
# Check if key is encrypted
.\deployment\scripts\handle-encrypted-key.ps1 -KeyPath "C:\Cert\server.key" -TestOnly

# Convert if encrypted
.\deployment\scripts\handle-encrypted-key.ps1 -KeyPath "C:\Cert\server.key"
```

### Step 3: Set Up Windows Service
```powershell
# Install nginx as Windows service with NSSM
.\deployment\scripts\setup-nginx-service.ps1 -NginxPath "C:\nginx" -Force
```

### Step 4: Verify Operation
```powershell
# Check service status
Get-Service nginx

# Test HTTP response
Invoke-WebRequest "http://127.0.0.1/health"

# Check logs
Get-Content "C:\Logs\nginx\service-stdout.log" -Tail 20
```

## LEGACY WARNINGS (Already Fixed in Configuration)

### 1. Deprecated HTTP/2 Syntax Warning

**Warning:**
```
[warn] the "listen ... http2" directive is deprecated, use the "http2" directive instead
```

**Solution:**
The configuration has been updated to use the new HTTP/2 syntax. Instead of:
```nginx
listen 8443 ssl http2;
```

We now use:
```nginx
listen 8443 ssl;
http2 on;
```

### 2. SSL Stapling Warning

**Warning:**
```
[warn] "ssl_stapling" ignored, no OCSP responder URL in the certificate
```

**Cause:** Company-issued certificates often don't include OCSP responder URLs, making SSL stapling unavailable.

**Solution:**
SSL stapling has been disabled in the configuration:
```nginx
# SSL stapling disabled for company certificates without OCSP responder
# ssl_stapling on;
# ssl_stapling_verify on;
```

**To re-enable (if your certificate supports OCSP):**
1. Uncomment the ssl_stapling lines
2. Ensure your certificate includes an OCSP responder URL
3. Verify network connectivity to the OCSP responder

### 3. Master Process Alert

**Alert:**
```
[alert] the event "ngx_master_*" was not signaled for 5s
```

**Cause:** This Windows-specific issue can be caused by:
- Process/thread synchronization issues
- Insufficient system resources
- Windows event handling problems

**Solutions:**

#### Option 1: Use Windows-Optimized Configuration
Replace your `nginx.conf` with our Windows-optimized template:

```powershell
Copy-Item "deployment\nginx\nginx.conf.windows.template" "C:\nginx\conf\nginx.conf"
```

#### Option 2: Adjust Worker Processes
In your `nginx.conf`, try reducing worker processes:
```nginx
worker_processes 1;  # Instead of 'auto'
```

#### Option 3: Use Select Event Method
Ensure your `nginx.conf` events block uses:
```nginx
events {
    use select;
    worker_connections 1024;
    accept_mutex_delay 500ms;
}
```

#### Option 4: Windows Service Configuration
If running nginx as a Windows service, try these NSSM settings:
```powershell
nssm set nginx AppPriority NORMAL_PRIORITY_CLASS
nssm set nginx AppNoConsole 1
nssm set nginx AppStdout C:\Logs\nginx\service-stdout.log
nssm set nginx AppStderr C:\Logs\nginx\service-stderr.log
```

## Validation and Testing

### 1. Test Configuration Syntax
```powershell
C:\nginx\nginx.exe -t -c C:\nginx\conf\nginx.conf
```

### 2. Run Validation Script
```powershell
.\deployment\scripts\validate-nginx-config.ps1
```

### 3. Check nginx Logs
- Error log: `C:\nginx\logs\error.log`
- Access log: `C:\nginx\logs\access.log`
- Application log: `C:\Logs\nginx\excel_addin_error.log`

### 4. Monitor Windows Event Viewer
Check Windows Event Viewer under:
- Windows Logs → Application
- Windows Logs → System

## Performance Optimization for Windows

### nginx.conf Optimizations
```nginx
# Windows-specific optimizations
worker_processes 1;
worker_priority 0;
worker_rlimit_nofile 65535;

events {
    use select;
    worker_connections 1024;
    accept_mutex_delay 500ms;
}

http {
    # Windows TCP optimizations
    sendfile on;
    tcp_nopush on;
    tcp_nodelay on;
    
    # Connection handling
    keepalive_timeout 65;
    keepalive_requests 1000;
    
    # Buffer sizes
    client_body_buffer_size 128k;
    client_max_body_size 50m;
    client_header_buffer_size 1k;
    large_client_header_buffers 4 4k;
}
```

## Windows Service Management

### Create nginx Service
```powershell
# Install NSSM if not already installed
# Download from https://nssm.cc/

# Create the service
nssm install nginx "C:\nginx\nginx.exe"
nssm set nginx AppDirectory "C:\nginx"
nssm set nginx AppParameters "-c C:\nginx\conf\nginx.conf"
nssm set nginx DisplayName "nginx Web Server"
nssm set nginx Description "nginx HTTP and reverse proxy server"
nssm set nginx Start SERVICE_AUTO_START
```

### Service Operations
```powershell
# Start service
Start-Service nginx

# Stop service
Stop-Service nginx

# Restart service
Restart-Service nginx

# Check status
Get-Service nginx
```

## Certificate-Related Issues

### Certificate Loading Problems
If certificates are not loading properly:

1. **Check file permissions:**
   ```powershell
   icacls "C:\Cert\server.crt" /grant "Everyone:(R)"
   icacls "C:\Cert\server.key" /grant "Everyone:(R)"
   ```

2. **Verify certificate format:**
   ```powershell
   # Test certificate validity
   $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2("C:\Cert\server.crt")
   $cert.Subject
   $cert.NotAfter
   ```

3. **Extract from PFX if needed:**
   ```powershell
   .\deployment\scripts\extract-pfx.ps1 -PfxPath "C:\Cert\server.pfx" -OutputPath "C:\Cert"
   ```

## Troubleshooting Checklist

- [ ] nginx configuration syntax is valid (`nginx -t`)
- [ ] All certificate files are present and readable
- [ ] Windows firewall allows port 8443
- [ ] Required directories exist with proper permissions
- [ ] Backend service is running on port 5000
- [ ] No conflicting services on ports 80/8443
- [ ] Event Viewer shows no critical errors
- [ ] nginx error log shows no SSL/certificate issues

## Getting Help

If issues persist:

1. Run the validation script: `.\deployment\scripts\validate-nginx-config.ps1`
2. Check all log files mentioned above
3. Verify Windows system requirements are met
4. Consider using single worker process configuration for stability
5. Test with minimal configuration first, then add features incrementally

## Additional Resources

- [nginx Windows Documentation](http://nginx.org/en/docs/windows.html)
- [NSSM Documentation](https://nssm.cc/usage)
- Windows Event Viewer (eventvwr.msc)
- Windows Services Manager (services.msc)