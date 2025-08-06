# nginx Windows Troubleshooting Guide

This guide addresses common nginx warnings and issues specific to Windows Server deployments.

## Common Warnings and Their Solutions

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