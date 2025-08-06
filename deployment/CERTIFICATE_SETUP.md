# Windows Server Certificate Deployment Guide

This guide provides step-by-step instructions for configuring SSL certificates for the Excel Add-in on Windows Server with nginx.

## Prerequisites

- Windows Server 2016 or later
- nginx for Windows installed (https://nginx.org/en/download.html)
- Company certificates in `C:\Cert\` directory:
  - `cacert.pem` - Company root CA certificate
  - `server.crt` and `server.key` - Server certificate and private key
  - OR `server.pfx` - Combined certificate file

## Step 1: Install and Configure nginx

### Download and Install nginx for Windows
1. Download nginx for Windows from https://nginx.org/en/download.html
2. Extract to `C:\nginx\`
3. Test basic installation:
```cmd
cd C:\nginx
nginx -t
```

### Replace nginx Configuration
1. Backup existing configuration:
```cmd
copy C:\nginx\conf\nginx.conf C:\nginx\conf\nginx.conf.backup
```

2. Replace with the Windows template:
```cmd
copy deployment\nginx\nginx.conf.windows.template C:\nginx\conf\nginx.conf
```

3. Copy the Excel add-in configuration:
```cmd
copy deployment\nginx\excel-addin.conf C:\nginx\conf\excel-addin.conf
```

## Step 2: Prepare Certificates

### Option A: If you have server.crt and server.key files
Your certificates are ready to use. Ensure they are in `C:\Cert\`:
```
C:\Cert\
├── cacert.pem     # Company root CA
├── server.crt     # Server certificate  
└── server.key     # Private key
```

### Option B: If you have server.pfx file
Extract certificate and key from the .pfx file:

**Using OpenSSL (if available):**
```cmd
# Extract certificate
openssl pkcs12 -in C:\Cert\server.pfx -out C:\Cert\server.crt -clcerts -nokeys

# Extract private key  
openssl pkcs12 -in C:\Cert\server.pfx -out C:\Cert\server.key -nocerts -nodes
```

**Using PowerShell (if OpenSSL not available):**
```powershell
# This extracts only the certificate; key extraction is more complex
$pfxPassword = Read-Host "Enter PFX password" -AsSecureString
$pfx = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2("C:\Cert\server.pfx", $pfxPassword, "Exportable")
[System.IO.File]::WriteAllBytes("C:\Cert\server.crt", $pfx.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))

# For private key extraction, you may need to use the Windows Certificate Manager
# or a tool like Win32 OpenSSL
```

## Step 3: Install Root CA Certificate (if needed)

If the company root CA is not already trusted on the server:
```powershell
# Import company root CA into Windows certificate store
Import-Certificate -FilePath "C:\Cert\cacert.pem" -CertStoreLocation Cert:\LocalMachine\Root

# Verify installation
Get-ChildItem -Path Cert:\LocalMachine\Root | Where-Object { $_.Subject -like "*Your Company*" }
```

## Step 4: Configure nginx for Certificate Paths

The nginx configuration is already set up to use the certificate paths. Verify the paths in `C:\nginx\conf\excel-addin.conf`:

```nginx
# SSL Configuration - paths should match your certificate location
ssl_certificate C:/Cert/server.crt;
ssl_certificate_key C:/Cert/server.key;
ssl_trusted_certificate C:/Cert/cacert.pem;
```

## Step 5: Test nginx Configuration

1. Test the configuration syntax:
```cmd
cd C:\nginx
nginx -t
```

2. If successful, start nginx:
```cmd
nginx
```

3. If nginx is already running, reload the configuration:
```cmd
nginx -s reload
```

## Step 6: Verify SSL Configuration

### Test basic connectivity:
```cmd
# Test the health endpoint
curl -k https://server01.intranet.local:8443/excellence/health

# Test certificate chain
openssl s_client -connect server01.intranet.local:8443 -servername server01.intranet.local
```

### Test from PowerShell:
```powershell
# Test HTTPS endpoint
Invoke-WebRequest -Uri "https://server01.intranet.local:8443/excellence/health" -UseBasicParsing

# Test certificate details
$webRequest = [Net.HttpWebRequest]::Create("https://server01.intranet.local:8443")
$webRequest.GetResponse()
$cert = $webRequest.ServicePoint.Certificate
$cert | Format-List *
```

## Step 7: Excel Add-in Certificate Trust

For the Excel add-in to work properly in production:

### Option A: Use Company CA (Recommended)
If your server certificate is signed by your company's root CA and that CA is trusted in your Windows environment, no additional steps are needed.

### Option B: Install Certificate on Client Machines
If you need to distribute the root CA to client machines:

```powershell
# Create a script to install the CA on client machines
# Save as install-ca.ps1 and deploy via Group Policy or SCCM

Import-Certificate -FilePath "\\server\share\cacert.pem" -CertStoreLocation Cert:\LocalMachine\Root
```

## Troubleshooting

### Common Issues:

1. **nginx fails to start with "upstream" error:**
   - Ensure `excel-addin.conf` is properly included in the main `nginx.conf` within the http block
   - The upstream directive must be at the http level, not server level

2. **SSL certificate errors:**
   - Verify certificate file paths and permissions
   - Ensure the certificate matches the domain name (server01.intranet.local)
   - Check that the private key matches the certificate

3. **Excel add-in trust issues:**
   - Verify the root CA is installed in the Windows certificate store
   - Check that the certificate chain is complete
   - Ensure all manifest URLs use HTTPS with the correct domain

### Log Locations:
- nginx error log: `C:\nginx\logs\error.log`
- nginx access log: `C:\nginx\logs\access.log`
- Excel add-in logs: `C:\Logs\ExcelAddin\`

### Validation Commands:
```cmd
# Test nginx configuration
nginx -t

# Check certificate expiration
certutil -verify C:\Cert\server.crt

# View certificate details
certutil -dump C:\Cert\server.crt
```

## Security Considerations

1. **File Permissions**: Ensure certificate files are readable by the nginx service account
2. **Private Key Security**: Protect the private key file with appropriate file permissions
3. **Certificate Expiration**: Set up monitoring for certificate expiration
4. **Regular Updates**: Keep nginx and certificates up to date

## Next Steps

After SSL is configured and working:
1. Deploy the Excel add-in files to `C:\inetpub\wwwroot\ExcelAddin\dist\`
2. Configure the Windows service for the backend API
3. Test the complete application flow
4. Set up monitoring and health checks

For complete deployment instructions, see `WINDOWS_DEPLOYMENT.md` and `EXCELLENCE_PATH_DEPLOYMENT.md`.