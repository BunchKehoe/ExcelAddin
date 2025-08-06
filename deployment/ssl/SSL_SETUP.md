# SSL Certificate Setup for Excel Add-in on Windows Server
# This guide covers different SSL certificate setup options

## Company Certificate Setup (Recommended for Production)

If you have company certificates located in `C:\Cert\`:
- Root CA certificate: `cacert.pem` (company root CA in .cer and .pem format)
- Server certificate: `server.crt` and `server.key` OR `server.pfx`

### Using Separate Certificate Files (.crt and .key)
```nginx
# nginx SSL configuration in excel-addin.conf
ssl_certificate C:/Cert/server.crt;
ssl_certificate_key C:/Cert/server.key;
ssl_trusted_certificate C:/Cert/cacert.pem;
```

### Using .pfx Certificate File
If you have a .pfx file, extract the certificate and key:

```powershell
# Extract certificate from .pfx file
openssl pkcs12 -in C:/Cert/server.pfx -out C:/Cert/server.crt -clcerts -nokeys -passin pass:YOUR_PFX_PASSWORD

# Extract private key from .pfx file
openssl pkcs12 -in C:/Cert/server.pfx -out C:/Cert/server.key -nocerts -nodes -passin pass:YOUR_PFX_PASSWORD

# If openssl is not available on Windows, use PowerShell:
$pfx = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2("C:\Cert\server.pfx", "YOUR_PFX_PASSWORD", "Exportable")
[System.IO.File]::WriteAllBytes("C:\Cert\server.crt", $pfx.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))
# Note: Private key extraction from PowerShell requires additional steps
```

### Microsoft Add-in Certificate Integration

For production deployment, you need to handle both the server SSL certificates and Excel add-in manifest certificates:

#### 1. Server SSL Certificates (for HTTPS communication)
- Use your company's server certificate (`server.crt`/`server.key` or `server.pfx`) 
- This handles browser-to-server HTTPS communication
- Configure in nginx as shown above

#### 2. Excel Add-in Manifest Certificates (for add-in trust)
The Microsoft add-in certificates you mentioned are likely for:
- Code signing the add-in manifest
- Establishing trust with Office/Excel

**For production, you have several options:**

**Option A: Use Company Root CA (Recommended)**
If your company root CA is already trusted in your environment:
1. Generate a new certificate signed by your company CA specifically for the add-in
2. Update the manifest to reference URLs using your company-signed certificate
3. No additional client configuration needed if company CA is in Windows trust store

**Option B: Install Microsoft Add-in CA in Production**
1. Export the Microsoft add-in root CA certificate
2. Install it in the Windows Certificate Store on all client machines
3. Use Group Policy to distribute the CA certificate across the organization

**Option C: Use Company-Signed Certificate for Add-in**
1. Create a new certificate signed by your company CA for the add-in URLs
2. Update manifest-staging.xml to use production URLs with company certificate
3. This leverages your existing company trust infrastructure

#### Certificate Trust Chain Setup
```powershell
# Install company root CA in Windows Certificate Store (if not already present)
Import-Certificate -FilePath "C:\Cert\cacert.pem" -CertStoreLocation Cert:\LocalMachine\Root

# Verify certificate chain
Get-ChildItem -Path Cert:\LocalMachine\Root | Where-Object { $_.Subject -like "*Your Company*" }
```

#### Manifest Configuration for Production
Update your manifest-staging.xml to use production URLs:
```xml
<!-- All URLs should use your production domain with company certificate -->
<SourceLocation DefaultValue="https://server01.intranet.local:8443/excellence/taskpane.html"/>
<SupportUrl DefaultValue="https://server01.intranet.local:8443/excellence/support"/>
```

### Certificate File Structure for Production
```
C:\Cert\
├── cacert.pem              # Company root CA certificate
├── server.crt              # Server certificate for HTTPS
├── server.key              # Server private key
├── server.pfx              # Alternative: combined certificate file
└── addin-ca.crt            # Optional: specific add-in CA certificate
```

## Option 1: Self-Signed Certificate (Development/Internal Testing Only)

### Create Self-Signed Certificate using PowerShell
```powershell
# Create self-signed certificate for development/testing
$cert = New-SelfSignedCertificate -DnsName "your-staging-domain.com" -CertStoreLocation "cert:\LocalMachine\My" -KeyLength 2048 -Provider "Microsoft RSA SChannel Cryptographic Provider" -KeyExportPolicy Exportable -KeyUsage DigitalSignature,KeyEncipherment -Type SSLServerAuthentication

# Export certificate to files
$certPath = "C:\ssl\certs"
New-Item -ItemType Directory -Path $certPath -Force

# Export certificate (.crt file)
$certBytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
[System.IO.File]::WriteAllBytes("$certPath\excel-addin.crt", $certBytes)

# Export private key (.key file) - requires certificate to be exportable
$keyBytes = $cert.PrivateKey.ExportCspBlob($true)
[System.IO.File]::WriteAllBytes("$certPath\excel-addin.key", $keyBytes)
```

### Using OpenSSL (if installed)
```bash
# Create private key
openssl genrsa -out excel-addin.key 2048

# Create certificate signing request
openssl req -new -key excel-addin.key -out excel-addin.csr -subj "/C=US/ST=State/L=City/O=Organization/CN=your-staging-domain.com"

# Create self-signed certificate
openssl x509 -req -in excel-addin.csr -signkey excel-addin.key -out excel-addin.crt -days 365
```

## Option 2: Let's Encrypt (Recommended for Internet-facing)

### Using Certbot with nginx
```bash
# Install Certbot (if using WSL or Linux subsystem)
sudo apt-get update
sudo apt-get install certbot python3-certbot-nginx

# Generate certificate
sudo certbot --nginx -d your-staging-domain.com

# Auto-renewal setup
sudo crontab -e
# Add: 0 12 * * * /usr/bin/certbot renew --quiet
```

### Using win-acme (Windows-native Let's Encrypt client)
```powershell
# Download win-acme from https://www.win-acme.com/
# Run the executable and follow the interactive setup
# Choose option for nginx integration
```

## Option 3: Commercial Certificate

### Certificate Signing Request (CSR) Generation
```powershell
# Generate CSR using PowerShell
$subject = "CN=your-staging-domain.com,O=Your Organization,L=City,S=State,C=US"
certreq -new -f -q -config - -attrib "CertificateTemplate:WebServer" @"
[NewRequest]
Subject="$subject"
KeyLength=2048
KeySpec=1
Exportable=TRUE
MachineKeySet=TRUE
SMIME=FALSE
RequestType=PKCS10
KeyUsage=0xa0

[RequestAttributes]
"@
```

### Install Commercial Certificate
1. Purchase certificate from CA (GoDaddy, DigiCert, etc.)
2. Generate CSR using above method or CA's tools
3. Submit CSR to CA
4. Download certificate files from CA
5. Install certificate in Windows Certificate Store
6. Export certificate and private key for nginx

## nginx SSL Configuration

Update the nginx configuration file:
```nginx
server {
    listen 443 ssl http2;
    server_name your-staging-domain.com;
    
    # SSL Certificate paths (update with your actual paths)
    ssl_certificate C:/ssl/certs/excel-addin.crt;
    ssl_certificate_key C:/ssl/private/excel-addin.key;
    
    # SSL Security Configuration
    ssl_protocols TLSv1.2 TLSv1.3;
    ssl_ciphers ECDHE-RSA-AES256-GCM-SHA512:DHE-RSA-AES256-GCM-SHA512:ECDHE-RSA-AES256-GCM-SHA384;
    ssl_prefer_server_ciphers off;
    ssl_session_cache shared:SSL:10m;
    ssl_session_timeout 10m;
    
    # Additional SSL security
    ssl_stapling on;
    ssl_stapling_verify on;
    
    # Security headers
    add_header Strict-Transport-Security "max-age=31536000; includeSubDomains; preload" always;
    
    # Rest of configuration...
}
```

## Certificate File Structure

Recommended directory structure:
```
C:\ssl\
├── certs\
│   ├── excel-addin.crt      # Certificate file
│   └── ca-bundle.crt        # CA intermediate certificates (if applicable)
├── private\
│   └── excel-addin.key      # Private key file
└── csr\
    └── excel-addin.csr      # Certificate signing request (backup)
```

## PowerShell Certificate Management Script

```powershell
# Certificate management helper script
param(
    [string]$Action = "list",
    [string]$DomainName = "your-staging-domain.com",
    [string]$CertPath = "C:\ssl\certs",
    [string]$KeyPath = "C:\ssl\private"
)

function List-Certificates {
    Write-Host "SSL Certificates in LocalMachine\My store:"
    Get-ChildItem -Path "Cert:\LocalMachine\My" | Where-Object { $_.Subject -like "*$DomainName*" } | Format-Table Thumbprint, Subject, NotAfter
}

function Test-Certificate {
    param($Domain)
    try {
        $cert = Get-ChildItem -Path "Cert:\LocalMachine\My" | Where-Object { $_.DnsNameList -contains $Domain } | Select-Object -First 1
        if ($cert) {
            Write-Host "Certificate found for $Domain" -ForegroundColor Green
            Write-Host "Expires: $($cert.NotAfter)" -ForegroundColor Yellow
            Write-Host "Thumbprint: $($cert.Thumbprint)" -ForegroundColor Cyan
            
            if ($cert.NotAfter -lt (Get-Date).AddDays(30)) {
                Write-Host "WARNING: Certificate expires in less than 30 days!" -ForegroundColor Red
            }
            
            return $true
        } else {
            Write-Host "No certificate found for $Domain" -ForegroundColor Red
            return $false
        }
    } catch {
        Write-Host "Error checking certificate: $_" -ForegroundColor Red
        return $false
    }
}

function Export-CertificateFiles {
    param($Domain, $OutputPath, $KeyOutputPath)
    
    $cert = Get-ChildItem -Path "Cert:\LocalMachine\My" | Where-Object { $_.DnsNameList -contains $Domain } | Select-Object -First 1
    if (-not $cert) {
        Write-Host "No certificate found for $Domain" -ForegroundColor Red
        return $false
    }
    
    try {
        # Create directories
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        New-Item -ItemType Directory -Path $KeyOutputPath -Force | Out-Null
        
        # Export certificate
        $certBytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
        [System.IO.File]::WriteAllBytes("$OutputPath\excel-addin.crt", $certBytes)
        Write-Host "Certificate exported to $OutputPath\excel-addin.crt" -ForegroundColor Green
        
        # Note: Private key export requires the certificate to be marked as exportable
        # This is more complex and may require additional tools or manual export
        Write-Host "Private key export may require manual steps or additional tools" -ForegroundColor Yellow
        
        return $true
    } catch {
        Write-Host "Error exporting certificate: $_" -ForegroundColor Red
        return $false
    }
}

switch ($Action.ToLower()) {
    "list" { List-Certificates }
    "test" { Test-Certificate -Domain $DomainName }
    "export" { Export-CertificateFiles -Domain $DomainName -OutputPath $CertPath -KeyOutputPath $KeyPath }
    default { 
        Write-Host "Usage: .\ssl-setup.ps1 -Action [list|test|export] -DomainName your-domain.com"
        Write-Host "Available actions:"
        Write-Host "  list   - List certificates in machine store"
        Write-Host "  test   - Test if certificate exists and check expiration"
        Write-Host "  export - Export certificate to files for nginx"
    }
}
```

## Testing SSL Configuration

### Test SSL certificate installation:
```powershell
# Test local SSL endpoint
Invoke-WebRequest -Uri "https://your-staging-domain.com/health" -UseBasicParsing

# Test certificate chain
openssl s_client -connect your-staging-domain.com:443 -servername your-staging-domain.com
```

### Common SSL Issues and Solutions:

1. **Certificate not trusted by Office**: Ensure the certificate is from a trusted CA or properly installed in the system trust store.

2. **nginx SSL errors**: Check certificate file paths and permissions.

3. **Mixed content warnings**: Ensure all resources (CSS, JS, images) are served over HTTPS.

4. **Certificate chain issues**: Include intermediate certificates in the certificate file.

## Automated Certificate Renewal

Create a scheduled task for certificate renewal:
```powershell
# Create scheduled task for certificate renewal (adapt based on your certificate provider)
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File C:\Scripts\renew-certificate.ps1"
$trigger = New-ScheduledTaskTrigger -Daily -At "2:00AM"
$principal = New-ScheduledTaskPrincipal -UserID "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
Register-ScheduledTask -TaskName "Excel-Addin-Cert-Renewal" -Action $action -Trigger $trigger -Principal $principal
```