# PowerShell script to extract certificate and key from .pfx file
# Usage: .\extract-pfx.ps1 -PfxPath "C:\Cert\server.pfx" -OutputDir "C:\Cert"

param(
    [Parameter(Mandatory=$true)]
    [string]$PfxPath,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputDir = "C:\Cert",
    
    [Parameter(Mandatory=$false)]
    [SecureString]$Password
)

# Check if PFX file exists
if (-not (Test-Path $PfxPath)) {
    Write-Error "PFX file not found: $PfxPath"
    exit 1
}

# Create output directory if it doesn't exist
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    Write-Host "Created output directory: $OutputDir"
}

# Get password if not provided
if (-not $Password) {
    $Password = Read-Host "Enter PFX password" -AsSecureString
}

try {
    Write-Host "Loading PFX certificate..." -ForegroundColor Yellow
    
    # Load the PFX certificate
    $pfx = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($PfxPath, $Password, "Exportable,PersistKeySet")
    
    # Export certificate (.crt file)
    $certPath = Join-Path $OutputDir "server.crt"
    $certBytes = $pfx.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
    [System.IO.File]::WriteAllBytes($certPath, $certBytes)
    Write-Host "Certificate exported to: $certPath" -ForegroundColor Green
    
    # For private key extraction, we need to use a different approach
    # This requires the certificate to be installed in the certificate store temporarily
    Write-Host "Installing certificate temporarily for key extraction..." -ForegroundColor Yellow
    
    # Install certificate in personal store
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "LocalMachine")
    $store.Open("ReadWrite")
    $store.Add($pfx)
    $store.Close()
    
    # Find the installed certificate
    $installedCert = Get-ChildItem -Path "Cert:\LocalMachine\My" | Where-Object { $_.Thumbprint -eq $pfx.Thumbprint }
    
    if ($installedCert -and $installedCert.HasPrivateKey) {
        # Export private key using .NET APIs
        $keyPath = Join-Path $OutputDir "server.key"
        
        # Get the private key
        $privateKey = $installedCert.PrivateKey
        
        if ($privateKey -is [System.Security.Cryptography.RSACryptoServiceProvider]) {
            # Export RSA key in PKCS#8 format
            $keyBytes = $privateKey.ExportCspBlob($true)
            
            # Convert to PEM format (this is a simplified conversion)
            # For production, consider using a proper conversion method
            $keyBase64 = [System.Convert]::ToBase64String($keyBytes)
            $keyPem = "-----BEGIN PRIVATE KEY-----`n"
            for ($i = 0; $i -lt $keyBase64.Length; $i += 64) {
                $line = $keyBase64.Substring($i, [Math]::Min(64, $keyBase64.Length - $i))
                $keyPem += "$line`n"
            }
            $keyPem += "-----END PRIVATE KEY-----"
            
            [System.IO.File]::WriteAllText($keyPath, $keyPem)
            Write-Host "Private key exported to: $keyPath" -ForegroundColor Green
        } else {
            Write-Warning "Unable to export private key directly. Consider using OpenSSL:"
            Write-Host "openssl pkcs12 -in `"$PfxPath`" -out `"$keyPath`" -nocerts -nodes" -ForegroundColor Cyan
        }
        
        # Clean up - remove the temporarily installed certificate
        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("My", "LocalMachine")
        $store.Open("ReadWrite")
        $store.Remove($installedCert)
        $store.Close()
        
        Write-Host "Cleaned up temporary certificate installation" -ForegroundColor Gray
    } else {
        Write-Warning "Could not access private key. The certificate may not be exportable."
        Write-Host "Alternative: Use OpenSSL to extract the private key:" -ForegroundColor Cyan
        Write-Host "openssl pkcs12 -in `"$PfxPath`" -out `"$(Join-Path $OutputDir "server.key")`" -nocerts -nodes" -ForegroundColor Cyan
    }
    
    # Display certificate information
    Write-Host "`nCertificate Information:" -ForegroundColor Cyan
    Write-Host "Subject: $($pfx.Subject)"
    Write-Host "Issuer: $($pfx.Issuer)"
    Write-Host "Valid From: $($pfx.NotBefore)"
    Write-Host "Valid To: $($pfx.NotAfter)"
    Write-Host "Serial Number: $($pfx.SerialNumber)"
    
    if ($pfx.NotAfter -lt (Get-Date).AddDays(30)) {
        Write-Warning "Certificate expires in less than 30 days!"
    }
    
    Write-Host "`nFiles created in $OutputDir :" -ForegroundColor Green
    Get-ChildItem -Path $OutputDir -Filter "server.*" | Format-Table Name, Length, LastWriteTime
    
} catch {
    Write-Error "Failed to extract certificate: $_"
    exit 1
} finally {
    # Ensure certificate object is disposed
    if ($pfx) {
        $pfx.Dispose()
    }
}

Write-Host "`nNext steps:" -ForegroundColor Yellow
Write-Host "1. Verify the extracted files are correct"
Write-Host "2. Update nginx configuration with the certificate paths"
Write-Host "3. Test the nginx configuration: nginx -t"
Write-Host "4. Restart nginx service"