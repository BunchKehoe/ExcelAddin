# PowerShell script to handle encrypted SSL private keys
# This script converts encrypted private keys to unencrypted format for nginx
# Usage: .\handle-encrypted-key.ps1 -KeyPath "C:\Cert\server.key" -OutputPath "C:\Cert\server-unencrypted.key"

param(
    [Parameter(Mandatory=$true)]
    [string]$KeyPath,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath,
    
    [Parameter(Mandatory=$false)]
    [SecureString]$Password,
    
    [Parameter(Mandatory=$false)]
    [switch]$TestOnly
)

# Check if OpenSSL is available
function Test-OpenSSL {
    try {
        $version = & openssl version 2>$null
        return $true
    } catch {
        return $false
    }
}

# Check if key is encrypted
function Test-KeyEncrypted {
    param([string]$keyPath)
    
    if (-not (Test-Path $keyPath)) {
        return $false
    }
    
    $keyContent = Get-Content $keyPath -Raw
    return $keyContent -match "ENCRYPTED|Proc-Type.*ENCRYPTED"
}

Write-Host "SSL Private Key Handler" -ForegroundColor Cyan
Write-Host "=" * 30

if (-not (Test-Path $KeyPath)) {
    Write-Error "Key file not found: $KeyPath"
    exit 1
}

$isEncrypted = Test-KeyEncrypted -keyPath $KeyPath
Write-Host "Key file: $KeyPath"
Write-Host "Encrypted: $isEncrypted"

if ($TestOnly) {
    if ($isEncrypted) {
        Write-Host "The private key is encrypted and will cause nginx to prompt for a password." -ForegroundColor Yellow
        Write-Host "Use this script without -TestOnly to convert it to an unencrypted format." -ForegroundColor Yellow
        Write-Host "`nAlternatively, you can use OpenSSL directly:" -ForegroundColor Cyan
        Write-Host "openssl rsa -in `"$KeyPath`" -out `"$KeyPath.unencrypted`"" -ForegroundColor Gray
    } else {
        Write-Host "The private key is not encrypted and should work with nginx." -ForegroundColor Green
    }
    exit 0
}

if (-not $isEncrypted) {
    Write-Host "Private key is already unencrypted. No action needed." -ForegroundColor Green
    exit 0
}

# Set default output path
if (-not $OutputPath) {
    $dir = Split-Path $KeyPath
    $name = [System.IO.Path]::GetFileNameWithoutExtension($KeyPath)
    $OutputPath = Join-Path $dir "$name-unencrypted.key"
}

Write-Host "`nConverting encrypted key to unencrypted format..." -ForegroundColor Yellow

# Try using OpenSSL if available
if (Test-OpenSSL) {
    Write-Host "Using OpenSSL for key conversion..." -ForegroundColor Green
    
    try {
        if ($Password) {
            # Convert secure string to plain text for OpenSSL
            $plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
            )
            $env:OPENSSL_PASSWORD = $plainPassword
            & openssl rsa -in $KeyPath -out $OutputPath -passin env:OPENSSL_PASSWORD
            Remove-Item Env:\OPENSSL_PASSWORD -ErrorAction SilentlyContinue
        } else {
            & openssl rsa -in $KeyPath -out $OutputPath
        }
        
        if ($LASTEXITCODE -eq 0 -and (Test-Path $OutputPath)) {
            Write-Host "✓ Key converted successfully to: $OutputPath" -ForegroundColor Green
            
            # Verify the unencrypted key
            $unencryptedIsEncrypted = Test-KeyEncrypted -keyPath $OutputPath
            if (-not $unencryptedIsEncrypted) {
                Write-Host "✓ Verified: Output key is unencrypted" -ForegroundColor Green
            } else {
                Write-Warning "Warning: Output key may still be encrypted"
            }
            
            # Set secure permissions on the new key file
            Write-Host "Setting secure permissions on unencrypted key..." -ForegroundColor Yellow
            icacls $OutputPath /inheritance:r /grant:r "Administrators:F" /grant:r "SYSTEM:F" | Out-Null
            Write-Host "✓ Secure permissions set" -ForegroundColor Green
            
        } else {
            Write-Error "OpenSSL conversion failed"
            exit 1
        }
    } catch {
        Write-Error "Error using OpenSSL: $_"
        exit 1
    }
    
} else {
    Write-Warning "OpenSSL not found. Attempting PowerShell-based conversion..."
    
    # PowerShell-based approach (limited support)
    try {
        # Get password if not provided
        if (-not $Password) {
            $Password = Read-Host "Enter private key password" -AsSecureString
        }
        
        # Convert secure string to plain text
        $plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
        )
        
        # Read the encrypted key content
        $keyContent = Get-Content $KeyPath -Raw
        
        # This is a basic approach - may not work for all key formats
        Write-Host "PowerShell-based conversion has limited support." -ForegroundColor Yellow
        Write-Host "For best results, install OpenSSL for Windows:" -ForegroundColor Yellow
        Write-Host "1. Download from: https://slproweb.com/products/Win32OpenSSL.html" -ForegroundColor Gray
        Write-Host "2. Or use: winget install ShiningLight.OpenSSL" -ForegroundColor Gray
        
        # Try using .NET crypto classes (basic support)
        try {
            $keyBytes = [System.Text.Encoding]::UTF8.GetBytes($keyContent)
            # This approach is limited and may not work for all key formats
            Write-Warning "PowerShell conversion is not fully supported for encrypted keys."
            Write-Host "Please install OpenSSL for reliable key conversion." -ForegroundColor Yellow
        } catch {
            Write-Error "PowerShell-based conversion failed: $_"
            exit 1
        }
        
    } catch {
        Write-Error "Failed to convert key: $_"
        exit 1
    }
}

# Provide next steps
Write-Host "`nNext Steps:" -ForegroundColor Cyan
Write-Host "1. Update nginx configuration to use the unencrypted key:"
Write-Host "   ssl_certificate_key $($OutputPath.Replace('\', '/'));"
Write-Host "2. Test nginx configuration: nginx -t"
Write-Host "3. Restart nginx service"
Write-Host "4. Verify nginx starts without password prompt"

Write-Host "`nSecurity Note:" -ForegroundColor Yellow
Write-Host "• The unencrypted key file contains sensitive data"
Write-Host "• Ensure proper file permissions (administrators only)"
Write-Host "• Consider using a hardware security module (HSM) for production"
Write-Host "• Regularly rotate SSL certificates and keys"

# Create nginx configuration snippet
$nginxSnippet = @"
# SSL Configuration with unencrypted key
ssl_certificate C:/Cert/server.crt;
ssl_certificate_key $($OutputPath.Replace('\', '/'));
ssl_trusted_certificate C:/Cert/cacert.pem;
"@

$snippetPath = Join-Path (Split-Path $OutputPath) "nginx-ssl-snippet.conf"
$nginxSnippet | Set-Content -Path $snippetPath
Write-Host "`nnginx configuration snippet saved to: $snippetPath" -ForegroundColor Green