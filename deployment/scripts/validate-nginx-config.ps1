# nginx Configuration Validation Script for Windows
# Usage: .\validate-nginx-config.ps1

param(
    [Parameter(Mandatory=$false)]
    [string]$NginxPath = "C:\nginx",
    
    [Parameter(Mandatory=$false)]
    [string]$CertPath = "C:\Cert"
)

Write-Host "Excel Add-in nginx Configuration Validator" -ForegroundColor Cyan
Write-Host "=" * 50

$errors = @()
$warnings = @()

# Check if nginx is installed
Write-Host "`nChecking nginx installation..." -ForegroundColor Yellow
if (-not (Test-Path "$NginxPath\nginx.exe")) {
    $errors += "nginx.exe not found at $NginxPath"
} else {
    Write-Host "✓ nginx found at $NginxPath" -ForegroundColor Green
}

# Check nginx configuration files
Write-Host "`nChecking nginx configuration files..." -ForegroundColor Yellow
if (-not (Test-Path "$NginxPath\conf\nginx.conf")) {
    $errors += "nginx.conf not found at $NginxPath\conf\nginx.conf"
} else {
    Write-Host "✓ nginx.conf found" -ForegroundColor Green
}

if (-not (Test-Path "$NginxPath\conf\excel-addin.conf")) {
    $warnings += "excel-addin.conf not found at $NginxPath\conf\excel-addin.conf - you may need to copy it"
} else {
    Write-Host "✓ excel-addin.conf found" -ForegroundColor Green
}

# Check certificate files
Write-Host "`nChecking SSL certificates..." -ForegroundColor Yellow
$certFiles = @(
    @{Name="Server Certificate"; Path="$CertPath\server.crt"; Required=$true},
    @{Name="Private Key"; Path="$CertPath\server.key"; Required=$true},
    @{Name="Root CA"; Path="$CertPath\cacert.pem"; Required=$false}
)

foreach ($cert in $certFiles) {
    if (Test-Path $cert.Path) {
        Write-Host "✓ $($cert.Name) found at $($cert.Path)" -ForegroundColor Green
        
        # Check certificate details if it's a .crt file
        if ($cert.Path -like "*.crt") {
            try {
                $certObj = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($cert.Path)
                Write-Host "  Subject: $($certObj.Subject)" -ForegroundColor Gray
                Write-Host "  Expires: $($certObj.NotAfter)" -ForegroundColor Gray
                
                if ($certObj.NotAfter -lt (Get-Date).AddDays(30)) {
                    $warnings += "$($cert.Name) expires in less than 30 days: $($certObj.NotAfter)"
                }
                
                # Check if certificate matches the domain
                $serverName = "server01.intranet.local"
                if ($certObj.Subject -like "*$serverName*" -or $certObj.DnsNameList -contains $serverName) {
                    Write-Host "  ✓ Certificate matches domain $serverName" -ForegroundColor Green
                } else {
                    $warnings += "Certificate subject may not match domain $serverName"
                }
                
                $certObj.Dispose()
            } catch {
                $warnings += "Could not read certificate details from $($cert.Path): $_"
            }
        }
    } elseif ($cert.Required) {
        $errors += "$($cert.Name) not found at $($cert.Path)"
    } else {
        $warnings += "$($cert.Name) not found at $($cert.Path) - optional but recommended"
    }
}

# Check if alternative .pfx file exists
if (-not (Test-Path "$CertPath\server.crt") -and (Test-Path "$CertPath\server.pfx")) {
    $warnings += "Found server.pfx but no server.crt. Use extract-pfx.ps1 to extract certificate and key files."
}

# Check for encrypted private keys
if (Test-Path "$CertPath\server.key") {
    try {
        $keyContent = Get-Content "$CertPath\server.key" -Raw
        if ($keyContent -match "ENCRYPTED|Proc-Type.*ENCRYPTED") {
            $errors += "Private key is encrypted and will cause nginx to prompt for password. Use handle-encrypted-key.ps1 to convert to unencrypted format."
        } else {
            Write-Host "✓ Private key is not encrypted" -ForegroundColor Green
        }
    } catch {
        $warnings += "Could not read private key file to check encryption status"
    }
}

# Check directory permissions
Write-Host "`nChecking directory structure..." -ForegroundColor Yellow
$directories = @(
    "C:\inetpub\wwwroot\ExcelAddin\dist",
    "C:\Logs\nginx",
    "C:\Logs\ExcelAddin"
)

foreach ($dir in $directories) {
    if (Test-Path $dir) {
        Write-Host "✓ Directory exists: $dir" -ForegroundColor Green
    } else {
        $warnings += "Directory does not exist: $dir - will be created during deployment"
    }
}

# Test nginx configuration syntax
Write-Host "`nTesting nginx configuration syntax..." -ForegroundColor Yellow
if (Test-Path "$NginxPath\nginx.exe") {
    try {
        $result = & "$NginxPath\nginx.exe" -t -c "$NginxPath\conf\nginx.conf" 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ nginx configuration syntax is valid" -ForegroundColor Green
        } else {
            $errors += "nginx configuration syntax error: $result"
        }
        
        # Check for common nginx warnings in the configuration
        $configContent = Get-Content "$NginxPath\conf\excel-addin.conf" -Raw -ErrorAction SilentlyContinue
        if ($configContent) {
            if ($configContent -match "listen.*http2" -and $configContent -notmatch "http2 on") {
                $warnings += "Configuration may still use deprecated 'listen ... http2' syntax. Update to use 'http2 on;' directive."
            }
            
            if ($configContent -match "ssl_stapling on" -and $configContent -notmatch "#.*ssl_stapling") {
                $warnings += "SSL stapling is enabled but may cause warnings with company certificates. Consider disabling if certificate lacks OCSP responder."
            }
        }
        
        # Check main nginx.conf for Windows compatibility
        $mainConfigContent = Get-Content "$NginxPath\conf\nginx.conf" -Raw -ErrorAction SilentlyContinue
        if ($mainConfigContent) {
            if ($mainConfigContent -match "daemon on" -or $mainConfigContent -notmatch "daemon off") {
                $warnings += "nginx.conf should have 'daemon off;' for Windows service compatibility with NSSM"
            }
            
            if ($mainConfigContent -match "worker_processes\s+auto") {
                $warnings += "Consider using 'worker_processes 1;' instead of 'auto' for Windows stability"
            }
            
            if ($mainConfigContent -notmatch "use select") {
                $warnings += "Consider using 'use select;' in events block for Windows compatibility"
            }
        }
    } catch {
        $warnings += "Could not test nginx configuration: $_"
    }
} else {
    $warnings += "Cannot test nginx configuration - nginx.exe not found"
}

# Check Windows service status
Write-Host "`nChecking Windows services..." -ForegroundColor Yellow
try {
    $nginxService = Get-Service -Name "nginx" -ErrorAction SilentlyContinue
    if ($nginxService) {
        Write-Host "✓ nginx Windows service exists - Status: $($nginxService.Status)" -ForegroundColor Green
    } else {
        $warnings += "nginx Windows service not found - will be created during deployment"
    }
} catch {
    $warnings += "Could not check nginx service status"
}

try {
    $backendService = Get-Service -Name "ExcelAddinBackend" -ErrorAction SilentlyContinue
    if ($backendService) {
        Write-Host "✓ ExcelAddinBackend Windows service exists - Status: $($backendService.Status)" -ForegroundColor Green
    } else {
        $warnings += "ExcelAddinBackend Windows service not found - will be created during deployment"
    }
} catch {
    $warnings += "Could not check ExcelAddinBackend service status"
}

# Summary
Write-Host "`n" + "=" * 50 -ForegroundColor Cyan
Write-Host "VALIDATION SUMMARY" -ForegroundColor Cyan
Write-Host "=" * 50 -ForegroundColor Cyan

if ($errors.Count -eq 0) {
    Write-Host "✓ No critical errors found" -ForegroundColor Green
} else {
    Write-Host "✗ $($errors.Count) critical error(s) found:" -ForegroundColor Red
    foreach ($error in $errors) {
        Write-Host "  • $error" -ForegroundColor Red
    }
}

if ($warnings.Count -eq 0) {
    Write-Host "✓ No warnings" -ForegroundColor Green
} else {
    Write-Host "⚠ $($warnings.Count) warning(s):" -ForegroundColor Yellow
    foreach ($warning in $warnings) {
        Write-Host "  • $warning" -ForegroundColor Yellow
    }
}

# Recommendations
Write-Host "`nRECOMMENDATIONS:" -ForegroundColor Cyan
if ($errors.Count -gt 0) {
    Write-Host "1. Fix all critical errors before proceeding with deployment" -ForegroundColor Red
    Write-Host "2. Review the CERTIFICATE_SETUP.md guide for certificate configuration" -ForegroundColor Yellow
    Write-Host "3. Ensure nginx and required files are in the correct locations" -ForegroundColor Yellow
} else {
    Write-Host "1. Configuration appears ready for deployment" -ForegroundColor Green
    Write-Host "2. Address any warnings if applicable" -ForegroundColor Yellow
    Write-Host "3. Proceed with the deployment using deploy-windows.ps1" -ForegroundColor Green
}

Write-Host "`nCOMMON NGINX WINDOWS ISSUES AND FIXES:" -ForegroundColor Cyan
Write-Host "• Process closes immediately: Use nginx.conf.windows.template for Windows optimizations"
Write-Host "• Password prompt on startup: Use handle-encrypted-key.ps1 to convert encrypted keys"
Write-Host "• Service management: Use setup-nginx-service.ps1 to configure nginx with NSSM"
Write-Host "• HTTP/2 deprecation warning: Fixed in current config using 'http2 on;' directive"
Write-Host "• SSL stapling warning: Disabled for company certificates without OCSP responder"
Write-Host "• Master process alert: Fixed with Windows-specific settings and service configuration"

Write-Host "`nFor immediate solutions to critical issues, see:" -ForegroundColor Cyan
Write-Host "• NGINX_WINDOWS_QUICK_FIX.md - Immediate solutions for common problems"

Write-Host "`nFor detailed setup instructions, see:" -ForegroundColor Cyan
Write-Host "• CERTIFICATE_SETUP.md - SSL certificate configuration"
Write-Host "• WINDOWS_DEPLOYMENT.md - Complete deployment guide"
Write-Host "• EXCELLENCE_PATH_DEPLOYMENT.md - Subpath deployment guide"

# Exit with error code if critical errors found
if ($errors.Count -gt 0) {
    exit 1
} else {
    exit 0
}