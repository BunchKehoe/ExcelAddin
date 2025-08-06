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