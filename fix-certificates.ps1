# Excel Add-in Certificate Fix Script
# Automatically diagnoses and fixes certificate issues for local development
# 
# Usage: .\fix-certificates.ps1

Write-Host "🔧 Excel Add-in Certificate Diagnostic Tool" -ForegroundColor Cyan
Write-Host ""

function Test-CertificateFiles {
    $certDir = Join-Path $env:USERPROFILE ".office-addin-dev-certs"
    $certFile = Join-Path $certDir "localhost.crt"
    $keyFile = Join-Path $certDir "localhost.key"
    
    Write-Host "📁 Checking certificate files..." -ForegroundColor Yellow
    Write-Host "   Directory: $certDir"
    
    $dirExists = Test-Path $certDir
    $certExists = Test-Path $certFile
    $keyExists = Test-Path $keyFile
    
    Write-Host "   Directory exists: $(if($dirExists){'✅'}else{'❌'})"
    Write-Host "   Certificate exists: $(if($certExists){'✅'}else{'❌'})"
    Write-Host "   Private key exists: $(if($keyExists){'✅'}else{'❌'})"
    Write-Host ""
    
    return $dirExists -and $certExists -and $keyExists
}

function Invoke-CertCommand {
    param(
        [string]$Command,
        [string]$Description
    )
    
    Write-Host "⏳ $Description..." -ForegroundColor Yellow
    
    try {
        $result = Invoke-Expression $Command 2>&1
        $success = $LASTEXITCODE -eq 0
        
        if ($success) {
            Write-Host "✅ $Description completed" -ForegroundColor Green
        } else {
            Write-Host "❌ $Description failed: $result" -ForegroundColor Red
            return $null
        }
        
        Write-Host ""
        return $result
    }
    catch {
        Write-Host "❌ $Description failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        return $null
    }
}

Write-Host "Starting certificate diagnosis..." -ForegroundColor Green
Write-Host ""

# Step 1: Check if certificate files exist
$filesExist = Test-CertificateFiles

# Step 2: Verify certificate status
Write-Host "🔍 Verifying certificate installation status..." -ForegroundColor Yellow
$verifyResult = Invoke-CertCommand "npm run cert:verify" "Certificate verification"

$needsInstall = (-not $verifyResult) -or ($verifyResult -like "*You need to install*") -or (-not $filesExist)

if ($needsInstall) {
    Write-Host "🚨 Certificates need to be installed or refreshed" -ForegroundColor Red
    Write-Host ""
    
    # Step 3: Uninstall old certificates if they exist
    if ($filesExist) {
        Write-Host "🧹 Removing old certificates..." -ForegroundColor Yellow
        $null = Invoke-CertCommand "npm run cert:uninstall" "Certificate uninstallation"
    }
    
    # Step 4: Install fresh certificates
    Write-Host "📜 Installing fresh certificates..." -ForegroundColor Yellow
    $installResult = Invoke-CertCommand "npm run cert:install" "Certificate installation"
    
    if ($installResult) {
        Write-Host "✅ Certificate installation completed!" -ForegroundColor Green
        
        # Step 5: Verify installation
        Write-Host "🔍 Verifying new certificate installation..." -ForegroundColor Yellow
        $null = Invoke-CertCommand "npm run cert:verify" "Final verification"
        
        Write-Host ""
        Write-Host "🎉 Certificate fix completed!" -ForegroundColor Green
        Write-Host ""
        Write-Host "📋 Next steps:" -ForegroundColor Cyan
        Write-Host "   1. Restart Excel completely"
        Write-Host "   2. Run: npm run dev"
        Write-Host "   3. Load your add-in in Excel"
        Write-Host ""
        Write-Host "If you still see certificate errors, please check the Certificate Guide:" -ForegroundColor Yellow
        Write-Host "   📖 CERTIFICATE_GUIDE.md"
        Write-Host ""
    } else {
        Write-Host ""
        Write-Host "❌ Certificate installation failed!" -ForegroundColor Red
        Write-Host ""
        Write-Host "🔧 Manual steps to try:" -ForegroundColor Yellow
        Write-Host "   1. Run this script as Administrator:"
        Write-Host "      Right-click PowerShell → Run as Administrator"
        Write-Host "   2. cd to your project directory"
        Write-Host "   3. npm run cert:install"
        Write-Host ""
        Write-Host "For more help, see: CERTIFICATE_GUIDE.md" -ForegroundColor Cyan
        Write-Host ""
    }
} else {
    Write-Host "✅ Certificates are already properly installed!" -ForegroundColor Green
    Write-Host ""
    Write-Host "🤔 If you're still seeing certificate errors in Excel:" -ForegroundColor Yellow
    Write-Host "   1. Restart Excel completely"
    Write-Host "   2. Clear browser cache (Ctrl+Shift+Delete)"
    Write-Host "   3. Check Windows Certificate Store: certlm.msc"
    Write-Host ""
    Write-Host "For advanced troubleshooting, see: CERTIFICATE_GUIDE.md" -ForegroundColor Cyan
    Write-Host ""
}