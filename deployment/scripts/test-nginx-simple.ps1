# Test nginx configuration with minimal setup
# Usage: .\test-nginx-simple.ps1
param()

Write-Host "🔧 Testing Simple nginx Configuration" -ForegroundColor Cyan
Write-Host ""

# Check nginx installation
if (-not (Test-Path "C:\nginx\nginx.exe")) {
    Write-Host "❌ nginx not found at C:\nginx\nginx.exe" -ForegroundColor Red
    Write-Host "💡 Please install nginx to C:\nginx first" -ForegroundColor Yellow
    exit 1
}

Write-Host "✅ nginx executable found" -ForegroundColor Green

# Test configuration syntax
Write-Host ""
Write-Host "🔍 Testing nginx configuration syntax..." -ForegroundColor Yellow
$originalPath = Get-Location
try {
    Set-Location "C:\nginx"
    $configTest = & "C:\nginx\nginx.exe" -t 2>&1
    Set-Location $originalPath
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✅ nginx configuration syntax is valid" -ForegroundColor Green
    } else {
        Write-Host "❌ nginx configuration has syntax errors:" -ForegroundColor Red
        Write-Host $configTest -ForegroundColor White
        Write-Host ""
        Write-Host "💡 Check the configuration files:" -ForegroundColor Yellow
        Write-Host "   C:\nginx\conf\nginx.conf" -ForegroundColor White
        Write-Host "   C:\nginx\conf\excel-addin.conf" -ForegroundColor White
        exit 1
    }
} catch {
    Set-Location $originalPath
    Write-Host "❌ Failed to test nginx configuration" -ForegroundColor Red
    exit 1
}

# Check required directories
Write-Host ""
Write-Host "📁 Checking required directories..." -ForegroundColor Yellow

$requiredDirs = @(
    "C:\inetpub\wwwroot\ExcelAddin\dist",
    "C:\Cert"
)

foreach ($dir in $requiredDirs) {
    if (Test-Path $dir) {
        Write-Host "✅ Directory exists: $dir" -ForegroundColor Green
    } else {
        Write-Host "❌ Directory missing: $dir" -ForegroundColor Red
    }
}

# Check certificate files
Write-Host ""
Write-Host "🔐 Checking SSL certificates..." -ForegroundColor Yellow

$certFiles = @(
    "C:\Cert\server-vs81t.crt",
    "C:\Cert\server-vs81t.key"
)

foreach ($certFile in $certFiles) {
    if (Test-Path $certFile) {
        Write-Host "✅ Certificate file exists: $certFile" -ForegroundColor Green
    } else {
        Write-Host "❌ Certificate file missing: $certFile" -ForegroundColor Red
        Write-Host "💡 Make sure certificate files are named correctly" -ForegroundColor Yellow
    }
}

# Check frontend files
Write-Host ""
Write-Host "🌐 Checking frontend files..." -ForegroundColor Yellow

$frontendFiles = @(
    "C:\inetpub\wwwroot\ExcelAddin\dist\taskpane.html",
    "C:\inetpub\wwwroot\ExcelAddin\dist\commands.html"
)

foreach ($file in $frontendFiles) {
    if (Test-Path $file) {
        Write-Host "✅ Frontend file exists: $file" -ForegroundColor Green
    } else {
        Write-Host "❌ Frontend file missing: $file" -ForegroundColor Red
        Write-Host "💡 Run 'npm run build:staging' to build frontend files" -ForegroundColor Yellow
    }
}

# Check if port is available
Write-Host ""
Write-Host "🔌 Checking port availability..." -ForegroundColor Yellow
try {
    $portCheck = netstat -an | Select-String ":9443"
    if ($portCheck) {
        Write-Host "⚠️  Port 9443 is already in use:" -ForegroundColor Yellow
        Write-Host $portCheck -ForegroundColor White
        Write-Host "💡 Stop nginx if it's already running: Stop-Service nginx" -ForegroundColor Yellow
    } else {
        Write-Host "✅ Port 9443 is available" -ForegroundColor Green
    }
} catch {
    Write-Host "⚠️  Could not check port status" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "📋 Summary:" -ForegroundColor Cyan
Write-Host "   If all checks passed, you can start nginx with:" -ForegroundColor Yellow
Write-Host "   Start-Service nginx" -ForegroundColor White
Write-Host ""
Write-Host "   Test URLs after starting:" -ForegroundColor Yellow
Write-Host "   https://server-vs81t.intranet.local:9443/health" -ForegroundColor White
Write-Host "   https://server-vs81t.intranet.local:9443/excellence/taskpane.html" -ForegroundColor White
Write-Host ""