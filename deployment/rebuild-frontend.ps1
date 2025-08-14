# Simple Frontend Rebuild Script
param(
    [switch]$Force
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir
$DistPath = Join-Path $ProjectRoot "dist"

Write-Host "===================================================="
Write-Host "Frontend Rebuild Utility"
Write-Host "===================================================="
Write-Host "Project Root: $ProjectRoot"
Write-Host "Dist Path: $DistPath"
Write-Host ""

# Navigate to project root
Set-Location $ProjectRoot

# Clean previous build
if (Test-Path $DistPath) {
    Write-Host "Removing previous build..."
    Remove-Item $DistPath -Recurse -Force
}

# Check if node_modules exists
if (!(Test-Path "node_modules")) {
    Write-Host "Installing dependencies..."
    npm install
    if ($LASTEXITCODE -ne 0) {
        Write-Error "npm install failed"
        exit 1
    }
    Write-Host "Dependencies installed successfully"
}

# Build frontend
Write-Host "Building frontend..."
npm run build:staging
if ($LASTEXITCODE -ne 0) {
    Write-Error "Frontend build failed"
    exit 1
}

# Verify build output
if (!(Test-Path $DistPath)) {
    Write-Error "Build completed but dist directory was not created"
    exit 1
}

$distFiles = Get-ChildItem $DistPath -ErrorAction SilentlyContinue
if (!$distFiles) {
    Write-Error "Build completed but dist directory is empty"
    exit 1
}

# Check for essential files
$indexPath = Join-Path $DistPath "index.html"
if (!(Test-Path $indexPath)) {
    Write-Error "index.html not found in build output"
    exit 1
}

Write-Host ""
Write-Host "✅ Build completed successfully!"
Write-Host "Dist directory contains $($distFiles.Count) files"

# List key files
Write-Host ""
Write-Host "Key files created:"
$keyFiles = @("index.html", "taskpane.html", "commands.html", "manifest.xml")
foreach ($file in $keyFiles) {
    $filePath = Join-Path $DistPath $file
    if (Test-Path $filePath) {
        $size = (Get-Item $filePath).Length
        Write-Host "✅ $file ($size bytes)"
    } else {
        Write-Host "❌ $file (missing)"
    }
}

Write-Host ""
Write-Host "===================================================="
Write-Host "Rebuild completed successfully!"
Write-Host "===================================================="
Write-Host ""
Write-Host "Next steps:"
Write-Host "1. Restart frontend service: Restart-Service ExcelAddin-Frontend"
Write-Host "2. Test the service: Invoke-WebRequest http://127.0.0.1:3000"
Write-Host "3. Check logs if needed: C:\Logs\ExcelAddin\frontend-*.log"
Write-Host "===================================================="