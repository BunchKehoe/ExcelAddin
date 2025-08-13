# Quick frontend rebuild script - ensures dist directory is properly built
# Useful for fixing 404 issues when dist directory is missing or empty

param(
    [switch]$Force
)

. (Join-Path $PSScriptRoot "scripts\common.ps1")

Write-Header "Frontend Rebuild Utility"

$FrontendPath = Get-FrontendPath
$DistPath = Join-Path $FrontendPath "dist"

Write-Host "Frontend Path: $FrontendPath"
Write-Host "Dist Path: $DistPath"
Write-Host ""

# Check current dist status
if (Test-Path $DistPath) {
    $distContents = Get-ChildItem $DistPath -ErrorAction SilentlyContinue
    if ($distContents) {
        Write-Host "Dist directory exists with $($distContents.Count) files" -ForegroundColor Green
        
        if (-not $Force) {
            Write-Host ""
            $rebuild = Read-Host "Dist directory already exists. Rebuild anyway? (y/n)"
            if ($rebuild -ne 'y' -and $rebuild -ne 'Y') {
                Write-Host "Skipping rebuild"
                exit 0
            }
        }
    } else {
        Write-Warning "Dist directory exists but is empty - rebuilding required"
    }
} else {
    Write-Warning "Dist directory does not exist - rebuild required"
}

# Navigate to frontend directory
Write-Host "Navigating to frontend directory..."
Set-Location $FrontendPath

# Check package.json
if (-not (Test-Path "package.json")) {
    Write-Error "package.json not found in $FrontendPath"
    exit 1
}

# Install dependencies if node_modules is missing
if (-not (Test-Path "node_modules")) {
    Write-Host "Installing dependencies..." -ForegroundColor Yellow
    npm install
    if ($LASTEXITCODE -ne 0) {
        Write-Error "npm install failed"
        exit 1
    }
    Write-Success "Dependencies installed"
}

# Build the frontend
Write-Host "Building frontend for staging..." -ForegroundColor Yellow
npm run build:staging
if ($LASTEXITCODE -ne 0) {
    Write-Error "Frontend build failed"
    exit 1
}

# Verify build output
if (Test-Path $DistPath) {
    $newDistContents = Get-ChildItem $DistPath -ErrorAction SilentlyContinue
    if ($newDistContents) {
        Write-Success "Build completed successfully!"
        Write-Host "Dist directory now contains $($newDistContents.Count) files:" -ForegroundColor Green
        $newDistContents | Select-Object -First 10 | ForEach-Object { 
            Write-Host "  $($_.Name) ($($_.Length) bytes)" 
        }
        if ($newDistContents.Count -gt 10) {
            Write-Host "  ... and $($newDistContents.Count - 10) more files"
        }
        
        # Check for index.html specifically
        $indexPath = Join-Path $DistPath "index.html"
        if (Test-Path $indexPath) {
            $indexSize = (Get-Item $indexPath).Length
            Write-Host ""
            Write-Host "index.html: $indexSize bytes" -ForegroundColor Green
        } else {
            Write-Warning "index.html not found in build output!"
        }
    } else {
        Write-Error "Build completed but dist directory is empty!"
        exit 1
    }
} else {
    Write-Error "Build completed but dist directory was not created!"
    exit 1
}

Write-Host ""
Write-Success "Frontend rebuild completed successfully!"
Write-Host ""
Write-Host "Next steps:"
Write-Host "1. Restart the frontend service: Restart-Service -Name 'ExcelAddin-Frontend'"
Write-Host "2. Test the frontend: Invoke-WebRequest http://127.0.0.1:3000"
Write-Host "3. Check service logs if issues persist: C:\Logs\ExcelAddin\frontend-*.log"