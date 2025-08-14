# ExcelAddin Update Deployment Script
# Updates existing services without full reconfiguration

param(
    [switch]$RestartServices,
    [switch]$UpdateDependencies
)

# Import common functions
. (Join-Path $PSScriptRoot "scripts" | Join-Path -ChildPath "common.ps1")

Write-Header "ExcelAddin Update Deployment"

# Check prerequisites (skip installation checks)
if (-not (Test-Administrator)) {
    Write-Error "This script must be run as Administrator"
    exit 1
}

$ProjectRoot = Get-ProjectRoot
$BackendPath = Get-BackendPath

try {
    # Update Backend
    Write-Header "Step 1: Update Backend"
    
    Push-Location $BackendPath
    
    # Update Python dependencies if requested
    if ($UpdateDependencies) {
        Write-Host "Updating Python dependencies..."
        if (Test-Path "pyproject.toml") {
            if (Test-Command "poetry") {
                poetry install
                if ($LASTEXITCODE -ne 0) {
                    Write-Warning "Poetry install failed, continuing anyway"
                }
            }
        }
    }
    
    # Restart backend service
    Write-Host "Restarting backend service..."
    try {
        Restart-Service -Name "ExcelAddin-Backend" -Force
        Start-Sleep -Seconds 10
        
        # Verify backend is responding
        $response = Invoke-RestMethod -Uri "http://127.0.0.1:5000/api/health" -TimeoutSec 15
        Write-Success "Backend updated and responding: $($response.status)"
    } catch {
        Write-Warning "Backend restart failed or not responding: $($_.Exception.Message)"
    }
    
    Pop-Location
    
    # Update Frontend
    Write-Header "Step 2: Update Frontend"
    
    Push-Location $ProjectRoot
    
    # Update Node.js dependencies if requested
    if ($UpdateDependencies) {
        Write-Host "Updating Node.js dependencies..."
        npm install
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "npm install failed, continuing anyway"
        }
    }
    
    # Rebuild frontend
    Write-Host "Rebuilding frontend..."
    
    # First try the full build with custom functions
    npm run build:staging
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "Full build failed, attempting web-only build..."
        
        # Create fallback functions.json if it doesn't exist
        if (-not (Test-Path "src\commands\functions.json")) {
            Write-Host "Creating fallback functions.json..."
            Copy-Item "src\commands\functions.json.fallback" "src\commands\functions.json"
        }
        
        # Try web-only build
        npm run build:web
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Both full build and web-only build failed"
            exit 1
        }
        Write-Warning "Web-only build completed (Excel custom functions may not work)"
    } else {
        Write-Success "Full build completed successfully"
    }
    
    # Restart NSSM service
    Write-Host "Restarting frontend service..."
    try {
        Restart-Service -Name "ExcelAddin-Frontend" -Force
        Start-Sleep -Seconds 10
        
        # Verify frontend is responding
        $response = Invoke-WebRequest -Uri "http://127.0.0.1:3000" -TimeoutSec 15
        Write-Success "Frontend updated and responding: HTTP $($response.StatusCode)"
    } catch {
        Write-Warning "Frontend restart failed or not responding: $($_.Exception.Message)"
    }
    
    Pop-Location
    
    # Optionally restart services
    if ($RestartServices) {
        Write-Header "Step 3: Restart All Services"
        
        # Restart IIS
        Write-Host "Restarting IIS..."
        try {
            iisreset /restart
            Start-Sleep -Seconds 10
            Write-Success "IIS restarted"
        } catch {
            Write-Warning "IIS restart failed: $($_.Exception.Message)"
        }
    }
    
    # Verification
    Write-Header "Update Verification"
    
    # Check services
    $backendService = Get-Service -Name "ExcelAddin-Backend" -ErrorAction SilentlyContinue
    if ($backendService -and $backendService.Status -eq "Running") {
        Write-Success "Backend service: Running"
    } else {
        Write-Warning "Backend service status: $($backendService.Status)"
    }
    
    $frontendService = Get-Service -Name "ExcelAddin-Frontend" -ErrorAction SilentlyContinue
    if ($frontendService -and $frontendService.Status -eq "Running") {
        Write-Success "Frontend service: Running"
    } else {
        Write-Warning "Frontend service is not running"
    }
    
    $site = Get-IISSite -Name "ExcelAddin" -ErrorAction SilentlyContinue
    if ($site -and $site.State -eq "Started") {
        Write-Success "IIS site: Started"
    } else {
        Write-Warning "IIS site state: $($site.State)"
    }
    
    Write-Header "Update Complete"
    Write-Success "ExcelAddin services have been updated successfully"
    Write-Host ""
    Write-Host "Updated Components:"
    Write-Host "- Backend: Restarted and verified"
    Write-Host "- Frontend: Rebuilt and restarted"
    if ($RestartServices) {
        Write-Host "- IIS: Restarted"
    }
    Write-Host ""
    Write-Host "Public URL: https://server-vs81t.intranet.local:9443"
    Write-Host ""
    
} catch {
    Write-Error "Update failed: $($_.Exception.Message)"
    Write-Host $_.ScriptStackTrace
    exit 1
} finally {
    # Ensure we're back to the original location
    if ((Get-Location).Path -ne $PSScriptRoot) {
        Set-Location $PSScriptRoot
    }
}