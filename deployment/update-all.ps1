# Simple Excel Add-in Update Script
param(
    [switch]$UpdateDependencies
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir

Write-Host "===================================================="
Write-Host "Excel Add-in Update Deployment"
Write-Host "===================================================="
Write-Host "Project Root: $ProjectRoot"
Write-Host ""

try {
    # Update Backend
    Write-Host "Step 1: Update Backend"
    Write-Host "----------------------------------------------------"
    
    # Navigate to backend directory
    $BackendPath = Join-Path $ProjectRoot "backend"
    Set-Location $BackendPath
    
    if ($UpdateDependencies) {
        Write-Host "Updating Python dependencies..."
        pip install -r requirements.txt --upgrade
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "Python dependency update failed, continuing..."
        }
    }
    
    # Restart backend service
    Write-Host "Restarting backend service..."
    Restart-Service -Name "ExcelAddin-Backend" -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 10
    
    # Check backend status
    $backendService = Get-Service -Name "ExcelAddin-Backend" -ErrorAction SilentlyContinue
    if ($backendService -and $backendService.Status -eq "Running") {
        Write-Host "✅ Backend service restarted successfully"
    } else {
        Write-Host "⚠️  Backend service restart issues"
    }
    
    # Update Frontend
    Write-Host ""
    Write-Host "Step 2: Update Frontend"
    Write-Host "----------------------------------------------------"
    
    # Navigate to project root
    Set-Location $ProjectRoot
    
    if ($UpdateDependencies) {
        Write-Host "Updating Node.js dependencies..."
        npm install
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "npm install failed, continuing..."
        }
    }
    
    # Clean and rebuild frontend
    Write-Host "Cleaning previous build..."
    if (Test-Path "dist") {
        Remove-Item "dist" -Recurse -Force
    }
    
    Write-Host "Rebuilding frontend..."
    npm run build:staging
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Frontend build failed"
        exit 1
    }
    
    # Restart frontend service
    Write-Host "Restarting frontend service..."
    Restart-Service -Name "ExcelAddin-Frontend" -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 10
    
    # Check frontend status
    $frontendService = Get-Service -Name "ExcelAddin-Frontend" -ErrorAction SilentlyContinue
    if ($frontendService -and $frontendService.Status -eq "Running") {
        Write-Host "✅ Frontend service restarted successfully"
    } else {
        Write-Host "⚠️  Frontend service restart issues"
    }
    
    # Test services
    Write-Host ""
    Write-Host "Step 3: Verify Services"
    Write-Host "----------------------------------------------------"
    
    # Test backend
    try {
        $healthCheck = Invoke-RestMethod -Uri "http://127.0.0.1:5000/api/health" -TimeoutSec 10
        Write-Host "✅ Backend health check: OK"
    } catch {
        Write-Host "❌ Backend health check failed"
    }
    
    # Test frontend
    try {
        $response = Invoke-WebRequest -Uri "http://127.0.0.1:3000" -UseBasicParsing -TimeoutSec 10
        if ($response.StatusCode -eq 200) {
            Write-Host "✅ Frontend health check: OK"
        } else {
            Write-Host "⚠️  Frontend returned HTTP $($response.StatusCode)"
        }
    } catch {
        Write-Host "❌ Frontend health check failed"
    }
    
    Write-Host ""
    Write-Host "===================================================="
    Write-Host "Update Completed Successfully!"
    Write-Host "===================================================="
    Write-Host "Services:"
    Write-Host "- Backend:  ExcelAddin-Backend  (http://127.0.0.1:5000)"
    Write-Host "- Frontend: ExcelAddin-Frontend (http://127.0.0.1:3000)"
    Write-Host ""
    
} catch {
    Write-Error "Update failed: $($_.Exception.Message)"
    exit 1
} finally {
    # Return to script directory
    Set-Location $ScriptDir
}