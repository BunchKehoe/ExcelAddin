# Simple Excel Add-in Complete Deployment Script
param(
    [switch]$ConfigureIIS
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "===================================================="
Write-Host "Excel Add-in Complete Deployment"
Write-Host "===================================================="
Write-Host ""

try {
    # Deploy Backend
    Write-Host "Step 1: Deploy Backend Service"
    Write-Host "----------------------------------------------------"
    & "$ScriptDir\deploy-backend.ps1"
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Backend deployment failed"
        exit 1
    }
    Write-Host "Backend deployment completed successfully"
    Write-Host ""
    
    # Deploy Frontend
    Write-Host "Step 2: Deploy Frontend Service"
    Write-Host "----------------------------------------------------"
    $frontendArgs = @()
    if ($ConfigureIIS) { $frontendArgs += "-ConfigureIIS" }
    
    & "$ScriptDir\deploy-frontend.ps1" @frontendArgs
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Frontend deployment failed"
        exit 1
    }
    Write-Host "Frontend deployment completed successfully"
    Write-Host ""
    
    # Verification
    Write-Host "Step 3: Verify Deployment"
    Write-Host "----------------------------------------------------"
    
    # Wait for services to stabilize
    Start-Sleep -Seconds 10
    
    # Check backend
    $backendService = Get-Service -Name "ExcelAddin-Backend" -ErrorAction SilentlyContinue
    if ($backendService -and $backendService.Status -eq "Running") {
        Write-Host "✅ Backend service: Running"
        try {
            $healthCheck = Invoke-RestMethod -Uri "http://127.0.0.1:5000/api/health" -TimeoutSec 10
            Write-Host "✅ Backend health check: OK"
        } catch {
            Write-Host "⚠️  Backend health check failed"
        }
    } else {
        Write-Host "❌ Backend service is not running"
    }
    
    # Check frontend
    $frontendService = Get-Service -Name "ExcelAddin-Frontend" -ErrorAction SilentlyContinue
    if ($frontendService -and $frontendService.Status -eq "Running") {
        Write-Host "✅ Frontend service: Running"
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
    } else {
        Write-Host "❌ Frontend service is not running"
    }
    
    Write-Host ""
    Write-Host "===================================================="
    Write-Host "Deployment Completed Successfully!"
    Write-Host "===================================================="
    Write-Host "Services:"
    Write-Host "- Backend:  ExcelAddin-Backend  (http://127.0.0.1:5000)"
    Write-Host "- Frontend: ExcelAddin-Frontend (http://127.0.0.1:3000)"
    if ($ConfigureIIS) {
        Write-Host "- IIS Site: ExcelAddin"
    }
    Write-Host ""
    Write-Host "To manage services:"
    Write-Host "- View status: Get-Service ExcelAddin-*"
    Write-Host "- Stop service: Stop-Service ExcelAddin-Frontend"
    Write-Host "- Start service: Start-Service ExcelAddin-Frontend"
    Write-Host "===================================================="
    
} catch {
    Write-Error "Deployment failed: $($_.Exception.Message)"
    exit 1
}