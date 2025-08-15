# ExcelAddin Complete Deployment Script
# Deploys both backend and frontend services for Windows Server 10

param(
    [ValidateSet("development", "staging", "production")]
    [string]$Environment = "staging",
    [switch]$SkipBuild,
    [switch]$SkipBackend,
    [switch]$SkipFrontend,
    [switch]$SkipIISProxy,
    [switch]$SkipTests,
    [switch]$Force,
    [switch]$Debug
)

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Green
Write-Host "  ExcelAddin Complete Deployment" -ForegroundColor Green
Write-Host "  Environment: $Environment" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

# Script locations
$ScriptDir = $PSScriptRoot
$BackendScript = Join-Path $ScriptDir "deploy-backend.ps1"
$FrontendScript = Join-Path $ScriptDir "deploy-frontend.ps1"
$IISProxyScript = Join-Path $ScriptDir "deploy-iis-proxy.ps1"
$DebugScript = Join-Path $ScriptDir "debug-integration.ps1"

# Verify scripts exist
$requiredScripts = @($BackendScript, $FrontendScript, $IISProxyScript, $DebugScript)
foreach ($script in $requiredScripts) {
    if (-not (Test-Path $script)) {
        Write-Error "Required script not found: $script"
    }
}

$startTime = Get-Date

try {
    # Deploy Backend
    if (-not $SkipBackend) {
        Write-Host "DEPLOYING BACKEND..." -ForegroundColor Cyan
        Write-Host ""
        
        $backendParams = @{
            Environment = $Environment
            Force = $Force
            Debug = $Debug
        }
        
        & $BackendScript @backendParams
        if ($LASTEXITCODE -ne 0) {
            throw "Backend deployment failed"
        }
        
        Write-Host "Backend deployment completed!" -ForegroundColor Green
        Write-Host ""
    }
    
    # Deploy Frontend
    if (-not $SkipFrontend) {
        Write-Host "DEPLOYING FRONTEND..." -ForegroundColor Cyan
        Write-Host ""
        
        $frontendParams = @{
            Environment = $Environment
            SkipBuild = $SkipBuild
            Force = $Force
            Debug = $Debug
        }
        
        & $FrontendScript @frontendParams
        if ($LASTEXITCODE -ne 0) {
            throw "Frontend deployment failed"
        }
        
        Write-Host "Frontend deployment completed!" -ForegroundColor Green
        Write-Host ""
    }
    
    # Deploy IIS Proxy
    if (-not $SkipIISProxy) {
        Write-Host "DEPLOYING IIS PROXY..." -ForegroundColor Cyan
        Write-Host ""
        
        $iisProxyParams = @{
            Force = $Force
            Debug = $Debug
        }
        
        & $IISProxyScript @iisProxyParams
        if ($LASTEXITCODE -ne 0) {
            throw "IIS Proxy deployment failed"
        }
        
        Write-Host "IIS Proxy deployment completed!" -ForegroundColor Green
        Write-Host ""
    }
    
    # Run Integration Tests
    if (-not $SkipTests) {
        Write-Host "RUNNING INTEGRATION TESTS..." -ForegroundColor Cyan
        Write-Host ""
        
        Start-Sleep -Seconds 5  # Allow services to fully start
        
        $debugParams = @{
            Detailed = $Debug
            FixIssues = $Force
        }
        
        & $DebugScript @debugParams
        Write-Host ""
    }
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "  DEPLOYMENT SUCCESSFUL!" -ForegroundColor Green
    Write-Host "  Duration: $($duration.ToString('mm\:ss'))" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    
    Write-Host "Service Status:" -ForegroundColor Cyan
    if (-not $SkipBackend) {
        $backendService = Get-Service "ExcelAddin-Backend" -ErrorAction SilentlyContinue
        if ($backendService) {
            Write-Host "  Backend: $($backendService.Status)" -ForegroundColor Green
        }
    }
    
    if (-not $SkipFrontend) {
        $frontendService = Get-Service "ExcelAddin Frontend" -ErrorAction SilentlyContinue  
        if ($frontendService) {
            Write-Host "  Frontend: $($frontendService.Status)" -ForegroundColor Green
        }
    }
    
    Write-Host ""
    Write-Host "Access URLs (through IIS proxy):" -ForegroundColor Cyan
    Write-Host "  Taskpane: https://server-vs81t.intranet.local:9443/excellence/taskpane.html"
    Write-Host "  API: https://server-vs81t.intranet.local:9443/api/health"
    Write-Host ""
    
    Write-Host "Local Testing URLs:" -ForegroundColor Cyan
    Write-Host "  Frontend Health: http://localhost:3000/health"
    Write-Host "  Backend Health: http://localhost:5000/api/health"
    Write-Host ""
    
    Write-Host "Troubleshooting:" -ForegroundColor Cyan
    Write-Host "  Debug: .\debug-integration.ps1 -Detailed"
    Write-Host "  Logs: Get-Content 'C:\Logs\ExcelAddin\*-stderr.log' -Tail 20"
    
} catch {
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "  DEPLOYMENT FAILED!" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Duration: $($duration.ToString('mm\:ss'))" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    
    Write-Host ""
    Write-Host "Troubleshooting steps:" -ForegroundColor Yellow
    Write-Host "  1. Run debug script: .\debug-integration.ps1 -Detailed"
    Write-Host "  2. Check service logs: Get-Content 'C:\Logs\ExcelAddin\*-stderr.log' -Tail 50"
    Write-Host "  3. Check Windows Event Logs"
    Write-Host "  4. Verify prerequisites (Node.js, Python, NSSM for backend)"
    
    exit 1
}