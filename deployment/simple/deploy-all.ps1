param(
    [string]$SiteName = "Default Web Site",
    [string]$ApplicationName = "excellence",
    [switch]$Force
)

Write-Host "=== Simple Complete Deployment ===" -ForegroundColor Cyan
Write-Host "This will deploy both frontend and backend to IIS" -ForegroundColor Cyan

$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$RepoRoot = Split-Path -Parent (Split-Path -Parent $ScriptPath)

try {
    # Build frontend first
    Write-Host "`n1. Building frontend..." -ForegroundColor Yellow
    Push-Location $RepoRoot
    
    # Check if package.json exists
    if (-not (Test-Path "package.json")) {
        Write-Error "package.json not found in $RepoRoot"
        exit 1
    }
    
    # Run staging build
    Write-Host "Running: npm run build:staging"
    npm run build:staging
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Frontend build failed"
        exit 1
    }
    
    Pop-Location
    
    # Deploy frontend
    Write-Host "`n2. Deploying frontend..." -ForegroundColor Yellow
    $FrontendScript = Join-Path $ScriptPath "deploy-frontend.ps1"
    if ($Force) {
        & $FrontendScript -SiteName $SiteName -ApplicationName $ApplicationName -Force
    } else {
        & $FrontendScript -SiteName $SiteName -ApplicationName $ApplicationName
    }
    
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Frontend deployment failed"
        exit 1
    }
    
    # Deploy backend
    Write-Host "`n3. Deploying backend..." -ForegroundColor Yellow
    $BackendScript = Join-Path $ScriptPath "deploy-backend.ps1"
    if ($Force) {
        & $BackendScript -SiteName $SiteName -ApplicationName $ApplicationName -Force
    } else {
        & $BackendScript -SiteName $SiteName -ApplicationName $ApplicationName
    }
    
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Backend deployment failed"
        exit 1
    }
    
    Write-Host "`n=== Deployment Complete ===" -ForegroundColor Green
    Write-Host "Frontend: https://localhost:9443/$ApplicationName/" -ForegroundColor Cyan
    Write-Host "Backend: https://localhost:9443/$ApplicationName/backend/api/health" -ForegroundColor Cyan
    Write-Host "`nRun test-deployment.ps1 to verify everything is working." -ForegroundColor Yellow
    
} catch {
    Write-Error "Deployment failed: $_"
    exit 1
} finally {
    Pop-Location -ErrorAction SilentlyContinue
}