# ExcelAddin Backend Deployment Script
# Deploys Python Flask backend as NSSM service for Windows Server 10

param(
    [switch]$Force,
    [switch]$SkipInstall,
    [switch]$Debug
)

$ErrorActionPreference = "Stop"

# Service configuration
$ServiceName = "ExcelAddin-Backend"
$ServiceDisplayName = "ExcelAddin Backend Service"
$ServiceDescription = "Excel Add-in Backend API Service"
$Port = 5000

# Paths
$ProjectRoot = Split-Path $PSScriptRoot -Parent
$BackendPath = Join-Path $ProjectRoot "backend"
$ServiceScript = Join-Path $BackendPath "run.py"
$LogDir = "C:\Logs\ExcelAddin"

Write-Host "========================================" -ForegroundColor Green
Write-Host "  ExcelAddin Backend Deployment (NSSM)" -ForegroundColor Green  
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

# Check prerequisites
Write-Host "Checking prerequisites..." -ForegroundColor Yellow

# Check Python
$pythonCmd = Get-Command python -ErrorAction SilentlyContinue
if ($pythonCmd) {
    $pythonPath = $pythonCmd.Source
} else {
    Write-Error "Python not found. Please install Python 3.8+ and add to PATH."
}
$pythonVersion = python --version 2>&1
Write-Host "  Python: $pythonVersion" -ForegroundColor Green

# Check NSSM
$nssmCmd = Get-Command nssm -ErrorAction SilentlyContinue
if ($nssmCmd) {
    $nssmPath = $nssmCmd.Source
} else {
    Write-Error "NSSM not found. Please install NSSM and add to PATH."
}
Write-Host "  NSSM: Found at $nssmPath" -ForegroundColor Green

# Verify paths
Write-Host "Verifying paths..." -ForegroundColor Yellow
Write-Host "  Project Root: $ProjectRoot"
Write-Host "  Backend Path: $BackendPath"

if (-not (Test-Path $BackendPath)) {
    Write-Error "Backend directory not found: $BackendPath"
}

if (-not (Test-Path $ServiceScript)) {
    Write-Error "Service script not found: $ServiceScript"
}

Write-Host "  Service Script: $ServiceScript" -ForegroundColor Green

# Create log directory
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
    Write-Host "  Created log directory: $LogDir" -ForegroundColor Green
} else {
    Write-Host "  Log directory exists: $LogDir" -ForegroundColor Green
}

# Check if port is in use (skip backend itself)
$existingProcess = Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue | 
    Where-Object { $_.State -eq "Listen" }
if ($existingProcess) {
    $processId = $existingProcess.OwningProcess
    $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
    if ($process) {
        $processName = $process.ProcessName
    } else {
        $processName = "Unknown"
    }
    if ($processName -ne "python" -and -not $Force) {
        Write-Warning "Port $Port is in use by process: $processName (PID: $processId)"
        Write-Warning "Use -Force to override"
        exit 1
    }
}

# Stop and remove existing service
$existingService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($existingService) {
    Write-Host "Stopping existing service..." -ForegroundColor Yellow
    
    if ($existingService.Status -eq "Running") {
        Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3
    }
    
    Write-Host "Removing existing service..." -ForegroundColor Yellow
    nssm remove $ServiceName confirm
    Start-Sleep -Seconds 2
    
    # Verify removal
    $checkService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($checkService) {
        Write-Error "Failed to remove existing service"
    }
}

# Install dependencies if not skipping
if (-not $SkipInstall) {
    Write-Host "Installing Python dependencies..." -ForegroundColor Yellow
    Set-Location $BackendPath
    
    # Check for Poetry first (preferred), then pip
    $poetryCmd = Get-Command poetry -ErrorAction SilentlyContinue
    if ($poetryCmd) {
        $poetryPath = $poetryCmd.Source
        Write-Host "  Using Poetry for dependency management" -ForegroundColor Green
        poetry install
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Poetry install failed"
        }
    } else {
        Write-Host "  Using pip for dependency management" -ForegroundColor Yellow
        if (Test-Path "requirements.txt") {
            pip install -r requirements.txt
            if ($LASTEXITCODE -ne 0) {
                Write-Error "pip install failed"
            }
        } else {
            Write-Warning "No requirements.txt found, skipping dependency install"
        }
    }
}

# Install new service
Write-Host "Installing NSSM service..." -ForegroundColor Yellow
nssm install $ServiceName $pythonPath $ServiceScript

if ($LASTEXITCODE -ne 0) {
    Write-Error "Failed to install NSSM service"
}

# Configure service parameters
Write-Host "Configuring service..." -ForegroundColor Yellow
nssm set $ServiceName DisplayName $ServiceDisplayName
nssm set $ServiceName Description $ServiceDescription  
nssm set $ServiceName AppDirectory $BackendPath
nssm set $ServiceName Start SERVICE_AUTO_START

# Set environment variables
nssm set $ServiceName AppEnvironmentExtra "FLASK_ENV=production;PORT=$Port;HOST=127.0.0.1"

# Configure logging
nssm set $ServiceName AppStdout "$LogDir\backend-stdout.log"
nssm set $ServiceName AppStderr "$LogDir\backend-stderr.log"
nssm set $ServiceName AppStdoutCreationDisposition 4  # FILE_OPEN_ALWAYS
nssm set $ServiceName AppStderrCreationDisposition 4  # FILE_OPEN_ALWAYS

# Configure service restart behavior
nssm set $ServiceName AppThrottle 1500
nssm set $ServiceName AppRestartDelay 0
nssm set $ServiceName AppStopMethodSkip 0
nssm set $ServiceName AppExit Default Restart

if ($Debug) {
    Write-Host "Service configuration:" -ForegroundColor Cyan
    nssm dump $ServiceName
}

# Start service
Write-Host "Starting service..." -ForegroundColor Yellow
Start-Service -Name $ServiceName

# Wait for service to start
Start-Sleep -Seconds 5

# Verify service status
$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if (-not $service -or $service.Status -ne "Running") {
    Write-Error "Service failed to start. Check logs at: $LogDir"
}

Write-Host "Service started successfully!" -ForegroundColor Green
Write-Host "  Status: $($service.Status)" -ForegroundColor Green

# Test connectivity
Write-Host "Testing backend connectivity..." -ForegroundColor Yellow
Start-Sleep -Seconds 2

try {
    $response = Invoke-RestMethod -Uri "http://localhost:$Port/api/health" -TimeoutSec 10 -ErrorAction SilentlyContinue
    if ($response) {
        Write-Host "  Backend health check: PASSED" -ForegroundColor Green
        if ($Debug) {
            Write-Host "  Response: $($response | ConvertTo-Json -Compress)" -ForegroundColor Cyan
        }
    } else {
        Write-Warning "  Backend health check: No response"
    }
} catch {
    Write-Warning "  Backend health check: FAILED - $($_.Exception.Message)"
    Write-Host "  Check service logs: $LogDir\backend-stderr.log" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Backend Deployment Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Service Details:" -ForegroundColor Cyan
Write-Host "  Name: $ServiceName"
Write-Host "  Port: $Port"
Write-Host "  Logs: $LogDir"
Write-Host ""
Write-Host "Management Commands:" -ForegroundColor Cyan
Write-Host "  Start:   Start-Service -Name '$ServiceName'"
Write-Host "  Stop:    Stop-Service -Name '$ServiceName'"  
Write-Host "  Status:  Get-Service -Name '$ServiceName'"
Write-Host "  Logs:    Get-Content '$LogDir\backend-stderr.log' -Tail 50"
Write-Host ""