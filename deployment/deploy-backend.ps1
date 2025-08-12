#!/usr/bin/env powershell
# ExcelAddin Backend Deployment Script
# Deploys Python Flask backend as NSSM service

param(
    [switch]$Force,
    [switch]$SkipInstall
)

# Import common functions
. "$PSScriptRoot\scripts\common.ps1"

$ServiceName = "ExcelAddin-Backend"
$ServiceDisplayName = "ExcelAddin Backend Service"
$ServiceDescription = "Excel Add-in Backend API Service"

Write-Header "ExcelAddin Backend Deployment"

# Check prerequisites
if (-not (Test-Prerequisites -SkipPM2)) {
    Write-Error "Prerequisites check failed. Please resolve issues before continuing."
    exit 1
}

# Get paths
$ProjectRoot = Get-ProjectRoot
$BackendPath = Get-BackendPath
$ServiceScript = Join-Path $PSScriptRoot "config/backend-service.py"

Write-Host "Project Root: $ProjectRoot"
Write-Host "Backend Path: $BackendPath"

# Verify backend directory exists
if (-not (Test-Path $BackendPath)) {
    Write-Error "Backend directory not found: $BackendPath"
    exit 1
}

# Navigate to backend directory
Push-Location $BackendPath

try {
    # Install Python dependencies
    Write-Header "Installing Python Dependencies"
    
    if (Test-Path "pyproject.toml") {
        if (Test-Command "poetry") {
            Write-Host "Installing dependencies with Poetry..."
            poetry install
            if ($LASTEXITCODE -ne 0) {
                Write-Error "Poetry install failed"
                exit 1
            }
        } else {
            Write-Warning "Poetry not found. Attempting pip install..."
            if (Test-Path "requirements.txt") {
                pip install -r requirements.txt
                if ($LASTEXITCODE -ne 0) {
                    Write-Error "Pip install failed"
                    exit 1
                }
            } else {
                Write-Error "Neither Poetry nor requirements.txt found"
                exit 1
            }
        }
    } else {
        Write-Error "No pyproject.toml found in backend directory"
        exit 1
    }
    
    # Setup environment file
    Write-Header "Configuring Environment"
    
    $envFile = Join-Path $BackendPath ".env"
    $stagingEnvFile = Join-Path $BackendPath ".env.staging"
    
    if (-not (Test-Path $envFile) -and (Test-Path $stagingEnvFile)) {
        Copy-Item $stagingEnvFile $envFile
        Write-Success "Copied staging environment configuration"
    } elseif (Test-Path $envFile) {
        Write-Success "Environment file already exists"
    } else {
        Write-Error "No environment configuration found"
        exit 1
    }
    
    # Test backend application
    Write-Header "Testing Backend Application"
    
    Write-Host "Starting backend test..."
    $testProcess = Start-Process -FilePath "python" -ArgumentList "run.py" -PassThru -NoNewWindow
    Start-Sleep -Seconds 5
    
    if ($testProcess -and -not $testProcess.HasExited) {
        Write-Host "Testing health endpoint..."
        try {
            $response = Invoke-RestMethod -Uri "http://127.0.0.1:5000/api/health" -TimeoutSec 10
            Write-Success "Backend health check passed: $($response.status)"
        } catch {
            Write-Warning "Health check failed, but continuing deployment: $($_.Exception.Message)"
        }
        
        # Stop test process
        Stop-Process -Id $testProcess.Id -Force
        Start-Sleep -Seconds 2
    } else {
        Write-Error "Backend failed to start during test"
        exit 1
    }
    
    # Configure NSSM service
    if (-not $SkipInstall) {
        Write-Header "Configuring NSSM Service"
        
        # Remove existing service if it exists
        $existingService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if ($existingService) {
            if ($Force) {
                Write-Host "Removing existing service..."
                Stop-ServiceSafely -ServiceName $ServiceName
                nssm remove $ServiceName confirm
                Start-Sleep -Seconds 2
            } else {
                Write-Error "Service $ServiceName already exists. Use -Force to overwrite."
                exit 1
            }
        }
        
        # Install NSSM service
        Write-Host "Installing NSSM service..."
        $pythonPath = (Get-Command python).Source
        nssm install $ServiceName $pythonPath $ServiceScript
        
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Failed to install NSSM service"
            exit 1
        }
        
        # Configure service parameters
        nssm set $ServiceName DisplayName $ServiceDisplayName
        nssm set $ServiceName Description $ServiceDescription
        nssm set $ServiceName AppDirectory $BackendPath
        nssm set $ServiceName Start SERVICE_AUTO_START
        
        # Configure logging
        $logDir = "C:\Logs\ExcelAddin"
        if (-not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }
        
        nssm set $ServiceName AppStdout "$logDir\backend-stdout.log"
        nssm set $ServiceName AppStderr "$logDir\backend-stderr.log"
        nssm set $ServiceName AppRotateFiles 1
        nssm set $ServiceName AppRotateOnline 1
        nssm set $ServiceName AppRotateSeconds 86400
        nssm set $ServiceName AppRotateBytes 10485760
        
        # Set environment variables
        nssm set $ServiceName AppEnvironmentExtra ENVIRONMENT=staging PYTHONPATH=$BackendPath
        
        Write-Success "NSSM service configured successfully"
    }
    
    # Start the service
    Write-Header "Starting Backend Service"
    
    try {
        Start-Service -Name $ServiceName
        Start-Sleep -Seconds 5
        
        $service = Get-Service -Name $ServiceName
        if ($service.Status -eq "Running") {
            Write-Success "Backend service started successfully"
            
            # Verify service is responding
            Start-Sleep -Seconds 5
            try {
                $response = Invoke-RestMethod -Uri "http://127.0.0.1:5000/api/health" -TimeoutSec 10
                Write-Success "Backend service is responding: $($response.status)"
            } catch {
                Write-Warning "Backend service is running but not responding to health check"
            }
        } else {
            Write-Error "Backend service failed to start"
            exit 1
        }
    } catch {
        Write-Error "Failed to start backend service: $($_.Exception.Message)"
        exit 1
    }
    
    Write-Header "Backend Deployment Complete"
    Write-Success "ExcelAddin Backend has been deployed successfully"
    Write-Host ""
    Write-Host "Service Information:"
    Write-Host "  Name: $ServiceName"
    Write-Host "  Display Name: $ServiceDisplayName"
    Write-Host "  Status: Running"
    Write-Host "  Health Check: http://127.0.0.1:5000/api/health"
    Write-Host ""

} catch {
    Write-Error "Deployment failed: $($_.Exception.Message)"
    Write-Host $_.ScriptStackTrace
    exit 1
} finally {
    Pop-Location
}