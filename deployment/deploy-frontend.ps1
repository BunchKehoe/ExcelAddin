# ExcelAddin Frontend Deployment Script
# Deploys React frontend as NSSM service

param(
    [switch]$Force,      # Kept for compatibility but not required for overwriting services
    [switch]$SkipBuild,
    [switch]$SkipInstall
)

# Import common functions
. "$PSScriptRoot\scripts\common.ps1"

$ServiceName = "ExcelAddin-Frontend"
$ServiceDisplayName = "ExcelAddin Frontend Service"
$ServiceDescription = "Excel Add-in Frontend Web Server"
$FrontendServerScript = "$PSScriptRoot/config/frontend-server.js"

Write-Header "ExcelAddin Frontend Deployment"

# Check prerequisites
if (-not (Test-Prerequisites -SkipPM2)) {
    Write-Error "Prerequisites check failed. Please resolve issues before continuing."
    exit 1
}

# Get paths
$ProjectRoot = Get-ProjectRoot
$FrontendPath = Get-FrontendPath

Write-Host "Project Root: $ProjectRoot"
Write-Host "Frontend Path: $FrontendPath"

# Navigate to project root
Push-Location $FrontendPath

try {
    # Install Node.js dependencies
    Write-Header "Installing Node.js Dependencies"
    
    Write-Host "Running npm install..."
    npm install
    if ($LASTEXITCODE -ne 0) {
        Write-Error "npm install failed"
        exit 1
    }
    Write-Success "Dependencies installed successfully"
    
    # Build the frontend
    if (-not $SkipBuild) {
        Write-Header "Building Frontend Application"
        
        Write-Host "Building for staging environment..."
        npm run build:staging
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Frontend build failed"
            exit 1
        }
        
        # Verify dist directory exists
        if (-not (Test-Path "dist")) {
            Write-Error "Build output directory 'dist' not found"
            exit 1
        }
        
        Write-Success "Frontend built successfully"
    } else {
        Write-Warning "Skipping build - using existing dist directory"
        if (-not (Test-Path "dist")) {
            Write-Error "dist directory not found. Cannot skip build."
            exit 1
        }
    }
    
    # Test the built frontend
    Write-Header "Testing Frontend Build"
    
    # Test using our new frontend server
    Write-Host "Testing frontend build with new server..."
    $nodeCmd = Get-Command node -ErrorAction Stop
    $testProcess = Start-Process -FilePath $nodeCmd.Source -ArgumentList $FrontendServerScript -PassThru -NoNewWindow -WorkingDirectory $FrontendPath
    Start-Sleep -Seconds 5
    
    if ($testProcess -and -not $testProcess.HasExited) {
        try {
            $response = Invoke-WebRequest -Uri "http://127.0.0.1:3000" -TimeoutSec 10
            Write-Success "Frontend test passed - HTTP Status: $($response.StatusCode)"
        } catch {
            Write-Warning "Frontend test failed, but continuing: $($_.Exception.Message)"
        }
        
        # Stop test process
        Stop-Process -Id $testProcess.Id -Force
        Start-Sleep -Seconds 2
    }
    
    # Configure NSSM service
    if (-not $SkipInstall) {
        Write-Header "Configuring NSSM Service"
        
        # Remove existing service if it exists (always overwrite by default)
        $existingService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if ($existingService) {
            Write-Host "Existing service detected. Stopping and removing service: $ServiceName"
            Stop-ServiceSafely -ServiceName $ServiceName
            nssm remove $ServiceName confirm
            Start-Sleep -Seconds 2
            Write-Success "Existing service removed successfully"
        }
        
        # Install NSSM service
        Write-Host "Installing NSSM service..."
        
        # Get Node.js executable path
        try {
            $nodeCmd = Get-Command node -ErrorAction Stop
            $nodePath = $nodeCmd.Source
            Write-Host "Using Node.js executable: $nodePath"
        } catch {
            Write-Error "Node.js executable not found in PATH. Please ensure Node.js is installed and accessible."
            exit 1
        }
        
        # Verify Node.js executable works
        try {
            $nodeVersion = & $nodePath --version 2>&1
            Write-Host "Node.js version: $nodeVersion"
        } catch {
            Write-Error "Node.js executable is not working properly: $($_.Exception.Message)"
            exit 1
        }
        
        # Verify frontend server script exists
        if (-not (Test-Path $FrontendServerScript)) {
            Write-Error "Frontend server script not found: $FrontendServerScript"
            exit 1
        }
        Write-Host "Using frontend server script: $FrontendServerScript"
        
        nssm install $ServiceName $nodePath $FrontendServerScript
        
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Failed to install NSSM service"
            exit 1
        }
        
        # Configure service parameters
        nssm set $ServiceName DisplayName $ServiceDisplayName
        nssm set $ServiceName Description $ServiceDescription
        nssm set $ServiceName AppDirectory $FrontendPath
        nssm set $ServiceName Start SERVICE_AUTO_START
        
        # Configure logging
        $logDir = "C:\Logs\ExcelAddin"
        if (-not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }
        
        nssm set $ServiceName AppStdout "$logDir\frontend-stdout.log"
        nssm set $ServiceName AppStderr "$logDir\frontend-stderr.log"
        nssm set $ServiceName AppRotateFiles 1
        nssm set $ServiceName AppRotateOnline 1
        nssm set $ServiceName AppRotateSeconds 86400
        nssm set $ServiceName AppRotateBytes 10485760
        
        # Set environment variables
        nssm set $ServiceName AppEnvironmentExtra NODE_ENV=production PORT=3000 HOST=127.0.0.1
        
        Write-Success "NSSM service configured successfully"
    }
    # Start the service
    Write-Header "Starting Frontend Service"
    
    try {
        Start-Service -Name $ServiceName
        Start-Sleep -Seconds 5
        
        $service = Get-Service -Name $ServiceName
        if ($service.Status -eq "Running") {
            Write-Success "Frontend service started successfully"
            
            # Verify service is responding
            Start-Sleep -Seconds 5
            try {
                $response = Invoke-WebRequest -Uri "http://127.0.0.1:3000" -TimeoutSec 10
                Write-Success "Frontend service is responding - HTTP Status: $($response.StatusCode)"
            } catch {
                Write-Warning "Frontend service is running but not responding to HTTP requests"
            }
        } else {
            Write-Error "Frontend service failed to start"
            exit 1
        }
    } catch {
        Write-Error "Failed to start frontend service: $($_.Exception.Message)"
        exit 1
    }
    
    Write-Header "Frontend Deployment Complete"
    Write-Success "ExcelAddin Frontend has been deployed successfully"
    Write-Host ""
    Write-Host "Service Information:"
    Write-Host "  Name: $ServiceName"
    Write-Host "  Display Name: $ServiceDisplayName"
    Write-Host "  Status: Running"
    Write-Host "  URL: http://127.0.0.1:3000"
    Write-Host ""
    Write-Host "Service Management Commands:"
    Write-Host "  Status: Get-Service -Name '$ServiceName'"
    Write-Host "  Start: Start-Service -Name '$ServiceName'"
    Write-Host "  Stop: Stop-Service -Name '$ServiceName'"
    Write-Host "  Restart: Restart-Service -Name '$ServiceName'"
    Write-Host "  Logs: Check C:\Logs\ExcelAddin\frontend-*.log"
    Write-Host ""

} catch {
    Write-Error "Deployment failed: $($_.Exception.Message)"
    Write-Host $_.ScriptStackTrace
    exit 1
} finally {
    Pop-Location
}