# ExcelAddin Frontend Deployment Script
# Deploys React frontend as PM2 service

param(
    [switch]$Force,
    [switch]$SkipBuild
)

# Import common functions
. "$PSScriptRoot\scripts\common.ps1"

$AppName = "exceladdin-frontend"
$PM2Config = "$PSScriptRoot/config/pm2-frontend.json"

Write-Header "ExcelAddin Frontend Deployment"

# Check prerequisites
if (-not (Test-Prerequisites -SkipNSSM)) {
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
    
    # Check if serve package is available globally
    if (-not (Test-Command "serve")) {
        Write-Host "Installing serve package globally..."
        npm install -g serve
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Failed to install serve package"
            exit 1
        }
    }
    
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
    
    # Start serve in background to test
    Write-Host "Testing frontend build..."
    $testProcess = Start-Process -FilePath "npx" -ArgumentList "serve", "-s", "dist", "-l", "3001" -PassThru -NoNewWindow
    Start-Sleep -Seconds 5
    
    if ($testProcess -and -not $testProcess.HasExited) {
        try {
            $response = Invoke-WebRequest -Uri "http://127.0.0.1:3001" -TimeoutSec 10
            Write-Success "Frontend test passed - HTTP Status: $($response.StatusCode)"
        } catch {
            Write-Warning "Frontend test failed, but continuing: $($_.Exception.Message)"
        }
        
        # Stop test process
        Stop-Process -Id $testProcess.Id -Force
        Start-Sleep -Seconds 2
    }
    
    # Configure PM2 application
    Write-Header "Configuring PM2 Service"
    
    # Check if application already exists
    $existingApp = pm2 list | Where-Object { $_ -match $AppName }
    if ($existingApp) {
        if ($Force) {
            Write-Host "Stopping existing PM2 application..."
            pm2 stop $AppName
            pm2 delete $AppName
            Start-Sleep -Seconds 2
        } else {
            Write-Error "PM2 application $AppName already exists. Use -Force to overwrite."
            exit 1
        }
    }
    
    # Create logs directory
    $logsDir = Join-Path $FrontendPath "logs"
    if (-not (Test-Path $logsDir)) {
        New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
        Write-Host "Created logs directory: $logsDir"
    }
    
    # Start PM2 application
    Write-Host "Starting PM2 application from config: $PM2Config"
    pm2 start $PM2Config
    
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to start PM2 application"
        exit 1
    }
    
    # Save PM2 configuration
    pm2 save
    
    # Verify application is running
    Start-Sleep -Seconds 5
    $appStatus = pm2 jlist | ConvertFrom-Json | Where-Object { $_.name -eq $AppName }
    
    if ($appStatus -and $appStatus.pm2_env.status -eq "online") {
        Write-Success "PM2 application started successfully"
        
        # Test if frontend is responding
        try {
            Start-Sleep -Seconds 3
            $response = Invoke-WebRequest -Uri "http://127.0.0.1:3000" -TimeoutSec 15
            Write-Success "Frontend service is responding - HTTP Status: $($response.StatusCode)"
        } catch {
            Write-Warning "Frontend service started but not responding to HTTP requests: $($_.Exception.Message)"
        }
    } else {
        Write-Error "Frontend service failed to start"
        pm2 logs $AppName --lines 20
        exit 1
    }
    
    # Configure PM2 startup (Windows service)
    Write-Header "Configuring PM2 Startup"
    
    try {
        # Generate startup script
        $startupResult = pm2 startup
        Write-Host $startupResult
        
        # Save current PM2 configuration
        pm2 save
        
        Write-Success "PM2 startup configuration completed"
        Write-Host "Note: PM2 startup may require additional manual configuration on Windows"
        
    } catch {
        Write-Warning "PM2 startup configuration failed: $($_.Exception.Message)"
        Write-Host "You may need to configure PM2 to start on system boot manually"
    }
    
    Write-Header "Frontend Deployment Complete"
    Write-Success "ExcelAddin Frontend has been deployed successfully"
    Write-Host ""
    Write-Host "Application Information:"
    Write-Host "  Name: $AppName"
    Write-Host "  Status: Online"
    Write-Host "  URL: http://127.0.0.1:3000"
    Write-Host "  Logs: pm2 logs $AppName"
    Write-Host ""
    Write-Host "PM2 Management Commands:"
    Write-Host "  Status: pm2 status"
    Write-Host "  Restart: pm2 restart $AppName"
    Write-Host "  Stop: pm2 stop $AppName"
    Write-Host "  Logs: pm2 logs $AppName"
    Write-Host ""

} catch {
    Write-Error "Deployment failed: $($_.Exception.Message)"
    Write-Host $_.ScriptStackTrace
    exit 1
} finally {
    Pop-Location
}