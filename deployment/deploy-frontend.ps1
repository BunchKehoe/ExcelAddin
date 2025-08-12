# ExcelAddin Frontend Deployment Script
# Simple, Windows 10 Server compatible deployment script

param(
    [switch]$SkipBuild,
    [switch]$SkipInstall,
    [switch]$ConfigureIIS
)

# Import common functions
. (Join-Path $PSScriptRoot "scripts\common.ps1")

$ServiceName = "ExcelAddin-Frontend"
$ServiceDisplayName = "ExcelAddin Frontend Service"
$ServiceDescription = "Excel Add-in Frontend Web Server"
$FrontendServerScript = Join-Path $PSScriptRoot "config\frontend-server.js"

Write-Header "ExcelAddin Frontend Deployment"

# Check prerequisites
Write-Host "Checking prerequisites..."
if (-not (Test-Prerequisites -SkipPM2 -SkipNSSM:$false)) {
    Write-Error "Prerequisites check failed. Please resolve issues before continuing."
    exit 1
}

# Get paths
$ProjectRoot = Get-ProjectRoot
$FrontendPath = Get-FrontendPath

Write-Host "Project Root: $ProjectRoot"
Write-Host "Frontend Path: $FrontendPath"

# Navigate to project root
Set-Location $FrontendPath

# Install dependencies
Write-Header "Installing Node.js Dependencies"
Write-Host "Running npm install..."
npm install
if ($LASTEXITCODE -ne 0) {
    Write-Error "npm install failed"
    exit 1
}
Write-Success "Dependencies installed successfully"

# Build frontend
if (-not $SkipBuild) {
    Write-Header "Building Frontend Application"
    Write-Host "Building for staging environment..."
    npm run build:staging
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Frontend build failed"
        exit 1
    }
    
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

# Configure NSSM service
if (-not $SkipInstall) {
    Write-Header "Configuring NSSM Service"
    
    # Stop and remove existing service
    $existingService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($existingService) {
        Write-Host "Stopping existing service: $ServiceName"
        Stop-ServiceSafely -ServiceName $ServiceName
        nssm remove $ServiceName confirm
        Start-Sleep -Seconds 3
        Write-Success "Existing service removed"
    }
    
    # Check for port conflicts and kill them
    Write-Host "Checking for port conflicts on 3000..."
    $portConflict = Get-NetTCPConnection -LocalPort 3000 -ErrorAction SilentlyContinue
    if ($portConflict) {
        $conflictProcess = Get-Process -Id $portConflict.OwningProcess -ErrorAction SilentlyContinue
        if ($conflictProcess -and $conflictProcess.Id -gt 4) {
            Write-Host "Killing process using port 3000: $($conflictProcess.ProcessName) (PID: $($conflictProcess.Id))"
            Stop-Process -Id $conflictProcess.Id -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
        }
    }
    
    # Get Node.js path
    $nodeCmd = Get-Command node -ErrorAction Stop
    $nodePath = $nodeCmd.Source
    Write-Host "Using Node.js: $nodePath"
    
    # Verify server script exists
    if (-not (Test-Path $FrontendServerScript)) {
        Write-Error "Frontend server script not found: $FrontendServerScript"
        exit 1
    }
    
    # Test server script syntax
    Write-Host "Testing server script syntax..."
    $syntaxTest = & node -c $FrontendServerScript 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Server script has syntax errors: $syntaxTest"
        exit 1
    }
    Write-Success "Server script syntax OK"
    
    # Install NSSM service
    Write-Host "Installing NSSM service..."
    nssm install $ServiceName $nodePath $FrontendServerScript
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to install NSSM service"
        exit 1
    }
    
    # Configure service - use absolute paths for reliability
    nssm set $ServiceName DisplayName $ServiceDisplayName
    nssm set $ServiceName Description $ServiceDescription
    nssm set $ServiceName AppDirectory $FrontendPath
    nssm set $ServiceName Start SERVICE_AUTO_START
    nssm set $ServiceName AppExit Default Restart
    nssm set $ServiceName AppRestartDelay 5000
    nssm set $ServiceName AppThrottle 5000
    
    # Setup logging directory
    $logDir = "C:\Logs\ExcelAddin"
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        Write-Host "Created log directory: $logDir"
    }
    
    # Configure logging
    nssm set $ServiceName AppStdout "$logDir\frontend-stdout.log"
    nssm set $ServiceName AppStderr "$logDir\frontend-stderr.log"
    nssm set $ServiceName AppRotateFiles 1
    nssm set $ServiceName AppRotateOnline 1
    nssm set $ServiceName AppRotateSeconds 86400
    nssm set $ServiceName AppRotateBytes 10485760
    
    # Set environment variables for service
    nssm set $ServiceName AppEnvironmentExtra "NODE_ENV=production;PORT=3000;HOST=127.0.0.1"
    
    Write-Success "NSSM service configured successfully"
    
    # Verify NSSM configuration
    Write-Host ""
    Write-Host "NSSM Configuration Verification:"
    $nssmApp = & nssm get $ServiceName Application 2>&1
    $nssmParams = & nssm get $ServiceName Parameters 2>&1
    $nssmDir = & nssm get $ServiceName AppDirectory 2>&1
    $nssmEnv = & nssm get $ServiceName AppEnvironmentExtra 2>&1
    
    Write-Host "  Application: $nssmApp"
    Write-Host "  Parameters: $nssmParams"
    Write-Host "  Directory: $nssmDir"
    Write-Host "  Environment: $nssmEnv"
}

# Start service
Write-Header "Starting Frontend Service"

# Clear any existing log files for fresh start
$logFiles = @(
    "C:\Logs\ExcelAddin\frontend-stdout.log",
    "C:\Logs\ExcelAddin\frontend-stderr.log"
)
foreach ($logFile in $logFiles) {
    if (Test-Path $logFile) {
        try {
            Clear-Content $logFile -Force
            Write-Host "Cleared log file: $logFile"
        } catch {
            Write-Warning "Could not clear log file: $logFile"
        }
    }
}

Write-Host "Starting service: $ServiceName"
Start-Service -Name $ServiceName

# Wait for service to start with better monitoring
$startTimeout = 30
$elapsed = 0
$serviceStarted = $false

Write-Host "Waiting for service to start..."
while ($elapsed -lt $startTimeout -and -not $serviceStarted) {
    Start-Sleep -Seconds 1
    $elapsed++
    
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($service -and $service.Status -eq "Running") {
        $serviceStarted = $true
        Write-Success "Service started successfully (${elapsed}s)"
    } else {
        Write-Host "." -NoNewline
    }
}
Write-Host ""

if (-not $serviceStarted) {
    Write-Error "Service failed to start within $startTimeout seconds"
    
    # Show any error logs
    $errorLog = "C:\Logs\ExcelAddin\frontend-stderr.log"
    if (Test-Path $errorLog) {
        Write-Host "Error log contents:"
        Get-Content $errorLog -ErrorAction SilentlyContinue | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    }
    exit 1
}

# Enhanced service testing
Write-Header "Testing Service"

Write-Host "Waiting for server to initialize..."
Start-Sleep -Seconds 5

# Check if port is actually listening
$maxPortTests = 10
$portTestAttempts = 0
$portListening = $false

while ($portTestAttempts -lt $maxPortTests -and -not $portListening) {
    $portTestAttempts++
    Write-Host "Port test attempt $portTestAttempts/$maxPortTests..."
    
    $portTest = Test-NetConnection -ComputerName "127.0.0.1" -Port 3000 -InformationLevel Quiet -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    if ($portTest) {
        Write-Success "Port 3000 is listening"
        $portListening = $true
    } else {
        Write-Host "Port 3000 not yet listening, waiting..."
        Start-Sleep -Seconds 2
    }
}

if (-not $portListening) {
    Write-Error "Service is running but not listening on port 3000"
    
    # Show logs for debugging
    Write-Host "Service output logs:"
    $stdoutLog = "C:\Logs\ExcelAddin\frontend-stdout.log"
    if (Test-Path $stdoutLog) {
        Write-Host "STDOUT:"
        Get-Content $stdoutLog -ErrorAction SilentlyContinue | ForEach-Object { Write-Host "  $_" }
    }
    
    $stderrLog = "C:\Logs\ExcelAddin\frontend-stderr.log"
    if (Test-Path $stderrLog) {
        Write-Host "STDERR:"
        Get-Content $stderrLog -ErrorAction SilentlyContinue | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    }
    
    Write-Host ""
    Write-Host "Troubleshooting steps:"
    Write-Host "1. Run diagnostics: .\diagnose-nssm.ps1"
    Write-Host "2. Test manually: .\test-frontend-server.ps1"
    Write-Host "3. Check logs: C:\Logs\ExcelAddin\frontend-*.log"
    
    exit 1
}

# Test HTTP response
Write-Host "Testing HTTP response..."
$httpTestAttempts = 0
$maxHttpTests = 5
$httpSuccess = $false

while ($httpTestAttempts -lt $maxHttpTests -and -not $httpSuccess) {
    $httpTestAttempts++
    try {
        Write-Host "HTTP test attempt $httpTestAttempts/$maxHttpTests..."
        $response = Invoke-WebRequest -Uri "http://127.0.0.1:3000" -TimeoutSec 10 -ErrorAction Stop
        Write-Success "HTTP test passed - Status: $($response.StatusCode)"
        Write-Host "Response length: $($response.Content.Length) bytes"
        
        if ($response.Content -like "*<html*" -or $response.Content -like "*<!DOCTYPE*") {
            Write-Success "Valid HTML content detected"
            $httpSuccess = $true
        } else {
            Write-Warning "Response doesn't appear to be HTML"
            Write-Host "First 200 characters: $($response.Content.Substring(0, [Math]::Min(200, $response.Content.Length)))"
        }
        
    } catch {
        Write-Warning "HTTP test $httpTestAttempts failed: $($_.Exception.Message)"
        if ($httpTestAttempts -lt $maxHttpTests) {
            Start-Sleep -Seconds 3
        }
    }
}

if (-not $httpSuccess) {
    Write-Warning "HTTP tests failed - service may not be working correctly"
    Write-Host ""
    Write-Host "Troubleshooting steps:"
    Write-Host "1. Check service logs: C:\Logs\ExcelAddin\frontend-*.log"
    Write-Host "2. Run manual test: .\test-frontend-server.ps1"
    Write-Host "3. Run diagnostics: .\diagnose-nssm.ps1"
} else {
    Write-Success "Service is working correctly"
}

# Configure IIS if requested
if ($ConfigureIIS) {
    Write-Header "Configuring IIS"
    
    $SiteName = "ExcelAddin"
    $Port = 9443
    
    # Import WebAdministration
    Import-Module WebAdministration -ErrorAction SilentlyContinue
    if (-not (Get-Module WebAdministration)) {
        Write-Error "WebAdministration module not available"
        exit 1
    }
    
    # Remove existing site
    $existingSite = Get-Website -Name $SiteName -ErrorAction SilentlyContinue
    if ($existingSite) {
        Write-Host "Removing existing IIS site..."
        Remove-Website -Name $SiteName
        Start-Sleep -Seconds 2
    }
    
    # Remove any conflicting applications under Default Web Site
    $conflictingApp = Get-WebApplication -Site "Default Web Site" -Name $SiteName -ErrorAction SilentlyContinue
    if ($conflictingApp) {
        Write-Host "Removing conflicting application under Default Web Site..."
        Remove-WebApplication -Site "Default Web Site" -Name $SiteName
    }
    
    # Create application directory
    $appPath = "C:\inetpub\wwwroot\$SiteName"
    if (-not (Test-Path $appPath)) {
        New-Item -ItemType Directory -Path $appPath -Force | Out-Null
    }
    
    # Copy web.config
    $webConfigSource = Join-Path $PSScriptRoot "config\web.config"
    if (Test-Path $webConfigSource) {
        Copy-Item $webConfigSource "$appPath\web.config" -Force
        Write-Host "Copied web.config"
    }
    
    # Create and configure application pool first
    $appPoolName = "$SiteName-AppPool"
    $existingAppPool = Get-WebAppPool -Name $appPoolName -ErrorAction SilentlyContinue
    if ($existingAppPool) {
        Write-Host "Removing existing application pool..."
        Remove-WebAppPool -Name $appPoolName
    }
    
    Write-Host "Creating application pool: $appPoolName"
    New-WebAppPool -Name $appPoolName
    Set-ItemProperty -Path "IIS:\AppPools\$appPoolName" -Name managedRuntimeVersion -Value ""
    
    # Create standalone IIS site with HTTPS binding
    Write-Host "Creating IIS site: $SiteName"
    $site = New-Website -Name $SiteName -PhysicalPath $appPath -Port $Port -Protocol https -ApplicationPool $appPoolName
    if (-not $site) {
        Write-Error "Failed to create IIS site"
        exit 1
    }
    
    Write-Success "IIS site created as standalone site"
    
    # Handle SSL certificate
    $CertPath = "C:\Cert"
    $certThumbprint = ""
    
    if (Test-Path $CertPath) {
        Write-Host "Looking for certificates in $CertPath..."
        $certFiles = Get-ChildItem -Path $CertPath -Include "*.pfx", "*.p12" -ErrorAction SilentlyContinue
        
        if ($certFiles) {
            $certFile = $certFiles | Select-Object -First 1
            Write-Host "Found certificate: $($certFile.Name)"
            
            try {
                $importResult = Import-PfxCertificate -FilePath $certFile.FullName -CertStoreLocation "Cert:\LocalMachine\My" -Password (ConvertTo-SecureString -String "" -AsPlainText -Force) -ErrorAction Stop
                $certThumbprint = $importResult.Thumbprint
                Write-Success "Imported certificate: $certThumbprint"
            } catch {
                Write-Warning "Failed to import certificate: $($_.Exception.Message)"
            }
        }
    }
    
    # Bind certificate if available
    if ($certThumbprint) {
        Write-Host "Binding SSL certificate..."
        
        # Remove existing binding
        netsh http delete sslcert ipport=0.0.0.0:$Port 2>$null
        
        # Add new binding using netsh
        $guid = "{12345678-1234-1234-1234-123456789abc}"
        netsh http add sslcert ipport=0.0.0.0:$Port certhash=$certThumbprint appid=$guid certstorename=MY
        
        if ($LASTEXITCODE -eq 0) {
            Write-Success "SSL certificate bound successfully"
        } else {
            Write-Warning "SSL certificate binding failed - you may need to bind manually"
        }
    } else {
        Write-Warning "No SSL certificate found - place .pfx or .p12 file in C:\Cert\"
    }
    
    # Start site
    Start-Website -Name $SiteName
    Write-Success "IIS site started"
    
    Write-Host ""
    Write-Host "IIS Configuration:"
    Write-Host "  Site: $SiteName (standalone)"
    Write-Host "  Port: $Port"
    Write-Host "  Path: $appPath"
    Write-Host "  App Pool: $appPoolName"
}

Write-Header "Deployment Complete"
Write-Success "Frontend deployment completed successfully"
Write-Host ""
Write-Host "Service: $ServiceName"
Write-Host "Status: Running"
Write-Host "URL: http://127.0.0.1:3000"

if ($ConfigureIIS) {
    Write-Host "IIS: https://server-vs81t.intranet.local:9443"
}

Write-Host ""
Write-Host "Service commands:"
Write-Host "  Status: Get-Service -Name '$ServiceName'"
Write-Host "  Stop: Stop-Service -Name '$ServiceName'"
Write-Host "  Start: Start-Service -Name '$ServiceName'"
Write-Host "  Logs: C:\Logs\ExcelAddin\frontend-*.log"