# PowerShell script to set up nginx as a Windows service using NSSM
# Usage: .\setup-nginx-service.ps1 -NginxPath "C:\nginx"

param(
    [Parameter(Mandatory=$false)]
    [string]$NginxPath = "C:\nginx",
    
    [Parameter(Mandatory=$false)]
    [string]$ServiceName = "nginx",
    
    [Parameter(Mandatory=$false)]
    [switch]$Uninstall,
    
    [Parameter(Mandatory=$false)]
    [switch]$Force
)

# Check if running as administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Administrator)) {
    Write-Error "This script must be run as Administrator"
    exit 1
}

Write-Host "nginx Windows Service Setup" -ForegroundColor Cyan
Write-Host "=" * 40

# Check if nginx exists
if (-not (Test-Path "$NginxPath\nginx.exe")) {
    Write-Error "nginx.exe not found at $NginxPath\nginx.exe"
    exit 1
}

# Check if NSSM is available
try {
    $nssmVersion = & nssm version 2>$null
    if ($LASTEXITCODE -ne 0) {
        throw "NSSM not found"
    }
    Write-Host "✓ NSSM found: $nssmVersion" -ForegroundColor Green
} catch {
    Write-Error "NSSM (Non-Sucking Service Manager) is required but not found."
    Write-Host "Please download and install NSSM from: https://nssm.cc/" -ForegroundColor Yellow
    Write-Host "1. Download nssm from https://nssm.cc/download" -ForegroundColor Yellow
    Write-Host "2. Extract to C:\Tools\nssm (or add to PATH)" -ForegroundColor Yellow
    Write-Host "3. Re-run this script" -ForegroundColor Yellow
    exit 1
}

# Uninstall existing service if requested
if ($Uninstall) {
    Write-Host "Uninstalling nginx service..." -ForegroundColor Yellow
    
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($service) {
        if ($service.Status -eq 'Running') {
            Write-Host "Stopping nginx service..."
            Stop-Service -Name $ServiceName -Force
            Start-Sleep -Seconds 3
        }
        
        & nssm remove $ServiceName confirm
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ nginx service removed successfully" -ForegroundColor Green
        } else {
            Write-Error "Failed to remove nginx service"
        }
    } else {
        Write-Host "nginx service not found" -ForegroundColor Yellow
    }
    exit 0
}

# Stop nginx processes if running
Write-Host "Checking for running nginx processes..." -ForegroundColor Yellow
$nginxProcesses = Get-Process -Name "nginx" -ErrorAction SilentlyContinue
if ($nginxProcesses) {
    Write-Host "Stopping nginx processes..."
    Stop-Process -Name "nginx" -Force
    Start-Sleep -Seconds 2
}

# Stop existing service if it exists
$existingService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($existingService) {
    if ($Force) {
        Write-Host "Removing existing nginx service..." -ForegroundColor Yellow
        if ($existingService.Status -eq 'Running') {
            Stop-Service -Name $ServiceName -Force
            Start-Sleep -Seconds 3
        }
        & nssm remove $ServiceName confirm
        Start-Sleep -Seconds 2
    } else {
        Write-Error "Service '$ServiceName' already exists. Use -Force to replace it."
        exit 1
    }
}

# Test nginx configuration before installing service
Write-Host "Testing nginx configuration..." -ForegroundColor Yellow
Push-Location $NginxPath
try {
    $testResult = & .\nginx.exe -t 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Error "nginx configuration test failed: $testResult"
        Write-Host "Please fix the nginx configuration before creating the service." -ForegroundColor Yellow
        exit 1
    } else {
        Write-Host "✓ nginx configuration is valid" -ForegroundColor Green
    }
} finally {
    Pop-Location
}

# Install nginx service with NSSM
Write-Host "Installing nginx as Windows service..." -ForegroundColor Yellow

# Create service wrapper script to handle nginx properly
$wrapperScript = @"
@echo off
cd /d "$NginxPath"
nginx.exe -c "$NginxPath\conf\nginx.conf"
"@

$wrapperPath = "$NginxPath\nginx-service-wrapper.bat"
$wrapperScript | Set-Content -Path $wrapperPath -Encoding ASCII

# Install service
& nssm install $ServiceName "$wrapperPath"
if ($LASTEXITCODE -ne 0) {
    Write-Error "Failed to install nginx service"
    exit 1
}

# Configure service settings
Write-Host "Configuring service settings..." -ForegroundColor Yellow

# Basic service configuration
& nssm set $ServiceName DisplayName "nginx Web Server"
& nssm set $ServiceName Description "nginx HTTP and reverse proxy server for Excel Add-in"
& nssm set $ServiceName Start SERVICE_AUTO_START

# Set working directory
& nssm set $ServiceName AppDirectory "$NginxPath"

# Configure logging
$logDir = "C:\Logs\nginx"
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

& nssm set $ServiceName AppStdout "$logDir\service-stdout.log"
& nssm set $ServiceName AppStderr "$logDir\service-stderr.log"

# Windows-specific optimizations for nginx service
& nssm set $ServiceName AppPriority NORMAL_PRIORITY_CLASS
& nssm set $ServiceName AppNoConsole 1

# Set service dependencies (wait for network)
& nssm set $ServiceName DependOnService Tcpip

# Configure service recovery options
& nssm set $ServiceName AppThrottle 1500
& nssm set $ServiceName AppExit Default Restart
& nssm set $ServiceName AppRestartDelay 5000

# Configure service shutdown
& nssm set $ServiceName AppStopMethodSkip 0
& nssm set $ServiceName AppStopMethodConsole 10000
& nssm set $ServiceName AppStopMethodWindow 10000
& nssm set $ServiceName AppStopMethodThreads 10000

Write-Host "✓ nginx service configured successfully" -ForegroundColor Green

# Create nginx management scripts
Write-Host "Creating management scripts..." -ForegroundColor Yellow

# nginx reload script
$reloadScript = @"
@echo off
echo Reloading nginx configuration...
cd /d "$NginxPath"
nginx.exe -s reload
if %errorlevel% equ 0 (
    echo nginx configuration reloaded successfully
) else (
    echo Failed to reload nginx configuration
)
pause
"@
$reloadScript | Set-Content -Path "$NginxPath\reload-nginx.bat" -Encoding ASCII

# nginx stop script
$stopScript = @"
@echo off
echo Stopping nginx gracefully...
cd /d "$NginxPath"
nginx.exe -s stop
if %errorlevel% equ 0 (
    echo nginx stopped successfully
) else (
    echo Failed to stop nginx gracefully, stopping service...
    net stop nginx
)
pause
"@
$stopScript | Set-Content -Path "$NginxPath\stop-nginx.bat" -Encoding ASCII

# nginx test configuration script
$testScript = @"
@echo off
echo Testing nginx configuration...
cd /d "$NginxPath"
nginx.exe -t
if %errorlevel% equ 0 (
    echo Configuration is valid
) else (
    echo Configuration has errors
)
pause
"@
$testScript | Set-Content -Path "$NginxPath\test-config.bat" -Encoding ASCII

Write-Host "✓ Management scripts created in $NginxPath" -ForegroundColor Green

# Start the service
Write-Host "Starting nginx service..." -ForegroundColor Yellow
Start-Service -Name $ServiceName

# Wait and check service status
Start-Sleep -Seconds 5
$service = Get-Service -Name $ServiceName
if ($service.Status -eq 'Running') {
    Write-Host "✓ nginx service is running successfully" -ForegroundColor Green
} else {
    Write-Warning "nginx service status: $($service.Status)"
    Write-Host "Check service logs at: $logDir" -ForegroundColor Yellow
}

# Test HTTP response
Write-Host "Testing HTTP response..." -ForegroundColor Yellow
Start-Sleep -Seconds 2
try {
    $response = Invoke-WebRequest -Uri "http://127.0.0.1/health" -TimeoutSec 10 -ErrorAction SilentlyContinue
    if ($response.StatusCode -eq 200) {
        Write-Host "✓ nginx is responding to HTTP requests" -ForegroundColor Green
    } else {
        Write-Warning "nginx returned status: $($response.StatusCode)"
    }
} catch {
    Write-Warning "Could not connect to nginx on HTTP (this is normal if only HTTPS is configured)"
}

Write-Host "`n" + "=" * 40 -ForegroundColor Cyan
Write-Host "SERVICE SETUP COMPLETE" -ForegroundColor Cyan
Write-Host "=" * 40 -ForegroundColor Cyan

Write-Host "`nService Management:" -ForegroundColor Green
Write-Host "• Start:   Start-Service $ServiceName"
Write-Host "• Stop:    Stop-Service $ServiceName"
Write-Host "• Restart: Restart-Service $ServiceName"
Write-Host "• Status:  Get-Service $ServiceName"

Write-Host "`nManagement Scripts in ${NginxPath}:" -ForegroundColor Green
Write-Host "• reload-nginx.bat  - Reload configuration"
Write-Host "• stop-nginx.bat    - Stop nginx gracefully"
Write-Host "• test-config.bat   - Test configuration syntax"

Write-Host "`nnginx Configuration:" -ForegroundColor Green
Write-Host "• Config file: $NginxPath\conf\nginx.conf"
Write-Host "• Test config: nginx -t (from $NginxPath)"
Write-Host "• Reload config: nginx -s reload (from $NginxPath)"

Write-Host "`nLogs:" -ForegroundColor Green
Write-Host "• Service logs: $logDir\"
Write-Host "• nginx logs: $NginxPath\logs\"
Write-Host "• Application logs: C:\Logs\nginx\excel_addin_*.log"

Write-Host "`nTroubleshooting:" -ForegroundColor Yellow
Write-Host "• If service fails to start, check logs in $logDir"
Write-Host "• Ensure nginx configuration is valid (nginx -t)"
Write-Host "• Check Windows Event Viewer for service errors"
Write-Host "• Use 'nssm edit $ServiceName' to modify service settings"

Write-Host "`nIMPORTANT NOTES:" -ForegroundColor Yellow
Write-Host "• This service will start automatically with Windows"
Write-Host "• Configuration changes require service restart or reload"
Write-Host "• Monitor service logs regularly for issues"
Write-Host "• Use proper SSL certificates for production"