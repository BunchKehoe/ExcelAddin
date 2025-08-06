# Excel Add-in Windows Server Deployment Script (PowerShell)
# This script provides more advanced deployment functionality
# Run as Administrator in PowerShell

param(
    [string]$DomainName = "your-staging-domain.com",
    [string]$DeployDir = "C:\inetpub\wwwroot\ExcelAddin",
    [string]$LogDir = "C:\Logs\ExcelAddin",
    [string]$NginxDir = "C:\nginx",
    [string]$ServiceName = "ExcelAddinBackend",
    [switch]$SkipServiceInstall,
    [switch]$UpdateOnly
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

Write-Host "================================================================" -ForegroundColor Green
Write-Host "Excel Add-in Windows Server Deployment" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Green

Write-Host "Configuration:" -ForegroundColor Yellow
Write-Host "- Domain: $DomainName"
Write-Host "- Deploy Directory: $DeployDir"
Write-Host "- Log Directory: $LogDir"
Write-Host "- nginx Directory: $NginxDir"
Write-Host "- Service Name: $ServiceName"
Write-Host ""

# Create directories
Write-Host "Creating directories..." -ForegroundColor Yellow
$directories = @($DeployDir, $LogDir, "$NginxDir\logs", "C:\Logs\nginx")
foreach ($dir in $directories) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Write-Host "Created: $dir" -ForegroundColor Green
    }
}

# Stop existing services
if (-not $UpdateOnly) {
    Write-Host "Stopping existing services..." -ForegroundColor Yellow
    
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($service -and $service.Status -eq 'Running') {
        Write-Host "Stopping $ServiceName service..."
        Stop-Service -Name $ServiceName -Force
        Start-Sleep -Seconds 3
    }
    
    # Stop nginx
    $nginxProcesses = Get-Process -Name "nginx" -ErrorAction SilentlyContinue
    if ($nginxProcesses) {
        Write-Host "Stopping nginx..."
        Stop-Process -Name "nginx" -Force
        Start-Sleep -Seconds 2
    }
}

# Build frontend
Write-Host "Building frontend..." -ForegroundColor Yellow
if (Test-Path "package.json") {
    npm run build
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Frontend build failed"
        exit 1
    }
} else {
    Write-Warning "package.json not found. Make sure to build the frontend first."
}

# Copy application files
Write-Host "Copying application files..." -ForegroundColor Yellow
if (Test-Path "dist") {
    Copy-Item -Path "dist\*" -Destination "$DeployDir\dist\" -Recurse -Force
    Write-Host "Frontend files copied" -ForegroundColor Green
}

if (Test-Path "backend") {
    Copy-Item -Path "backend\*" -Destination "$DeployDir\backend\" -Recurse -Force
    Write-Host "Backend files copied" -ForegroundColor Green
}

if (Test-Path "manifest-staging.xml") {
    Copy-Item -Path "manifest-staging.xml" -Destination "$DeployDir\" -Force
    Write-Host "Manifest file copied" -ForegroundColor Green
}

# Copy nginx configuration
Write-Host "Setting up nginx configuration..." -ForegroundColor Yellow
$nginxConfDir = "$NginxDir\conf\conf.d"
if (-not (Test-Path $nginxConfDir)) {
    New-Item -ItemType Directory -Path $nginxConfDir -Force | Out-Null
}

$nginxConfigSource = "deployment\nginx\excel-addin.conf"
$nginxConfigDest = "$nginxConfDir\excel-addin.conf"

if (Test-Path $nginxConfigSource) {
    # Copy and update nginx configuration
    $nginxConfig = Get-Content $nginxConfigSource -Raw
    $nginxConfig = $nginxConfig -replace 'your-staging-domain\.com', $DomainName
    $nginxConfig = $nginxConfig -replace 'C:/inetpub/wwwroot/ExcelAddin', $DeployDir.Replace('\', '/')
    $nginxConfig | Set-Content -Path $nginxConfigDest
    Write-Host "nginx configuration updated" -ForegroundColor Green
} else {
    Write-Warning "nginx configuration template not found at $nginxConfigSource"
}

# Set up backend environment
Write-Host "Setting up backend environment..." -ForegroundColor Yellow
$backendDir = "$DeployDir\backend"

if (Test-Path "$backendDir\.env.production") {
    Copy-Item -Path "$backendDir\.env.production" -Destination "$backendDir\.env" -Force
    
    # Update environment file
    $envContent = Get-Content "$backendDir\.env" -Raw
    $envContent = $envContent -replace 'your-staging-domain\.com', $DomainName
    $envContent = $envContent -replace 'C:\\inetpub\\wwwroot\\ExcelAddin\\backend', $DeployDir.Replace('/', '\') + '\backend'
    $envContent | Set-Content -Path "$backendDir\.env"
    Write-Host "Backend environment configured" -ForegroundColor Green
}

# Install Python dependencies
Write-Host "Installing Python dependencies..." -ForegroundColor Yellow
Push-Location $backendDir
try {
    python -m pip install -r requirements.txt --user
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Python dependencies installed" -ForegroundColor Green
    } else {
        Write-Warning "Failed to install some Python dependencies"
    }
} catch {
    Write-Warning "Error installing Python dependencies: $_"
}
Pop-Location

# Install Windows Service
if (-not $SkipServiceInstall -and -not $UpdateOnly) {
    Write-Host "Installing Windows Service..." -ForegroundColor Yellow
    
    # Check for NSSM
    try {
        $nssmVersion = nssm version 2>$null
        Write-Host "NSSM found: $nssmVersion"
        
        # Remove existing service if it exists
        $existingService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if ($existingService) {
            nssm remove $ServiceName confirm
        }
        
        # Install new service
        nssm install $ServiceName python "$DeployDir\backend\service_wrapper.py"
        nssm set $ServiceName AppDirectory "$DeployDir\backend"
        nssm set $ServiceName DisplayName "Excel Add-in Backend Service"
        nssm set $ServiceName Description "Python Flask backend service for Excel Add-in"
        nssm set $ServiceName Start SERVICE_AUTO_START
        nssm set $ServiceName AppStdout "$LogDir\service-stdout.log"
        nssm set $ServiceName AppStderr "$LogDir\service-stderr.log"
        
        # Set environment variables
        $envVars = "FLASK_ENV=production", "DEBUG=false", "HOST=127.0.0.1", "PORT=5000", "PYTHONPATH=$DeployDir\backend"
        nssm set $ServiceName AppEnvironmentExtra $envVars
        
        Write-Host "Windows Service installed" -ForegroundColor Green
    } catch {
        Write-Warning "NSSM not found. Please install NSSM from https://nssm.cc/"
        Write-Warning "Service installation skipped."
    }
}

# Set file permissions
Write-Host "Setting file permissions..." -ForegroundColor Yellow
icacls $DeployDir /grant "IIS_IUSRS:(OI)(CI)R" /T | Out-Null
icacls $LogDir /grant "IIS_IUSRS:(OI)(CI)F" /T | Out-Null
Write-Host "File permissions set" -ForegroundColor Green

# Create error pages
Write-Host "Creating error pages..." -ForegroundColor Yellow
@"
<!DOCTYPE html>
<html>
<head>
    <title>Page Not Found</title>
    <style>body { font-family: Arial, sans-serif; text-align: center; margin-top: 50px; }</style>
</head>
<body>
    <h1>404 - Page Not Found</h1>
    <p>The requested page could not be found.</p>
</body>
</html>
"@ | Set-Content -Path "$DeployDir\dist\404.html"

@"
<!DOCTYPE html>
<html>
<head>
    <title>Server Error</title>
    <style>body { font-family: Arial, sans-serif; text-align: center; margin-top: 50px; }</style>
</head>
<body>
    <h1>Server Error</h1>
    <p>The server encountered an error processing your request.</p>
</body>
</html>
"@ | Set-Content -Path "$DeployDir\dist\50x.html"

# Start services
if (-not $UpdateOnly) {
    Write-Host "Starting services..." -ForegroundColor Yellow
    
    # Start backend service
    try {
        Start-Service -Name $ServiceName
        Write-Host "Backend service started" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to start backend service: $_"
    }
    
    # Start nginx
    Push-Location $NginxDir
    try {
        Start-Process -FilePath "nginx.exe" -WindowStyle Hidden
        Start-Sleep -Seconds 3
        $nginxProcesses = Get-Process -Name "nginx" -ErrorAction SilentlyContinue
        if ($nginxProcesses) {
            Write-Host "nginx started" -ForegroundColor Green
        } else {
            Write-Warning "nginx may not have started properly"
        }
    } catch {
        Write-Warning "Failed to start nginx: $_"
    }
    Pop-Location
}

# Health check
Write-Host "Performing health checks..." -ForegroundColor Yellow
Start-Sleep -Seconds 5

$backendService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($backendService -and $backendService.Status -eq 'Running') {
    Write-Host "✓ Backend service is running" -ForegroundColor Green
} else {
    Write-Host "✗ Backend service is not running" -ForegroundColor Red
}

$nginxProcesses = Get-Process -Name "nginx" -ErrorAction SilentlyContinue
if ($nginxProcesses) {
    Write-Host "✓ nginx is running" -ForegroundColor Green
} else {
    Write-Host "✗ nginx is not running" -ForegroundColor Red
}

# Test backend endpoint
try {
    $response = Invoke-WebRequest -Uri "http://127.0.0.1:5000/api/health" -TimeoutSec 10 -ErrorAction SilentlyContinue
    if ($response.StatusCode -eq 200) {
        Write-Host "✓ Backend API is responding" -ForegroundColor Green
    } else {
        Write-Host "✗ Backend API returned status: $($response.StatusCode)" -ForegroundColor Red
    }
} catch {
    Write-Host "✗ Backend API is not responding" -ForegroundColor Red
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Green
Write-Host "Deployment Complete!" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Next Steps:" -ForegroundColor Yellow
Write-Host "1. Update DNS to point $DomainName to this server"
Write-Host "2. Install SSL certificate for $DomainName"
Write-Host "3. Update nginx SSL certificate paths in $nginxConfDir\excel-addin.conf"
Write-Host "4. Test the application at https://$DomainName"
Write-Host "5. Distribute the manifest file to developers"
Write-Host ""
Write-Host "Management Commands:" -ForegroundColor Cyan
Write-Host "- Start Backend: Start-Service $ServiceName"
Write-Host "- Stop Backend: Stop-Service $ServiceName"
Write-Host "- Restart nginx: nginx -s reload (from $NginxDir)"
Write-Host "- View logs: Get-Content $LogDir\*.log -Wait"
Write-Host ""
Write-Host "URLs:" -ForegroundColor Cyan
Write-Host "- Application: https://$DomainName"
Write-Host "- API Health: https://$DomainName/api/health"
Write-Host "- nginx Status: https://$DomainName/health"