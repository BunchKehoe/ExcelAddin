# Simple Excel Add-in Frontend Deployment Script
# Compatible with Windows 10 Server and older IIS versions

param(
    [switch]$SkipBuild,
    [int]$Port = 3000,
    [switch]$ConfigureIIS
)

# Configuration
$ServiceName = "ExcelAddin-Frontend"
$ServiceDisplayName = "ExcelAddin Frontend Service"
$SiteName = "ExcelAddin"

# Get paths
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir
$AppPath = Join-Path $ProjectRoot "dist"
$ServerScript = Join-Path $ScriptDir "config\frontend-server.js"
$NodeExe = "C:\Program Files\nodejs\node.exe"
$LogDir = "C:\Logs\ExcelAddin"

Write-Host "===================================================="
Write-Host "Excel Add-in Frontend Deployment"
Write-Host "===================================================="
Write-Host "Project Root: $ProjectRoot"
Write-Host "App Path: $AppPath"
Write-Host ""

# Check prerequisites
Write-Host "Checking prerequisites..."

# Check Node.js
if (!(Test-Path $NodeExe)) {
    Write-Error "Node.js not found at $NodeExe. Please install Node.js."
    exit 1
}

# Check NSSM
$nssmPath = Get-Command nssm -ErrorAction SilentlyContinue
if (!$nssmPath) {
    Write-Error "NSSM not found. Please install NSSM."
    exit 1
}

Write-Host "Prerequisites OK"
Write-Host ""

# Navigate to project root
Set-Location $ProjectRoot

# Clean previous build
if (Test-Path "dist") {
    Write-Host "Removing previous build..."
    Remove-Item "dist" -Recurse -Force
}

# Install dependencies
Write-Host "Installing dependencies..."
npm install
if ($LASTEXITCODE -ne 0) {
    Write-Error "npm install failed"
    exit 1
}
Write-Host "Dependencies installed successfully"
Write-Host ""

# Build application
if (!$SkipBuild) {
    Write-Host "Building application..."
    npm run build:staging
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Build failed"
        exit 1
    }
    Write-Host "Build completed successfully"
    Write-Host ""
}

# Verify build output
if (!(Test-Path $AppPath)) {
    Write-Error "Build output directory not found: $AppPath"
    exit 1
}

if (!(Test-Path (Join-Path $AppPath "index.html"))) {
    Write-Error "index.html not found in build output"
    exit 1
}

Write-Host "Build verification passed"
Write-Host ""

# Create log directory
if (!(Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
}

# Stop and remove existing service
$existingService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($existingService) {
    Write-Host "Stopping existing service..."
    Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 5
    
    Write-Host "Removing existing service..."
    nssm remove $ServiceName confirm
    Start-Sleep -Seconds 3
}

# Install new service
Write-Host "Installing NSSM service..."
nssm install $ServiceName $NodeExe
nssm set $ServiceName AppParameters $ServerScript
nssm set $ServiceName AppDirectory $ProjectRoot
nssm set $ServiceName DisplayName $ServiceDisplayName
nssm set $ServiceName Description "Excel Add-in Frontend Web Server"
nssm set $ServiceName Start SERVICE_AUTO_START

# Set environment variables
nssm set $ServiceName AppEnvironmentExtra "NODE_ENV=production;PORT=$Port;HOST=127.0.0.1"

# Set logging
nssm set $ServiceName AppStdout "$LogDir\frontend-stdout.log"
nssm set $ServiceName AppStderr "$LogDir\frontend-stderr.log"
nssm set $ServiceName AppRotateFiles 1
nssm set $ServiceName AppRotateOnline 1
nssm set $ServiceName AppRotateSeconds 86400

Write-Host "NSSM service configured"
Write-Host ""

# Start service
Write-Host "Starting service..."
nssm start $ServiceName
Start-Sleep -Seconds 10

# Check service status
$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($service -and $service.Status -eq "Running") {
    Write-Host "Service started successfully: $($service.Status)"
} else {
    Write-Error "Service failed to start. Status: $($service.Status)"
    Write-Host "Check logs at: $LogDir"
    exit 1
}

# Test the service
Write-Host "Testing service response..."
try {
    $response = Invoke-WebRequest -Uri "http://127.0.0.1:$Port" -UseBasicParsing -TimeoutSec 10
    if ($response.StatusCode -eq 200) {
        Write-Host "Service is responding correctly (HTTP 200)"
    } else {
        Write-Warning "Service responded with HTTP $($response.StatusCode)"
    }
} catch {
    Write-Error "Failed to connect to service: $($_.Exception.Message)"
    Write-Host "Check logs at: $LogDir"
    exit 1
}

Write-Host ""

# Configure IIS if requested
if ($ConfigureIIS) {
    Write-Host "Configuring IIS..."
    
    # Import WebAdministration module
    Import-Module WebAdministration -ErrorAction SilentlyContinue
    
    # Remove existing site if it exists
    if (Get-Website -Name $SiteName -ErrorAction SilentlyContinue) {
        Write-Host "Removing existing IIS site..."
        Remove-Website -Name $SiteName
    }
    
    # Create new website
    Write-Host "Creating IIS site: $SiteName"
    New-Website -Name $SiteName -PhysicalPath $AppPath
    
    # Configure web.config for SPA
    $webConfigPath = Join-Path $AppPath "web.config"
    $webConfigContent = @'
<?xml version="1.0" encoding="UTF-8"?>
<configuration>
    <system.webServer>
        <rewrite>
            <rules>
                <rule name="SPA" stopProcessing="true">
                    <match url=".*" />
                    <conditions logicalGrouping="MatchAll">
                        <add input="{REQUEST_FILENAME}" matchType="IsFile" negate="true" />
                    </conditions>
                    <action type="Rewrite" url="/" />
                </rule>
            </rules>
        </rewrite>
        <staticContent>
            <mimeMap fileExtension=".json" mimeType="application/json" />
            <mimeMap fileExtension=".js" mimeType="application/javascript" />
        </staticContent>
    </system.webServer>
</configuration>
'@
    
    Set-Content -Path $webConfigPath -Value $webConfigContent
    Write-Host "Created web.config"
    
    # Start the website
    Start-Website -Name $SiteName
    Write-Host "IIS site configured and started"
}

Write-Host ""
Write-Host "===================================================="
Write-Host "Deployment completed successfully!"
Write-Host "===================================================="
Write-Host "Service: $ServiceName"
Write-Host "Status: Running"
Write-Host "URL: http://127.0.0.1:$Port"
Write-Host "Logs: $LogDir"
if ($ConfigureIIS) {
    Write-Host "IIS Site: $SiteName"
}
Write-Host ""
Write-Host "To troubleshoot issues:"
Write-Host "1. Check service status: Get-Service $ServiceName"
Write-Host "2. View logs: Get-Content $LogDir\frontend-stdout.log -Tail 20"
Write-Host "3. Test manually: & '$NodeExe' '$ServerScript'"
Write-Host "===================================================="