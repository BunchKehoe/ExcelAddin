# ExcelAddin Frontend Deployment Script  
# Deploys Vite-built React frontend as Windows Service using node-windows

param(
    [switch]$Force,
    [switch]$SkipBuild,
    [switch]$Debug,
    [string]$Environment = "staging"
)

$ErrorActionPreference = "Stop"

# Service configuration
$ServiceName = "ExcelAddin Frontend"
$Port = 3000

# Paths
$ProjectRoot = Split-Path $PSScriptRoot -Parent
$DistPath = Join-Path $ProjectRoot "dist"
$LogDir = "C:\Logs\ExcelAddin"
$NodeCmd = Get-Command node -ErrorAction SilentlyContinue
if ($NodeCmd) {
    $NodeExe = $NodeCmd.Source
} else {
    Write-Error "Node.js not found. Please install Node.js and add to PATH."
}

Write-Host "========================================" -ForegroundColor Green
Write-Host "  ExcelAddin Frontend Deployment (node-windows)" -ForegroundColor Green  
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

# Validate environment parameter
if ($Environment -notin @("development", "staging", "production")) {
    Write-Error "Invalid environment '$Environment'. Must be one of: development, staging, production"
}

Write-Host "Deployment Configuration:" -ForegroundColor Cyan
Write-Host "  Environment: $Environment" -ForegroundColor Cyan
Write-Host "  Port: $Port" -ForegroundColor Cyan
Write-Host "  Service: $ServiceName" -ForegroundColor Cyan
Write-Host ""

# Check prerequisites  
Write-Host "Checking prerequisites..." -ForegroundColor Yellow

# Check Node.js
if (-not $NodeExe) {
    Write-Error "Node.js not found. Please install Node.js 18+ and add to PATH."
}
$nodeVersion = node --version
Write-Host "  Node.js: $nodeVersion" -ForegroundColor Green

# Check NPM
$npmCmd = Get-Command npm -ErrorAction SilentlyContinue
if ($npmCmd) {
    $npmPath = $npmCmd.Source
} else {
    Write-Error "npm not found. Please ensure npm is installed with Node.js."
}
Write-Host "  npm: Found" -ForegroundColor Green

Write-Host "  Prerequisites: OK" -ForegroundColor Green
Write-Host ""

# Verify paths
Write-Host "Verifying project structure..." -ForegroundColor Yellow
Write-Host "  Project Root: $ProjectRoot"

if (-not (Test-Path $ProjectRoot)) {
    Write-Error "Project root directory not found: $ProjectRoot"
}

if (-not (Test-Path (Join-Path $ProjectRoot "package.json"))) {
    Write-Error "package.json not found. Are you in the correct directory?"
}

if (-not (Test-Path (Join-Path $ProjectRoot "service.cjs"))) {
    Write-Error "service.cjs not found. Service wrapper script is missing."
}

if (-not (Test-Path (Join-Path $ProjectRoot "server.cjs"))) {
    Write-Error "server.cjs not found. Express server script is missing."
}

# Create log directory
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
    Write-Host "  Created log directory: $LogDir" -ForegroundColor Green
} else {
    Write-Host "  Log directory exists: $LogDir" -ForegroundColor Green
}

# Navigate to project root
Set-Location $ProjectRoot

# Install dependencies and build application
if (-not $SkipBuild) {
    Write-Host "Installing dependencies..." -ForegroundColor Yellow
    npm install
    if ($LASTEXITCODE -ne 0) {
        Write-Error "npm install failed"
    }
    Write-Host "  Dependencies installed successfully" -ForegroundColor Green
    
    Write-Host "Building application for $Environment..." -ForegroundColor Yellow
    
    $buildCommand = switch ($Environment) {
        "development" { "npm run build:dev" }
        "staging" { "npm run build:staging" }
        "production" { "npm run build:prod" }
    }
    
    Invoke-Expression $buildCommand
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Build failed"
    }
    Write-Host "  Build completed successfully" -ForegroundColor Green
    
    # Verify build output
    if (-not (Test-Path $DistPath)) {
        Write-Error "Build output directory not found: $DistPath"
    }
    
    $requiredFiles = @("taskpane.html", "commands.html", "functions.json")
    foreach ($file in $requiredFiles) {
        $filePath = Join-Path $DistPath $file
        if (-not (Test-Path $filePath)) {
            Write-Error "Required file not found in build output: $file"
        }
    }
    
    Write-Host "  Build verification: PASSED" -ForegroundColor Green
    Write-Host ""
}

# Check if port is in use (skip our own service)
$existingProcess = Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue | 
    Where-Object { $_.State -eq "Listen" }
if ($existingProcess -and -not $Force) {
    $processId = $existingProcess.OwningProcess
    $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
    if ($process) {
        $processName = $process.ProcessName
    } else {
        $processName = "Unknown"
    }
    Write-Warning "Port $Port is in use by process: $processName (PID: $processId)"
    Write-Warning "Use -Force to override or stop the conflicting process first"
    exit 1
}

# Stop and uninstall existing service if it exists
Write-Host "Checking for existing Windows service..." -ForegroundColor Yellow
$existingService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($existingService) {
    Write-Host "Stopping existing service..." -ForegroundColor Yellow
    
    # Stop the service first
    if ($existingService.Status -eq "Running") {
        Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
        
        # Wait for service to stop
        $timeout = 30
        $elapsed = 0
        do {
            Start-Sleep -Seconds 2
            $elapsed += 2
            $serviceStatus = (Get-Service -Name $ServiceName -ErrorAction SilentlyContinue).Status
        } while ($serviceStatus -eq "Running" -and $elapsed -lt $timeout)
        
        if ($serviceStatus -eq "Running") {
            Write-Warning "Service did not stop within $timeout seconds, continuing..."
        } else {
            Write-Host "Service stopped successfully" -ForegroundColor Green
        }
    }
    
    # Uninstall the service using node-windows
    Write-Host "Uninstalling existing service..." -ForegroundColor Yellow
    $env:NODE_ENV = if ($Environment -eq "development") { "development" } else { "production" }
    $env:PORT = $Port
    $env:HOST = "127.0.0.1"
    
    node service.cjs uninstall
    
    # Wait for uninstall to complete
    Start-Sleep -Seconds 5
    
    # Verify removal
    $serviceCheck = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($serviceCheck) {
        Write-Warning "Service may still exist, continuing with installation..."
    } else {
        Write-Host "Service uninstalled successfully" -ForegroundColor Green
    }
}

# Kill any existing node processes on our port to ensure clean start
$existingProcess = Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue | 
    Where-Object { $_.State -eq "Listen" }
if ($existingProcess) {
    $processId = $existingProcess.OwningProcess
    $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
    if ($process -and $process.ProcessName -eq "node") {
        Write-Host "Stopping existing Node.js process on port $Port..." -ForegroundColor Yellow
        Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3
    }
}

# Install the Windows service using node-windows
Write-Host "Installing Windows service using node-windows..." -ForegroundColor Yellow

# Set environment variables for the service
$env:NODE_ENV = if ($Environment -eq "development") { "development" } else { "production" }
$env:PORT = $Port
$env:HOST = "127.0.0.1"

# Install the service
node service.cjs install

if ($LASTEXITCODE -ne 0) {
    Write-Error "Failed to install Windows service"
}

# Wait for installation to complete and service to start
Write-Host "Waiting for service installation and startup..." -ForegroundColor Yellow
$timeout = 60
$elapsed = 0
$serviceStarted = $false

do {
    Start-Sleep -Seconds 3
    $elapsed += 3
    
    # Check if service exists and is running
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($service -and $service.Status -eq "Running") {
        # Also check if Node.js process is running on our port
        $runningProcess = Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue | 
            Where-Object { $_.State -eq "Listen" }
        
        if ($runningProcess) {
            $processId = $runningProcess.OwningProcess
            $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
            if ($process -and $process.ProcessName -eq "node") {
                $serviceStarted = $true
            }
        }
    }
} while (-not $serviceStarted -and $elapsed -lt $timeout)

if (-not $serviceStarted) {
    Write-Error "Service failed to start within $timeout seconds. Check Windows Event Log and service status."
    
    # Show service status for debugging
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($service) {
        Write-Host "Service Status: $($service.Status)" -ForegroundColor Red
        Write-Host "Check Windows Event Viewer for detailed error messages" -ForegroundColor Yellow
    } else {
        Write-Host "Service not found after installation" -ForegroundColor Red
    }
    exit 1
}

Write-Host "Windows service installed and started successfully!" -ForegroundColor Green
$service = Get-Service -Name $ServiceName
Write-Host "  Service Status: $($service.Status)" -ForegroundColor Green
Write-Host "  Process running on port: $Port" -ForegroundColor Green

# Test connectivity
Write-Host "Testing frontend connectivity..." -ForegroundColor Yellow
Start-Sleep -Seconds 5

$testEndpoints = @(
    @{ Path = "/health"; Name = "Health Check" },
    @{ Path = "/excellence/taskpane.html"; Name = "Taskpane HTML" },
    @{ Path = "/excellence/commands.html"; Name = "Commands HTML" },
    @{ Path = "/functions.json"; Name = "Functions Manifest" }
)

foreach ($endpoint in $testEndpoints) {
    try {
        $uri = "http://localhost:$Port$($endpoint.Path)"
        $response = Invoke-WebRequest -Uri $uri -TimeoutSec 10 -ErrorAction SilentlyContinue
        if ($response.StatusCode -eq 200) {
            Write-Host "  $($endpoint.Name): PASSED" -ForegroundColor Green
        } else {
            Write-Warning "  $($endpoint.Name): HTTP $($response.StatusCode)"
        }
    } catch {
        Write-Warning "  $($endpoint.Name): FAILED - $($_.Exception.Message)"
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green  
Write-Host "  Frontend Deployment Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Service Details:" -ForegroundColor Cyan
Write-Host "  Name: $ServiceName"
Write-Host "  Port: $Port"  
Write-Host "  Environment: $Environment"
Write-Host "  Status: $($service.Status)"
Write-Host ""
Write-Host "Access URLs:" -ForegroundColor Cyan
Write-Host "  Health Check:  http://localhost:$Port/health"
Write-Host "  Taskpane:      http://localhost:$Port/excellence/taskpane.html"
Write-Host "  Commands:      http://localhost:$Port/excellence/commands.html"
Write-Host "  Functions:     http://localhost:$Port/functions.json"
Write-Host ""
Write-Host "Service Management:" -ForegroundColor Cyan
Write-Host "  Windows Services:  services.msc -> '$ServiceName'"
Write-Host "  Start Service:     Start-Service '$ServiceName'"
Write-Host "  Stop Service:      Stop-Service '$ServiceName'"
Write-Host "  Service Status:    Get-Service '$ServiceName'"
Write-Host "  Restart:           Restart-Service '$ServiceName'"
Write-Host ""
Write-Host "Advanced Management:" -ForegroundColor Cyan
Write-Host "  Install:           node service.cjs install"
Write-Host "  Uninstall:         node service.cjs uninstall" 
Write-Host "  Manual Start:      node service.cjs start"
Write-Host "  Manual Stop:       node service.cjs stop"
Write-Host "  Logs:              Windows Event Viewer -> Application Log"
Write-Host ""