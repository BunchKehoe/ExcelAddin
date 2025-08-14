# ExcelAddin Frontend Deployment Script  
# Deploys Vite-built React frontend as NSSM service for Windows Server 10

param(
    [switch]$Force,
    [switch]$SkipBuild,
    [switch]$Debug,
    [string]$Environment = "staging"
)

$ErrorActionPreference = "Stop"

# Service configuration
$ServiceName = "ExcelAddin-Frontend"
$ServiceDisplayName = "ExcelAddin Frontend Service"
$ServiceDescription = "Excel Add-in Frontend Web Server (Vite)"
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

# Server script content for NSSM (Windows Server 10 compatible)
$ServerScript = @"
const express = require('express');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3000;
const HOST = process.env.HOST || '127.0.0.1';

// Serve static files from dist directory
app.use('/excellence', express.static(path.join(__dirname, 'dist')));

// Serve manifest files and assets at root level too (for backward compatibility)
app.use('/assets', express.static(path.join(__dirname, 'dist/assets')));
app.get('/manifest*.xml', (req, res) => {
    const manifestPath = path.join(__dirname, 'dist', req.path);
    if (fs.existsSync(manifestPath)) {
        res.setHeader('Content-Type', 'application/xml');
        res.sendFile(manifestPath);
    } else {
        res.status(404).send('Manifest not found');
    }
});

// Serve functions.json
app.get('/functions.json', (req, res) => {
    res.sendFile(path.join(__dirname, 'dist/functions.json'));
});

// Excel Add-in specific endpoints (required by Excel)
app.get('/excellence/taskpane.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'dist/taskpane.html'));
});

app.get('/excellence/commands.html', (req, res) => {
    res.sendFile(path.join(__dirname, 'dist/commands.html'));
});

// Health check endpoint
app.get('/health', (req, res) => {
    res.json({ 
        status: 'healthy', 
        service: 'exceladdin-frontend',
        timestamp: new Date().toISOString(),
        environment: process.env.NODE_ENV || 'production'
    });
});

// Fallback for any other Excel Add-in requests
app.get('/excellence/*', (req, res) => {
    const filePath = path.join(__dirname, 'dist', req.path.replace('/excellence/', ''));
    if (fs.existsSync(filePath)) {
        res.sendFile(filePath);
    } else {
        // For SPA routing, serve taskpane.html as fallback
        res.sendFile(path.join(__dirname, 'dist/taskpane.html'));
    }
});

// CORS headers for Excel Add-in compatibility
app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, PATCH, OPTIONS');
    res.header('Access-Control-Allow-Headers', 'X-Requested-With, content-type, Authorization');
    next();
});

// Start server
const server = app.listen(PORT, HOST, () => {
    console.log(`Excel Add-in Frontend Server running on http://$${HOST}:$${PORT}`);
    console.log(`Environment: $${process.env.NODE_ENV || 'production'}`);
    console.log(`Serving from: $${path.join(__dirname, 'dist')}`);
});

// Graceful shutdown
process.on('SIGTERM', () => {
    console.log('SIGTERM received, shutting down gracefully');
    server.close(() => {
        console.log('Frontend server stopped');
        process.exit(0);
    });
});

process.on('SIGINT', () => {
    console.log('SIGINT received, shutting down gracefully');
    server.close(() => {
        console.log('Frontend server stopped');
        process.exit(0);
    });
});
"@

Write-Host "========================================" -ForegroundColor Green
Write-Host "  ExcelAddin Frontend Deployment (NSSM)" -ForegroundColor Green  
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

# Validate environment parameter
if ($Environment -notin @("development", "staging", "production")) {
    Write-Error "Invalid environment '$Environment'. Must be one of: development, staging, production"
}

Write-Host "Deployment Configuration:" -ForegroundColor Cyan
Write-Host "  Environment: $Environment" -ForegroundColor Cyan
Write-Host "  Port: $Port" -ForegroundColor Cyan
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

# Check NSSM
$nssmCmd = Get-Command nssm -ErrorAction SilentlyContinue
if ($nssmCmd) {
    $nssmPath = $nssmCmd.Source
} else {
    Write-Error "NSSM not found. Please install NSSM and add to PATH."
}
Write-Host "  NSSM: Found at $nssmPath" -ForegroundColor Green

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
if ($existingProcess) {
    $processId = $existingProcess.OwningProcess
    $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
    if ($process) {
        $processName = $process.ProcessName
    } else {
        $processName = "Unknown"
    }
    if ($processName -ne "node" -and -not $Force) {
        Write-Warning "Port $Port is in use by process: $processName (PID: $processId)"
        Write-Warning "Use -Force to override or stop the conflicting process"
        exit 1
    }
}

# Create Express server script for NSSM
$ServerScriptPath = Join-Path $ProjectRoot "server.js"
Write-Host "Creating server script for NSSM..." -ForegroundColor Yellow
$ServerScript | Out-File -FilePath $ServerScriptPath -Encoding UTF8
Write-Host "  Server script created: $ServerScriptPath" -ForegroundColor Green

# Check if Express is installed (required for server script)
$packageJsonPath = Join-Path $ProjectRoot "package.json"
$packageJson = Get-Content $packageJsonPath | ConvertFrom-Json
$hasExpress = $packageJson.dependencies.express -or $packageJson.devDependencies.express

if (-not $hasExpress) {
    Write-Host "Installing Express for server..." -ForegroundColor Yellow
    npm install express --save
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to install Express"
    }
}

# Stop and remove existing service
$existingService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($existingService) {
    Write-Host "Stopping existing service..." -ForegroundColor Yellow
    
    if ($existingService.Status -eq "Running") {
        Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
        
        # Wait for service to fully stop with status checking
        $timeout = 15
        $elapsed = 0
        do {
            Start-Sleep -Seconds 2
            $elapsed += 2
            $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        } while ($service.Status -eq "Running" -and $elapsed -lt $timeout)
        
        if ($service.Status -eq "Running") {
            Write-Warning "Service did not stop within $timeout seconds, forcing removal"
        } else {
            Write-Host "Service stopped successfully" -ForegroundColor Green
        }
    }
    
    Write-Host "Removing existing service..." -ForegroundColor Yellow
    nssm remove $ServiceName confirm
    Start-Sleep -Seconds 3
    
    # Verify removal with retries
    $timeout = 10
    $elapsed = 0
    do {
        Start-Sleep -Seconds 1
        $elapsed += 1
        $checkService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    } while ($checkService -and $elapsed -lt $timeout)
    
    if ($checkService) {
        Write-Error "Failed to remove existing service after $timeout seconds"
    } else {
        Write-Host "Service removed successfully" -ForegroundColor Green
    }
    }
}

# Install new NSSM service
Write-Host "Installing NSSM service..." -ForegroundColor Yellow
nssm install $ServiceName $NodeExe $ServerScriptPath

if ($LASTEXITCODE -ne 0) {
    Write-Error "Failed to install NSSM service"
}

# Configure service parameters
Write-Host "Configuring service..." -ForegroundColor Yellow
nssm set $ServiceName DisplayName $ServiceDisplayName
nssm set $ServiceName Description $ServiceDescription
nssm set $ServiceName AppDirectory $ProjectRoot
nssm set $ServiceName Start SERVICE_AUTO_START

# Set environment variables
$nodeEnv = if ($Environment -eq "development") { "development" } else { "production" }
nssm set $ServiceName AppEnvironmentExtra "NODE_ENV=$nodeEnv;PORT=$Port;HOST=127.0.0.1"

# Configure logging
nssm set $ServiceName AppStdout "$LogDir\frontend-stdout.log"
nssm set $ServiceName AppStderr "$LogDir\frontend-stderr.log" 
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

# Wait for service to start with status checking
$timeout = 20
$elapsed = 0
do {
    Start-Sleep -Seconds 2
    $elapsed += 2
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
} while ((-not $service -or $service.Status -ne "Running") -and $elapsed -lt $timeout)

# Verify service status
if (-not $service -or $service.Status -ne "Running") {
    Write-Error "Service failed to start within $timeout seconds. Check logs at: $LogDir"
    if ($service) {
        Write-Host "  Current Status: $($service.Status)" -ForegroundColor Red
    }
    exit 1
}

Write-Host "Service started successfully!" -ForegroundColor Green
Write-Host "  Status: $($service.Status)" -ForegroundColor Green

# Test connectivity
Write-Host "Testing frontend connectivity..." -ForegroundColor Yellow
Start-Sleep -Seconds 3

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
Write-Host "  Logs: $LogDir"
Write-Host ""
Write-Host "Access URLs:" -ForegroundColor Cyan
Write-Host "  Health Check:  http://localhost:$Port/health"
Write-Host "  Taskpane:      http://localhost:$Port/excellence/taskpane.html"
Write-Host "  Commands:      http://localhost:$Port/excellence/commands.html"
Write-Host "  Functions:     http://localhost:$Port/functions.json"
Write-Host ""
Write-Host "Management Commands:" -ForegroundColor Cyan
Write-Host "  Start:   Start-Service -Name '$ServiceName'"
Write-Host "  Stop:    Stop-Service -Name '$ServiceName'"
Write-Host "  Status:  Get-Service -Name '$ServiceName'" 
Write-Host "  Logs:    Get-Content '$LogDir\frontend-stderr.log' -Tail 50"
Write-Host ""