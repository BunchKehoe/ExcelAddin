# Helper functions for ExcelAddin deployment scripts
# DO NOT run this script directly - it is sourced by other scripts

function Write-Header {
    param([string]$Message)
    Write-Host ""
    Write-Host "=" * 60 -ForegroundColor Cyan
    Write-Host $Message -ForegroundColor Cyan
    Write-Host "=" * 60 -ForegroundColor Cyan
}

function Write-Success {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Green
}

function Write-Warning {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Yellow
}

function Write-Error {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Red
}

function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-Command {
    param([string]$CommandName)
    try {
        Get-Command $CommandName -ErrorAction Stop | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Test-Port {
    param([int]$Port)
    try {
        $connection = New-Object System.Net.Sockets.TcpClient
        $connection.Connect("127.0.0.1", $Port)
        $connection.Close()
        return $true
    } catch {
        return $false
    }
}

function Stop-ServiceSafely {
    param([string]$ServiceName)
    try {
        $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if ($service -and $service.Status -eq "Running") {
            Write-Host "Stopping service: $ServiceName"
            Stop-Service -Name $ServiceName -Force -ErrorAction Stop
            
            # Wait for service to stop
            $timeout = 30
            $elapsed = 0
            while ($service.Status -ne "Stopped" -and $elapsed -lt $timeout) {
                Start-Sleep -Seconds 1
                $elapsed++
                $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
            }
            
            if ($service.Status -eq "Stopped") {
                Write-Success "Service $ServiceName stopped successfully"
                return $true
            } else {
                Write-Warning "Service $ServiceName did not stop within $timeout seconds"
                return $false
            }
        }
        return $true
    } catch {
        Write-Warning "Failed to stop service $ServiceName : $($_.Exception.Message)"
        return $false
    }
}

function Get-ProjectRoot {
    $scriptPath = $PSScriptRoot
    # Go up from deployment folder to project root
    return Split-Path -Parent $scriptPath
}

function Get-BackendPath {
    $projectRoot = Get-ProjectRoot
    return Join-Path $projectRoot "backend"
}

function Get-FrontendPath {
    $projectRoot = Get-ProjectRoot
    return $projectRoot  # Frontend is at project root level
}

function Test-Prerequisites {
    param([switch]$SkipPM2, [switch]$SkipNSSM)
    
    $allGood = $true
    
    Write-Header "Checking Prerequisites"
    
    # Check Administrator privileges
    if (-not (Test-Administrator)) {
        Write-Error "This script must be run as Administrator"
        $allGood = $false
    } else {
        Write-Success "Administrator privileges: OK"
    }
    
    # Check Node.js
    if (Test-Command "node") {
        $nodeVersion = node --version
        Write-Success "Node.js: $nodeVersion"
    } else {
        Write-Error "Node.js is not installed or not in PATH"
        $allGood = $false
    }
    
    # Check Python
    if (Test-Command "python") {
        $pythonVersion = python --version 2>&1
        Write-Success "Python: $pythonVersion"
    } else {
        Write-Error "Python is not installed or not in PATH"
        $allGood = $false
    }
    
    # Check PM2 if needed
    if (-not $SkipPM2) {
        if (Test-Command "pm2") {
            Write-Success "PM2: Available"
        } else {
            Write-Error "PM2 is not installed. Run: npm install -g pm2"
            $allGood = $false
        }
    }
    
    # Check NSSM if needed
    if (-not $SkipNSSM) {
        if (Test-Command "nssm") {
            Write-Success "NSSM: Available"
        } else {
            Write-Error "NSSM is not installed or not in PATH"
            $allGood = $false
        }
    }
    
    # Check IIS
    $iisFeature = Get-WindowsOptionalFeature -Online -FeatureName IIS-WebServerRole
    if ($iisFeature.State -eq "Enabled") {
        Write-Success "IIS: Enabled"
    } else {
        Write-Error "IIS is not enabled. Enable it through Windows Features."
        $allGood = $false
    }
    
    return $allGood
}