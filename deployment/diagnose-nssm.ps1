# Comprehensive NSSM service diagnostics for ExcelAddin frontend
# Helps debug why NSSM service is running but not serving the website

param(
    [string]$ServiceName = "ExcelAddin-Frontend"
)

. (Join-Path $PSScriptRoot "scripts\common.ps1")

Write-Header "NSSM Service Diagnostics"

Write-Host "Service Name: $ServiceName"
Write-Host "Timestamp: $(Get-Date)"
Write-Host ""

# Check service status
Write-Host "=== SERVICE STATUS ===" -ForegroundColor Cyan
$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($service) {
    Write-Host "Service Status: $($service.Status)" -ForegroundColor $(if ($service.Status -eq 'Running') { 'Green' } else { 'Red' })
    Write-Host "Service Start Type: $($service.StartType)"
    Write-Host "Service Display Name: $($service.DisplayName)"
} else {
    Write-Error "Service '$ServiceName' not found!"
    exit 1
}

# Get NSSM configuration
Write-Host ""
Write-Host "=== NSSM CONFIGURATION ===" -ForegroundColor Cyan

$nssmCommands = @(
    "get $ServiceName Application",
    "get $ServiceName Parameters",
    "get $ServiceName AppDirectory",
    "get $ServiceName AppEnvironmentExtra",
    "get $ServiceName Start",
    "get $ServiceName AppExit",
    "get $ServiceName AppStdout",
    "get $ServiceName AppStderr"
)

foreach ($cmd in $nssmCommands) {
    try {
        $result = & nssm $cmd.Split(' ') 2>&1
        $cmdName = $cmd.Split(' ')[-1]
        Write-Host "${cmdName}: $result"
    } catch {
        Write-Warning "Failed to get $cmd : $($_.Exception.Message)"
    }
}

# Check log files
Write-Host ""
Write-Host "=== LOG FILES ===" -ForegroundColor Cyan

$logFiles = @(
    "C:\Logs\ExcelAddin\frontend-stdout.log",
    "C:\Logs\ExcelAddin\frontend-stderr.log"
)

foreach ($logFile in $logFiles) {
    Write-Host ""
    Write-Host "Log file: $logFile" -ForegroundColor Yellow
    if (Test-Path $logFile) {
        $logSize = (Get-Item $logFile).Length
        Write-Host "Size: $logSize bytes"
        Write-Host "Last modified: $((Get-Item $logFile).LastWriteTime)"
        Write-Host ""
        Write-Host "Last 10 lines:" -ForegroundColor Yellow
        try {
            $lines = Get-Content $logFile -Tail 10 -ErrorAction Stop
            if ($lines) {
                $lines | ForEach-Object { Write-Host "  $_" }
            } else {
                Write-Host "  (empty file)"
            }
        } catch {
            Write-Warning "Error reading log file: $($_.Exception.Message)"
        }
    } else {
        Write-Warning "Log file not found"
    }
}

# Check port status
Write-Host ""
Write-Host "=== PORT STATUS ===" -ForegroundColor Cyan

$ports = @(3000, 5000, 9443)
foreach ($port in $ports) {
    Write-Host ""
    Write-Host "Port $port:" -ForegroundColor Yellow
    
    # Check if port is listening
    try {
        $connection = Get-NetTCPConnection -LocalPort $port -ErrorAction SilentlyContinue
        if ($connection) {
            $process = Get-Process -Id $connection.OwningProcess -ErrorAction SilentlyContinue
            Write-Host "  Status: LISTENING"
            Write-Host "  Process: $($process.ProcessName) (PID: $($process.Id))"
            Write-Host "  Local Address: $($connection.LocalAddress):$($connection.LocalPort)"
        } else {
            Write-Host "  Status: NOT LISTENING" -ForegroundColor Red
        }
    } catch {
        Write-Host "  Status: ERROR - $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # Test HTTP connection
    if ($port -in @(3000, 5000)) {
        try {
            Write-Host "  Testing HTTP connection..."
            $response = Invoke-WebRequest -Uri "http://127.0.0.1:$port" -TimeoutSec 5 -ErrorAction Stop
            Write-Host "  HTTP Status: $($response.StatusCode)" -ForegroundColor Green
            Write-Host "  Content Length: $($response.Content.Length) bytes"
        } catch {
            Write-Host "  HTTP Test: FAILED - $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# Check file system paths
Write-Host ""
Write-Host "=== FILE SYSTEM PATHS ===" -ForegroundColor Cyan

$ProjectRoot = Get-ProjectRoot
$FrontendPath = Get-FrontendPath
$ServerScript = Join-Path $PSScriptRoot "config\frontend-server.js"

$pathsToCheck = @{
    "Project Root" = $ProjectRoot
    "Frontend Path" = $FrontendPath
    "Server Script" = $ServerScript
    "Dist Directory" = (Join-Path $FrontendPath "dist")
    "Index.html" = (Join-Path $FrontendPath "dist\index.html")
}

foreach ($pathName in $pathsToCheck.Keys) {
    $pathValue = $pathsToCheck[$pathName]
    Write-Host ""
    Write-Host "$pathName :" -ForegroundColor Yellow
    Write-Host "  $pathValue"
    
    if (Test-Path $pathValue) {
        $item = Get-Item $pathValue
        if ($item.PSIsContainer) {
            Write-Host "  Status: EXISTS (Directory)" -ForegroundColor Green
            # List contents if it's dist directory
            if ($pathName -eq "Dist Directory") {
                try {
                    $contents = Get-ChildItem $pathValue -Force | Select-Object -First 10
                    Write-Host "  Contents (first 10):"
                    $contents | ForEach-Object { Write-Host "    $($_.Name)" }
                } catch {
                    Write-Host "  Error listing contents: $($_.Exception.Message)"
                }
            }
        } else {
            Write-Host "  Status: EXISTS (File)" -ForegroundColor Green
            Write-Host "  Size: $($item.Length) bytes"
        }
    } else {
        Write-Host "  Status: NOT FOUND" -ForegroundColor Red
    }
}

# Check Node.js accessibility from service context
Write-Host ""
Write-Host "=== NODE.JS VERIFICATION ===" -ForegroundColor Cyan

try {
    $nodeCmd = Get-Command node -ErrorAction Stop
    Write-Host "Node.js Path: $($nodeCmd.Source)" -ForegroundColor Green
    
    $nodeVersion = & node --version
    Write-Host "Node.js Version: $nodeVersion" -ForegroundColor Green
    
    # Test if the server script can be parsed
    Write-Host "Testing script syntax..."
    $syntaxTest = & node -c $ServerScript 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Script Syntax: OK" -ForegroundColor Green
    } else {
        Write-Host "Script Syntax: ERROR" -ForegroundColor Red
        Write-Host "Error: $syntaxTest"
    }
} catch {
    Write-Error "Node.js not available: $($_.Exception.Message)"
}

# Environment variables that NSSM should set
Write-Host ""
Write-Host "=== ENVIRONMENT VARIABLES ===" -ForegroundColor Cyan

$expectedEnvVars = @("NODE_ENV", "PORT", "HOST")
foreach ($envVar in $expectedEnvVars) {
    $value = [Environment]::GetEnvironmentVariable($envVar)
    if ($value) {
        Write-Host "$envVar = $value" -ForegroundColor Green
    } else {
        Write-Host "$envVar = (not set)" -ForegroundColor Yellow
    }
}

# Service process information
Write-Host ""
Write-Host "=== SERVICE PROCESS INFO ===" -ForegroundColor Cyan

if ($service.Status -eq 'Running') {
    try {
        # Get service PID using SC command
        $scOutput = & sc queryex $ServiceName 2>&1
        $pidLine = $scOutput | Where-Object { $_ -match "PID\s*:\s*(\d+)" }
        if ($pidLine -and $Matches[1]) {
            $pid = [int]$Matches[1]
            Write-Host "Service PID: $pid"
            
            $process = Get-Process -Id $pid -ErrorAction SilentlyContinue
            if ($process) {
                Write-Host "Process Name: $($process.ProcessName)"
                Write-Host "Process Path: $($process.Path)"
                Write-Host "Working Set: $([math]::Round($process.WorkingSet / 1MB, 2)) MB"
                Write-Host "Start Time: $($process.StartTime)"
            }
        } else {
            Write-Warning "Could not determine service PID"
        }
    } catch {
        Write-Warning "Error getting service process info: $($_.Exception.Message)"
    }
} else {
    Write-Host "Service is not running"
}

Write-Host ""
Write-Header "Diagnostics Complete"

# Recommendations
Write-Host ""
Write-Host "=== RECOMMENDATIONS ===" -ForegroundColor Cyan

if ($service.Status -ne 'Running') {
    Write-Host "1. Service is not running - start it with: Start-Service -Name '$ServiceName'"
} else {
    $portListening = Get-NetTCPConnection -LocalPort 3000 -ErrorAction SilentlyContinue
    if (-not $portListening) {
        Write-Host "1. Service is running but not listening on port 3000"
        Write-Host "   - Check stderr log for startup errors"
        Write-Host "   - Verify dist directory exists and contains files"
        Write-Host "   - Test server manually: .\test-frontend-server.ps1"
    } else {
        Write-Host "1. Service appears to be listening on port 3000"
        Write-Host "   - If HTTP test failed, check firewall settings"
        Write-Host "   - Verify IIS proxy configuration if using HTTPS"
    }
}

Write-Host "2. Run manual test: .\test-frontend-server.ps1"
Write-Host "3. Check NSSM service logs for errors"
Write-Host "4. If issues persist, consider alternative deployment methods"