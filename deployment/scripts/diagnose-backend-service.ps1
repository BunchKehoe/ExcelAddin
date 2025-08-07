# PowerShell script to diagnose Excel Add-in Backend service issues
# Usage: .\diagnose-backend-service.ps1 -ServiceName "ExcelAddinBackend"

param(
    [Parameter(Mandatory=$false)]
    [string]$ServiceName = "ExcelAddinBackend",
    
    [Parameter(Mandatory=$false)]
    [string]$BackendPath = "C:\inetpub\wwwroot\ExcelAddin\backend",
    
    [Parameter(Mandatory=$false)]
    [switch]$FixCommonIssues,
    
    [Parameter(Mandatory=$false)]
    [switch]$TestManually
)

Write-Host "Excel Add-in Backend Service Diagnostics" -ForegroundColor Cyan
Write-Host "=" * 45

# Check if service exists
$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if (-not $service) {
    Write-Host "❌ Service '$ServiceName' not found" -ForegroundColor Red
    Write-Host "Run setup-backend-service.ps1 to install the service" -ForegroundColor Yellow
    exit 1
}

Write-Host "✓ Service '$ServiceName' exists" -ForegroundColor Green
Write-Host "  Status: $($service.Status)" -ForegroundColor $(if ($service.Status -eq 'Running') { 'Green' } else { 'Yellow' })
Write-Host "  Start Type: $($service.StartType)"

# Check NSSM configuration
Write-Host "`nNSSM Service Configuration:" -ForegroundColor Cyan
try {
    $nssmConfig = & nssm dump $ServiceName 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ NSSM configuration found" -ForegroundColor Green
        
        # Parse key configuration values
        $nssmConfig | ForEach-Object {
            if ($_ -match '^nssm set .* Application (.*)$') {
                Write-Host "  Application: $($matches[1])" -ForegroundColor White
            }
            elseif ($_ -match '^nssm set .* AppDirectory (.*)$') {
                Write-Host "  Working Directory: $($matches[1])" -ForegroundColor White
            }
            elseif ($_ -match '^nssm set .* AppParameters (.*)$') {
                Write-Host "  Arguments: $($matches[1])" -ForegroundColor White
            }
        }
    }
} catch {
    Write-Host "❌ Could not retrieve NSSM configuration" -ForegroundColor Red
}

# Check Python executable
Write-Host "`nPython Environment:" -ForegroundColor Cyan
$pythonExe = ""
try {
    $nssmConfig | ForEach-Object {
        if ($_ -match '^nssm set .* Application (.*)$') {
            $pythonExe = $matches[1].Trim('"')
        }
    }
    
    if ($pythonExe) {
        if (Test-Path $pythonExe) {
            Write-Host "✓ Python executable exists: $pythonExe" -ForegroundColor Green
            
            try {
                $pythonVersion = & "$pythonExe" --version 2>&1
                Write-Host "  Version: $pythonVersion" -ForegroundColor White
            } catch {
                Write-Host "❌ Could not get Python version" -ForegroundColor Red
            }
        } else {
            Write-Host "❌ Python executable not found: $pythonExe" -ForegroundColor Red
            Write-Host "  This is likely the cause of service startup failure" -ForegroundColor Yellow
        }
    }
} catch {
    Write-Host "❌ Could not determine Python executable from NSSM config" -ForegroundColor Red
}

# Check backend directory and files
Write-Host "`nBackend Directory:" -ForegroundColor Cyan
if (Test-Path $BackendPath) {
    Write-Host "✓ Backend directory exists: $BackendPath" -ForegroundColor Green
    
    # Check key files
    $keyFiles = @("service_wrapper.py", "app.py", "run.py", "requirements.txt")
    foreach ($file in $keyFiles) {
        $filePath = Join-Path $BackendPath $file
        if (Test-Path $filePath) {
            Write-Host "  ✓ $file" -ForegroundColor Green
        } else {
            Write-Host "  ❌ $file (missing)" -ForegroundColor Red
        }
    }
} else {
    Write-Host "❌ Backend directory not found: $BackendPath" -ForegroundColor Red
}

# Check log files
Write-Host "`nLog Files:" -ForegroundColor Cyan
$logDir = "C:\Logs\ExcelAddin"
$logFiles = @(
    "$logDir\backend-service-stdout.log",
    "$logDir\backend-service-stderr.log", 
    "$logDir\backend-service.log"
)

foreach ($logFile in $logFiles) {
    if (Test-Path $logFile) {
        $size = (Get-Item $logFile).Length
        $lastWrite = (Get-Item $logFile).LastWriteTime
        Write-Host "  ✓ $(Split-Path $logFile -Leaf) (${size} bytes, modified: $lastWrite)" -ForegroundColor Green
        
        # Show recent errors from log
        if ($logFile -match "stderr") {
            $recentErrors = Get-Content $logFile -Tail 10 -ErrorAction SilentlyContinue | Where-Object { $_ -match "(error|exception|failed|traceback)" -and $_ -ne "" }
            if ($recentErrors) {
                Write-Host "    Recent errors:" -ForegroundColor Yellow
                $recentErrors | ForEach-Object { Write-Host "      $_" -ForegroundColor Red }
            }
        }
    } else {
        Write-Host "  ❌ $(Split-Path $logFile -Leaf) (not found)" -ForegroundColor Red
    }
}

# Check if port 5000 is in use
Write-Host "`nNetwork Connectivity:" -ForegroundColor Cyan
try {
    $port5000 = Get-NetTCPConnection -LocalPort 5000 -ErrorAction SilentlyContinue
    if ($port5000) {
        Write-Host "✓ Port 5000 is in use (process ID: $($port5000.OwningProcess))" -ForegroundColor Green
        
        # Try to identify the process
        try {
            $process = Get-Process -Id $port5000.OwningProcess -ErrorAction SilentlyContinue
            if ($process) {
                Write-Host "  Process: $($process.ProcessName) ($($process.Id))" -ForegroundColor White
            }
        } catch { }
        
    } else {
        Write-Host "❌ Port 5000 is not in use" -ForegroundColor Red
        Write-Host "  This indicates the Flask backend is not running" -ForegroundColor Yellow
    }
} catch {
    Write-Host "❌ Could not check port 5000 status" -ForegroundColor Red
}

# Test HTTP connectivity
try {
    Write-Host "Testing HTTP connectivity to Flask backend..." -ForegroundColor White
    $response = Invoke-WebRequest -Uri "http://127.0.0.1:5000/health" -TimeoutSec 5 -ErrorAction SilentlyContinue
    if ($response.StatusCode -eq 200) {
        Write-Host "✓ Flask backend is responding (HTTP 200)" -ForegroundColor Green
    } else {
        Write-Host "⚠ Flask backend returned HTTP $($response.StatusCode)" -ForegroundColor Yellow
    }
} catch {
    Write-Host "❌ Cannot connect to Flask backend on port 5000" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Yellow
}

# Check Windows Event Log
Write-Host "`nWindows Event Log:" -ForegroundColor Cyan
try {
    $recentEvents = Get-EventLog -LogName System -After (Get-Date).AddHours(-24) -ErrorAction SilentlyContinue | 
                   Where-Object { $_.Source -eq "Service Control Manager" -and $_.Message -match $ServiceName } |
                   Select-Object -First 5
    
    if ($recentEvents) {
        Write-Host "✓ Found recent service events in Event Log" -ForegroundColor Green
        $recentEvents | ForEach-Object {
            $level = switch ($_.EntryType) {
                'Error' { 'Red' }
                'Warning' { 'Yellow' }
                default { 'White' }
            }
            Write-Host "  [$($_.TimeGenerated)] $($_.EntryType): $($_.Message)" -ForegroundColor $level
        }
    } else {
        Write-Host "⚠ No recent service events found in Event Log" -ForegroundColor Yellow
    }
} catch {
    Write-Host "❌ Could not check Event Log: $_" -ForegroundColor Red
}

# Manual test option
if ($TestManually) {
    Write-Host "`nRunning Manual Test:" -ForegroundColor Cyan
    if ($pythonExe -and (Test-Path $pythonExe) -and (Test-Path $BackendPath)) {
        Write-Host "Testing service wrapper manually..." -ForegroundColor White
        Write-Host "Press Ctrl+C to stop the test" -ForegroundColor Yellow
        
        Push-Location $BackendPath
        try {
            & "$pythonExe" "service_wrapper.py"
        } catch {
            Write-Host "Manual test failed: $_" -ForegroundColor Red
        } finally {
            Pop-Location
        }
    } else {
        Write-Host "❌ Cannot run manual test - Python or backend path not available" -ForegroundColor Red
    }
}

# Common fixes
if ($FixCommonIssues) {
    Write-Host "`nApplying Common Fixes:" -ForegroundColor Cyan
    
    # Fix 1: Update Python path to full path
    Write-Host "1. Checking Python executable path..." -ForegroundColor White
    if ($pythonExe -eq "python" -or $pythonExe -eq "python.exe") {
        Write-Host "  Fixing relative Python path..." -ForegroundColor Yellow
        try {
            $fullPythonPath = (Get-Command python).Source
            if ($fullPythonPath) {
                & nssm set $ServiceName Application "$fullPythonPath"
                Write-Host "  ✓ Updated Python path to: $fullPythonPath" -ForegroundColor Green
            }
        } catch {
            Write-Host "  ❌ Could not find full Python path" -ForegroundColor Red
        }
    }
    
    # Fix 2: Ensure log directory exists
    Write-Host "2. Ensuring log directory exists..." -ForegroundColor White
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        Write-Host "  ✓ Created log directory: $logDir" -ForegroundColor Green
    }
    
    # Fix 3: Reset service configuration
    Write-Host "3. Resetting service environment..." -ForegroundColor White
    $envVars = @(
        "FLASK_ENV=production",
        "DEBUG=false", 
        "HOST=127.0.0.1",
        "PORT=5000",
        "PYTHONPATH=$BackendPath",
        "PYTHONIOENCODING=utf-8",
        "PYTHONUNBUFFERED=1"
    )
    & nssm set $ServiceName AppEnvironmentExtra ($envVars -join "`0")
    Write-Host "  ✓ Updated environment variables" -ForegroundColor Green
    
    Write-Host "`nFixes applied. Try restarting the service:" -ForegroundColor Green
    Write-Host "  Restart-Service $ServiceName" -ForegroundColor White
}

Write-Host "`n" + "=" * 45 -ForegroundColor Cyan
Write-Host "DIAGNOSTICS COMPLETE" -ForegroundColor Cyan
Write-Host "=" * 45 -ForegroundColor Cyan

# Summary and recommendations
Write-Host "`nSUMMARY & RECOMMENDATIONS:" -ForegroundColor Green

$issues = @()
$recommendations = @()

if ($service.Status -ne 'Running') {
    $issues += "Service is not running"
    $recommendations += "Check log files for error messages"
    $recommendations += "Try starting manually: .\debug-service.bat"
}

if ($pythonExe -and -not (Test-Path $pythonExe)) {
    $issues += "Python executable not found"
    $recommendations += "Run with -FixCommonIssues to update Python path"
    $recommendations += "Or reinstall service with correct Python path"
}

if (-not (Test-Path $BackendPath)) {
    $issues += "Backend directory missing"
    $recommendations += "Verify backend files are deployed correctly"
}

try {
    $port5000 = Get-NetTCPConnection -LocalPort 5000 -ErrorAction SilentlyContinue
    if (-not $port5000) {
        $issues += "Flask backend not listening on port 5000"
        $recommendations += "Check service logs for startup errors"
    }
} catch { }

if ($issues.Count -eq 0) {
    Write-Host "✓ No major issues detected" -ForegroundColor Green
    if ($service.Status -eq 'Running') {
        Write-Host "✓ Service appears to be healthy" -ForegroundColor Green
    }
} else {
    Write-Host "Issues found:" -ForegroundColor Yellow
    $issues | ForEach-Object { Write-Host "  • $_" -ForegroundColor Red }
    
    Write-Host "`nRecommendations:" -ForegroundColor Yellow
    $recommendations | ForEach-Object { Write-Host "  • $_" -ForegroundColor White }
}

Write-Host "`nNext Steps:" -ForegroundColor Green
Write-Host "• Check recent log files: Get-Content C:\Logs\ExcelAddin\backend-service-stderr.log -Tail 20"
Write-Host "• Manual test: cd $BackendPath && debug-service.bat"
Write-Host "• Apply fixes: .\diagnose-backend-service.ps1 -FixCommonIssues"
Write-Host "• Reinstall service: .\setup-backend-service.ps1 -Force"
Write-Host "• Edit NSSM config: nssm edit $ServiceName"