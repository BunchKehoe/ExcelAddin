# Alternative frontend deployment using Windows Task Scheduler
# Fallback option if NSSM doesn't work reliably

param(
    [switch]$SkipBuild,
    [switch]$Remove
)

. (Join-Path $PSScriptRoot "scripts\common.ps1")

$TaskName = "ExcelAddin-Frontend"
$TaskDescription = "Excel Add-in Frontend Web Server (Task Scheduler)"

Write-Header "Alternative Frontend Deployment (Task Scheduler)"

if ($Remove) {
    Write-Host "Removing existing task..."
    try {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
        Write-Success "Task removed successfully"
    } catch {
        Write-Warning "Error removing task: $($_.Exception.Message)"
    }
    exit 0
}

# Get paths
$ProjectRoot = Get-ProjectRoot
$FrontendPath = Get-FrontendPath
$ServerScript = Join-Path $PSScriptRoot "config\frontend-server.js"

Write-Host "Project Root: $ProjectRoot"
Write-Host "Frontend Path: $FrontendPath"
Write-Host "Server Script: $ServerScript"

# Check prerequisites
if (-not (Test-Administrator)) {
    Write-Error "This script must be run as Administrator"
    exit 1
}

if (-not (Test-Path $ServerScript)) {
    Write-Error "Server script not found: $ServerScript"
    exit 1
}

# Build if needed
if (-not $SkipBuild) {
    Write-Header "Building Frontend"
    Set-Location $FrontendPath
    
    npm install
    if ($LASTEXITCODE -ne 0) {
        Write-Error "npm install failed"
        exit 1
    }
    
    npm run build:staging
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Build failed"
        exit 1
    }
    
    Write-Success "Build completed"
}

# Verify dist directory exists
$distPath = Join-Path $FrontendPath "dist"
if (-not (Test-Path $distPath)) {
    Write-Error "Dist directory not found: $distPath"
    exit 1
}

# Remove existing task
Write-Host "Removing existing task if present..."
try {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
} catch {}

# Get Node.js path
$nodeCmd = Get-Command node -ErrorAction Stop
$nodePath = $nodeCmd.Source
Write-Host "Using Node.js: $nodePath"

# Create batch file wrapper for better process management
$batchFile = Join-Path $PSScriptRoot "frontend-task.bat"
$batchContent = @"
@echo off
cd /d "$FrontendPath"
set NODE_ENV=production
set PORT=3000
set HOST=127.0.0.1
"$nodePath" "$ServerScript"
pause
"@

Set-Content -Path $batchFile -Value $batchContent -Encoding ASCII
Write-Host "Created batch wrapper: $batchFile"

# Create PowerShell wrapper script for better logging
$wrapperScript = Join-Path $PSScriptRoot "frontend-wrapper.ps1"
$wrapperContent = @"
# Frontend service wrapper script - DO NOT EDIT
Set-Location "$FrontendPath"
`$env:NODE_ENV = "production"
`$env:PORT = "3000"
`$env:HOST = "127.0.0.1"

`$logDir = "C:\Logs\ExcelAddin"
if (-not (Test-Path `$logDir)) {
    New-Item -ItemType Directory -Path `$logDir -Force | Out-Null
}

`$stdoutLog = "`$logDir\frontend-task-stdout.log"
`$stderrLog = "`$logDir\frontend-task-stderr.log"

Write-Host "Starting ExcelAddin Frontend Server..."
Write-Host "Logs: `$stdoutLog"

try {
    & "$nodePath" "$ServerScript" 2>`$stderrLog | Tee-Object -FilePath `$stdoutLog
} catch {
    "Error: `$(`$_.Exception.Message)" | Out-File -FilePath `$stderrLog -Append
    throw
}
"@

Set-Content -Path $wrapperScript -Value $wrapperContent -Encoding UTF8
Write-Host "Created PowerShell wrapper: $wrapperScript"

# Create scheduled task
Write-Host "Creating scheduled task..."

$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$wrapperScript`""
$trigger = New-ScheduledTaskTrigger -AtStartup
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)

$task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description $TaskDescription

# Register the task
Register-ScheduledTask -TaskName $TaskName -InputObject $task -Force

Write-Success "Scheduled task created: $TaskName"

# Start the task
Write-Host "Starting task..."
Start-ScheduledTask -TaskName $TaskName

# Wait and test
Write-Host "Waiting for service to start..."
Start-Sleep -Seconds 10

# Test the service
$testAttempts = 0
$maxAttempts = 6
$serviceWorking = $false

while ($testAttempts -lt $maxAttempts -and -not $serviceWorking) {
    $testAttempts++
    Write-Host "Test attempt $testAttempts/$maxAttempts..."
    
    # Check if port is listening
    $portTest = Test-NetConnection -ComputerName "127.0.0.1" -Port 3000 -InformationLevel Quiet -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    if ($portTest) {
        Write-Success "Port 3000 is listening"
        
        # Test HTTP
        try {
            $response = Invoke-WebRequest -Uri "http://127.0.0.1:3000" -TimeoutSec 10 -ErrorAction Stop
            Write-Success "HTTP test passed - Status: $($response.StatusCode)"
            $serviceWorking = $true
        } catch {
            Write-Warning "HTTP test failed: $($_.Exception.Message)"
        }
    } else {
        Write-Host "Port not yet listening..."
    }
    
    if (-not $serviceWorking -and $testAttempts -lt $maxAttempts) {
        Start-Sleep -Seconds 5
    }
}

if ($serviceWorking) {
    Write-Success "Alternative deployment successful!"
    Write-Host ""
    Write-Host "Task Information:"
    Write-Host "  Name: $TaskName"
    Write-Host "  Status: Running at startup"
    Write-Host "  URL: http://127.0.0.1:3000"
    Write-Host ""
    Write-Host "Management Commands:"
    Write-Host "  Status: Get-ScheduledTask -TaskName '$TaskName'"
    Write-Host "  Start: Start-ScheduledTask -TaskName '$TaskName'"
    Write-Host "  Stop: Stop-ScheduledTask -TaskName '$TaskName'"
    Write-Host "  Remove: .\deploy-frontend-alternative.ps1 -Remove"
    Write-Host "  Logs: C:\Logs\ExcelAddin\frontend-task-*.log"
} else {
    Write-Error "Alternative deployment failed"
    
    # Show task status
    $taskInfo = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($taskInfo) {
        Write-Host "Task Status: $($taskInfo.State)"
        
        # Show logs if available
        $logFiles = @(
            "C:\Logs\ExcelAddin\frontend-task-stdout.log",
            "C:\Logs\ExcelAddin\frontend-task-stderr.log"
        )
        
        foreach ($logFile in $logFiles) {
            if (Test-Path $logFile) {
                Write-Host ""
                Write-Host "Log: $logFile"
                Get-Content $logFile -Tail 10 | ForEach-Object { Write-Host "  $_" }
            }
        }
    }
    
    exit 1
}