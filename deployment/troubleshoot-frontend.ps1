# Frontend deployment troubleshooting and recovery script
# Comprehensive tool to diagnose and fix frontend deployment issues

param(
    [switch]$AutoFix,
    [switch]$SwitchToAlternative,
    [string]$Method = "Auto"  # Auto, NSSM, TaskScheduler
)

. (Join-Path $PSScriptRoot "scripts\common.ps1")

$ServiceName = "ExcelAddin-Frontend"
$TaskName = "ExcelAddin-Frontend"

Write-Header "Frontend Deployment Troubleshooting"

Write-Host "Auto-fix mode: $AutoFix"
Write-Host "Switch to alternative: $SwitchToAlternative"
Write-Host "Method: $Method"
Write-Host ""

# Quick 404 issue detection first
Write-Host "Performing quick 404 issue detection..."
try {
    $quickTest = Invoke-WebRequest -Uri "http://127.0.0.1:3000" -TimeoutSec 5 -ErrorAction Stop
    if ($quickTest.StatusCode -eq 404) {
        Write-Warning "DETECTED: HTTP 404 error - Frontend files missing or misconfigured!"
        Write-Host ""
        Write-Host "This is a common issue where the NSSM service is running but cannot find the built frontend files."
        Write-Host "Run the specialized fix: .\fix-frontend-404.ps1 -AutoFix"
        Write-Host ""
        if ($AutoFix) {
            Write-Host "Auto-fix mode enabled - running 404 fix..." -ForegroundColor Green
            & (Join-Path $PSScriptRoot "fix-frontend-404.ps1") -AutoFix
            return
        } else {
            $answer = Read-Host "Would you like to run the 404 fix now? (y/n)"
            if ($answer -eq 'y' -or $answer -eq 'Y') {
                & (Join-Path $PSScriptRoot "fix-frontend-404.ps1") -AutoFix
                return
            }
        }
    }
} catch {
    # Continue with full troubleshooting if we can't connect at all
}
Write-Host ""

# Functions for different deployment methods
function Test-NSSMDeployment {
    Write-Host "=== Testing NSSM Deployment ===" -ForegroundColor Cyan
    
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if (-not $service) {
        Write-Host "NSSM service not found" -ForegroundColor Red
        return $false
    }
    
    Write-Host "Service status: $($service.Status)"
    
    if ($service.Status -ne 'Running') {
        Write-Host "Service is not running" -ForegroundColor Red
        return $false
    }
    
    # Test port
    $portTest = Test-NetConnection -ComputerName "127.0.0.1" -Port 3000 -InformationLevel Quiet -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    if (-not $portTest) {
        Write-Host "Port 3000 not listening" -ForegroundColor Red
        return $false
    }
    
    # Test HTTP
    try {
        $response = Invoke-WebRequest -Uri "http://127.0.0.1:3000" -TimeoutSec 5 -ErrorAction Stop
        Write-Host "NSSM deployment: WORKING" -ForegroundColor Green
        return $true
    } catch {
        Write-Host "HTTP test failed: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Test-TaskSchedulerDeployment {
    Write-Host "=== Testing Task Scheduler Deployment ===" -ForegroundColor Cyan
    
    $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if (-not $task) {
        Write-Host "Scheduled task not found" -ForegroundColor Red
        return $false
    }
    
    Write-Host "Task status: $($task.State)"
    
    if ($task.State -ne 'Running') {
        Write-Host "Task is not running" -ForegroundColor Red
        return $false
    }
    
    # Test port
    $portTest = Test-NetConnection -ComputerName "127.0.0.1" -Port 3000 -InformationLevel Quiet -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    if (-not $portTest) {
        Write-Host "Port 3000 not listening" -ForegroundColor Red
        return $false
    }
    
    # Test HTTP
    try {
        $response = Invoke-WebRequest -Uri "http://127.0.0.1:3000" -TimeoutSec 5 -ErrorAction Stop
        Write-Host "Task Scheduler deployment: WORKING" -ForegroundColor Green
        return $true
    } catch {
        Write-Host "HTTP test failed: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Get-CurrentDeploymentMethod {
    $nssmWorking = Test-NSSMDeployment
    $taskWorking = Test-TaskSchedulerDeployment
    
    if ($nssmWorking -and $taskWorking) {
        Write-Warning "Both NSSM and Task Scheduler deployments are running!"
        return "Both"
    } elseif ($nssmWorking) {
        return "NSSM"
    } elseif ($taskWorking) {
        return "TaskScheduler"
    } else {
        return "None"
    }
}

function Stop-AllDeployments {
    Write-Host "Stopping all deployment methods..." -ForegroundColor Yellow
    
    # Stop NSSM service
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($service -and $service.Status -eq 'Running') {
        Write-Host "Stopping NSSM service..."
        Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
    }
    
    # Stop scheduled task
    $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($task -and $task.State -eq 'Running') {
        Write-Host "Stopping scheduled task..."
        Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    }
    
    # Kill any remaining node processes on port 3000
    $portProcesses = Get-NetTCPConnection -LocalPort 3000 -ErrorAction SilentlyContinue
    foreach ($conn in $portProcesses) {
        $process = Get-Process -Id $conn.OwningProcess -ErrorAction SilentlyContinue
        if ($process -and $process.ProcessName -eq 'node') {
            Write-Host "Killing node process PID $($process.Id)..."
            Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
        }
    }
    
    Start-Sleep -Seconds 3
    Write-Success "All deployments stopped"
}

function Repair-NSSMDeployment {
    Write-Host "=== Repairing NSSM Deployment ===" -ForegroundColor Cyan
    
    # Run the deployment script
    $deployScript = Join-Path $PSScriptRoot "deploy-frontend.ps1"
    if (Test-Path $deployScript) {
        Write-Host "Running NSSM deployment..."
        & $deployScript -SkipBuild
        
        Start-Sleep -Seconds 5
        return Test-NSSMDeployment
    } else {
        Write-Error "Deploy script not found: $deployScript"
        return $false
    }
}

function Install-AlternativeDeployment {
    Write-Host "=== Installing Alternative Deployment ===" -ForegroundColor Cyan
    
    # Run the alternative deployment script
    $altScript = Join-Path $PSScriptRoot "deploy-frontend-alternative.ps1"
    if (Test-Path $altScript) {
        Write-Host "Running Task Scheduler deployment..."
        & $altScript -SkipBuild
        
        Start-Sleep -Seconds 5
        return Test-TaskSchedulerDeployment
    } else {
        Write-Error "Alternative deploy script not found: $altScript"
        return $false
    }
}

# Main troubleshooting logic
Write-Header "Deployment Status Check"

$currentMethod = Get-CurrentDeploymentMethod
Write-Host "Current deployment method: $currentMethod" -ForegroundColor $(if ($currentMethod -eq 'None') { 'Red' } else { 'Green' })

if ($currentMethod -eq 'Both') {
    Write-Warning "Multiple deployment methods are active - this can cause conflicts"
    if ($AutoFix) {
        Stop-AllDeployments
        # Default to NSSM unless specified otherwise
        if ($Method -eq 'TaskScheduler') {
            $success = Install-AlternativeDeployment
        } else {
            $success = Repair-NSSMDeployment
        }
        
        if ($success) {
            Write-Success "Deployment repaired successfully"
        } else {
            Write-Error "Failed to repair deployment"
        }
    }
}
elseif ($currentMethod -eq 'None') {
    Write-Host "No working deployment found" -ForegroundColor Red
    
    if ($AutoFix -or $SwitchToAlternative) {
        # Try to fix based on method preference
        if ($Method -eq 'TaskScheduler' -or $SwitchToAlternative) {
            Write-Host "Attempting Task Scheduler deployment..."
            $success = Install-AlternativeDeployment
        } else {
            Write-Host "Attempting NSSM deployment repair..."
            $success = Repair-NSSMDeployment
            
            if (-not $success) {
                Write-Host "NSSM deployment failed, trying alternative..."
                $success = Install-AlternativeDeployment
            }
        }
        
        if ($success) {
            Write-Success "Deployment restored successfully"
        } else {
            Write-Error "All deployment methods failed"
            
            Write-Host ""
            Write-Host "Manual troubleshooting steps:"
            Write-Host "1. Run: .\diagnose-nssm.ps1"
            Write-Host "2. Run: .\test-frontend-server.ps1"
            Write-Host "3. Check build: npm run build:staging"
            Write-Host "4. Check logs: C:\Logs\ExcelAddin\"
        }
    } else {
        Write-Host ""
        Write-Host "Recommended actions:"
        Write-Host "1. Run with -AutoFix to attempt automatic repair"
        Write-Host "2. Run with -SwitchToAlternative to use Task Scheduler"
        Write-Host "3. Manual diagnostics: .\diagnose-nssm.ps1"
    }
}
else {
    Write-Success "Deployment is working correctly using: $currentMethod"
}

# Show final status
Write-Host ""
Write-Header "Final Status"

$finalMethod = Get-CurrentDeploymentMethod
Write-Host "Active deployment method: $finalMethod"

if ($finalMethod -ne 'None') {
    Write-Host "Frontend URL: http://127.0.0.1:3000" -ForegroundColor Green
    
    if ($finalMethod -eq 'NSSM') {
        Write-Host "Management: Get-Service -Name '$ServiceName'"
        Write-Host "Logs: C:\Logs\ExcelAddin\frontend-*.log"
    } elseif ($finalMethod -eq 'TaskScheduler') {
        Write-Host "Management: Get-ScheduledTask -TaskName '$TaskName'"
        Write-Host "Logs: C:\Logs\ExcelAddin\frontend-task-*.log"
    }
} else {
    Write-Host "No working deployment found" -ForegroundColor Red
}

Write-Host ""
Write-Host "Available troubleshooting tools:"
Write-Host "  .\diagnose-nssm.ps1 - Comprehensive NSSM diagnostics"
Write-Host "  .\test-frontend-server.ps1 - Manual server testing"
Write-Host "  .\troubleshoot-frontend.ps1 -AutoFix - Auto repair"
Write-Host "  .\troubleshoot-frontend.ps1 -SwitchToAlternative - Use Task Scheduler"