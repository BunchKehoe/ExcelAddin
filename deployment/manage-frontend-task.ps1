# ExcelAddin Frontend Task Management Script
# Helper script for managing the Windows Task Scheduler task

param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("start", "stop", "restart", "status", "delete")]
    [string]$Action,
    [switch]$Detailed
)

$ErrorActionPreference = "Stop"

# Configuration
$TaskName = "ExcelAddin-Frontend"
$Port = 3000

function Get-TaskStatus {
    try {
        $taskInfo = schtasks /query /tn $TaskName /fo LIST 2>$null
        if ($LASTEXITCODE -eq 0) {
            $status = "Unknown"
            $lastRun = "Unknown"
            $nextRun = "Unknown"
            
            foreach ($line in $taskInfo) {
                if ($line -like "*Status:*") {
                    $status = ($line -split ":")[1].Trim()
                }
                if ($line -like "*Last Run Time:*") {
                    $lastRun = ($line -split ":",2)[1].Trim()
                }
                if ($line -like "*Next Run Time:*") {
                    $nextRun = ($line -split ":",2)[1].Trim()
                }
            }
            
            # Check if process is actually running
            $isRunning = $false
            $processId = $null
            $runningProcess = Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue | 
                Where-Object { $_.State -eq "Listen" }
            if ($runningProcess) {
                $processId = $runningProcess.OwningProcess
                $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
                if ($process -and $process.ProcessName -eq "node") {
                    $isRunning = $true
                }
            }
            
            return @{
                TaskExists = $true
                Status = $status
                IsRunning = $isRunning
                ProcessId = $processId
                LastRun = $lastRun
                NextRun = $nextRun
            }
        }
    } catch {}
    
    return @{
        TaskExists = $false
        Status = "Not Found"
        IsRunning = $false
        ProcessId = $null
        LastRun = "N/A"
        NextRun = "N/A"
    }
}

function Start-FrontendTask {
    Write-Host "Starting frontend task..." -ForegroundColor Yellow
    schtasks /run /tn $TaskName
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Task start command sent successfully" -ForegroundColor Green
        
        # Wait a moment and check status
        Start-Sleep -Seconds 3
        $status = Get-TaskStatus
        if ($status.IsRunning) {
            Write-Host "Frontend process is now running on port $Port" -ForegroundColor Green
        } else {
            Write-Warning "Task started but process not detected on port $Port"
        }
    } else {
        Write-Error "Failed to start task"
    }
}

function Stop-FrontendTask {
    Write-Host "Stopping frontend task..." -ForegroundColor Yellow
    
    # End the scheduled task
    schtasks /end /tn $TaskName 2>$null | Out-Null
    
    # Also kill any Node.js processes on our port
    $runningProcess = Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue | 
        Where-Object { $_.State -eq "Listen" }
    if ($runningProcess) {
        $processId = $runningProcess.OwningProcess
        $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
        if ($process -and $process.ProcessName -eq "node") {
            Write-Host "Stopping Node.js process (PID: $processId)..." -ForegroundColor Yellow
            Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
        }
    }
    
    Start-Sleep -Seconds 2
    $status = Get-TaskStatus
    if (-not $status.IsRunning) {
        Write-Host "Frontend task stopped successfully" -ForegroundColor Green
    } else {
        Write-Warning "Task may still be running"
    }
}

function Show-TaskStatus {
    $status = Get-TaskStatus
    
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  ExcelAddin Frontend Task Status" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    if ($status.TaskExists) {
        Write-Host "Task Information:" -ForegroundColor Green
        Write-Host "  Name: $TaskName"
        Write-Host "  Status: $($status.Status)"
        Write-Host "  Process Running: $($status.IsRunning)"
        if ($status.ProcessId) {
            Write-Host "  Process ID: $($status.ProcessId)"
        }
        Write-Host "  Last Run: $($status.LastRun)"
        Write-Host "  Next Run: $($status.NextRun)"
        
        if ($Detailed) {
            Write-Host ""
            Write-Host "Detailed Task Information:" -ForegroundColor Cyan
            schtasks /query /tn $TaskName /fo LIST
        }
    } else {
        Write-Host "Task Status: NOT FOUND" -ForegroundColor Red
        Write-Host "The scheduled task '$TaskName' does not exist." -ForegroundColor Red
        Write-Host "Run the deployment script to create it." -ForegroundColor Yellow
    }
    
    Write-Host ""
    Write-Host "Port Information:" -ForegroundColor Green
    $portInfo = Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue | 
        Where-Object { $_.State -eq "Listen" }
    if ($portInfo) {
        $processId = $portInfo.OwningProcess
        $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
        Write-Host "  Port $Port: IN USE by $($process.ProcessName) (PID: $processId)"
    } else {
        Write-Host "  Port $Port: AVAILABLE"
    }
}

function Remove-FrontendTask {
    Write-Host "Removing frontend task..." -ForegroundColor Yellow
    
    # Stop first
    Stop-FrontendTask
    
    # Remove task
    schtasks /delete /tn $TaskName /f 2>$null | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Task removed successfully" -ForegroundColor Green
    } else {
        Write-Host "Task may not have existed or removal failed" -ForegroundColor Yellow
    }
}

# Main logic
Write-Host "ExcelAddin Frontend Task Manager" -ForegroundColor Green
Write-Host ""

switch ($Action.ToLower()) {
    "start" { Start-FrontendTask }
    "stop" { Stop-FrontendTask }
    "restart" { 
        Stop-FrontendTask
        Start-Sleep -Seconds 2
        Start-FrontendTask
    }
    "status" { Show-TaskStatus }
    "delete" { Remove-FrontendTask }
}

Write-Host ""