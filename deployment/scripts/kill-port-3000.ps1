#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Kill processes using port 3000

.DESCRIPTION
    This script identifies and stops processes that are currently using port 3000.
    Useful for resolving port conflicts before deploying the frontend service.

.PARAMETER Force
    Force kill processes without confirmation

.EXAMPLE
    .\kill-port-3000.ps1
    .\kill-port-3000.ps1 -Force
#>

param(
    [switch]$Force
)

# Source the common functions
. "$PSScriptRoot\common.ps1"

Write-Header "Port 3000 Conflict Resolution"

try {
    $portConnections = Get-NetTCPConnection -LocalPort 3000 -ErrorAction SilentlyContinue
    
    if (-not $portConnections) {
        Write-Success "Port 3000 is not in use"
        return
    }
    
    Write-Host "Found processes using port 3000:"
    $processesToKill = @()
    
    foreach ($connection in $portConnections) {
        $process = Get-Process -Id $connection.OwningProcess -ErrorAction SilentlyContinue
        if ($process) {
            Write-Host "  Process: $($process.ProcessName) (PID: $($process.Id))"
            
            # Don't kill system processes (PID 0, 4) or critical Windows processes
            if ($process.Id -gt 4 -and $process.ProcessName -notin @("System", "Idle", "svchost", "winlogon", "csrss")) {
                $processesToKill += $process
            } else {
                Write-Warning "  Skipping system/critical process: $($process.ProcessName)"
            }
        } else {
            Write-Warning "  Could not identify process with PID: $($connection.OwningProcess)"
        }
    }
    
    if ($processesToKill.Count -eq 0) {
        Write-Warning "No killable processes found using port 3000"
        return
    }
    
    # Confirm before killing unless -Force is used
    if (-not $Force) {
        Write-Host ""
        Write-Host "The following processes will be terminated:"
        $processesToKill | ForEach-Object {
            Write-Host "  - $($_.ProcessName) (PID: $($_.Id))"
        }
        Write-Host ""
        $confirm = Read-Host "Continue? (y/N)"
        if ($confirm -notmatch '^y(es)?$') {
            Write-Host "Operation cancelled"
            return
        }
    }
    
    # Kill the processes
    Write-Host "Stopping processes..."
    foreach ($process in $processesToKill) {
        try {
            Write-Host "  Stopping $($process.ProcessName) (PID: $($process.Id))..."
            Stop-Process -Id $process.Id -Force -ErrorAction Stop
            Write-Success "    Successfully stopped $($process.ProcessName)"
        } catch {
            Write-Warning "    Failed to stop $($process.ProcessName): $($_.Exception.Message)"
        }
    }
    
    # Wait a moment and verify
    Start-Sleep -Seconds 2
    
    Write-Host ""
    Write-Host "Verifying port 3000 status..."
    $remainingConnections = Get-NetTCPConnection -LocalPort 3000 -ErrorAction SilentlyContinue
    if (-not $remainingConnections) {
        Write-Success "Port 3000 is now available"
    } else {
        Write-Warning "Port 3000 is still in use by some processes:"
        foreach ($connection in $remainingConnections) {
            $process = Get-Process -Id $connection.OwningProcess -ErrorAction SilentlyContinue
            if ($process) {
                Write-Warning "  Process: $($process.ProcessName) (PID: $($process.Id))"
            }
        }
    }
    
} catch {
    Write-Warning "Error checking port usage: $($_.Exception.Message)"
    Write-Host "This is normal on some Windows versions where Get-NetTCPConnection is not available."
}

Write-Host ""
Write-Host "Port conflict resolution complete."