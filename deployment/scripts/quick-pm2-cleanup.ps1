# One-time PM2 Cleanup Script
# Quick cleanup script that can be run independently

Write-Host "ExcelAddin PM2 Cleanup" -ForegroundColor Cyan
Write-Host "=====================" -ForegroundColor Cyan

# Check if PM2 is even installed
try {
    $pm2Version = pm2 --version 2>$null
    if (-not $pm2Version) {
        Write-Host "PM2 is not installed - cleanup not needed" -ForegroundColor Green
        exit 0
    }
} catch {
    Write-Host "PM2 is not installed - cleanup not needed" -ForegroundColor Green
    exit 0
}

Write-Host "Found PM2 version: $pm2Version" -ForegroundColor Yellow

# Kill PM2 daemon and stop all processes
Write-Host "Stopping all PM2 processes..." -ForegroundColor Yellow
try {
    pm2 kill 2>$null
    Write-Host "PM2 processes stopped" -ForegroundColor Green
} catch {
    Write-Host "PM2 kill command failed, continuing..." -ForegroundColor Yellow
}

# Remove PM2 config directory
$pm2Dir = "$env:USERPROFILE\.pm2"
if (Test-Path $pm2Dir) {
    Write-Host "Removing PM2 configuration directory..." -ForegroundColor Yellow
    try {
        Remove-Item -Path $pm2Dir -Recurse -Force
        Write-Host "PM2 configuration directory removed" -ForegroundColor Green
    } catch {
        Write-Host "Warning: Could not remove PM2 directory: $($_.Exception.Message)" -ForegroundColor Yellow
    }
} else {
    Write-Host "PM2 configuration directory not found" -ForegroundColor Green
}

Write-Host ""
Write-Host "PM2 cleanup completed successfully!" -ForegroundColor Green
Write-Host "The system is ready for NSSM-based frontend deployment." -ForegroundColor Green