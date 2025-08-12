# PM2 Cleanup Script
# Removes PM2 applications and configurations from the system

param(
    [switch]$RemovePM2,    # Also uninstall PM2 entirely
    [switch]$Force         # Skip confirmations
)

# Import common functions
. "$PSScriptRoot\common.ps1"

Write-Header "PM2 Cleanup Script"

$AppName = "exceladdin-frontend"

# Function to safely run PM2 commands
function Invoke-PM2Command {
    param(
        [string]$Command,
        [string]$Description
    )
    
    try {
        Write-Host "Executing: $Description"
        Invoke-Expression $Command
        if ($LASTEXITCODE -eq 0) {
            Write-Success "$Description completed successfully"
            return $true
        } else {
            Write-Warning "$Description failed with exit code $LASTEXITCODE"
            return $false
        }
    } catch {
        Write-Warning "$Description failed: $($_.Exception.Message)"
        return $false
    }
}

# Check if PM2 is installed
try {
    $pm2Version = pm2 --version 2>$null
    if (-not $pm2Version) {
        Write-Success "PM2 is not installed - no cleanup needed"
        exit 0
    }
    Write-Host "Found PM2 version: $pm2Version"
} catch {
    Write-Success "PM2 is not installed - no cleanup needed"
    exit 0
}

Write-Header "Cleaning Up PM2 Applications"

# Stop and remove specific ExcelAddin frontend application
Write-Host "Looking for ExcelAddin PM2 applications..."

try {
    $pm2List = pm2 jlist | ConvertFrom-Json
    $exceladdinApps = $pm2List | Where-Object { $_.name -like "*exceladdin*" -or $_.name -like "*ExcelAddin*" }
    
    if ($exceladdinApps) {
        foreach ($app in $exceladdinApps) {
            Write-Host "Found PM2 application: $($app.name)"
            
            if (-not $Force) {
                $confirm = Read-Host "Remove PM2 application '$($app.name)'? (y/n)"
                if ($confirm -ne 'y' -and $confirm -ne 'Y') {
                    Write-Host "Skipping $($app.name)"
                    continue
                }
            }
            
            # Stop the application
            Invoke-PM2Command -Command "pm2 stop $($app.name)" -Description "Stopping PM2 application $($app.name)"
            
            # Delete the application
            Invoke-PM2Command -Command "pm2 delete $($app.name)" -Description "Deleting PM2 application $($app.name)"
        }
    } else {
        Write-Host "No ExcelAddin PM2 applications found"
    }
} catch {
    Write-Warning "Error listing PM2 applications: $($_.Exception.Message)"
    Write-Host "Attempting individual cleanup commands..."
    
    # Try to stop and delete by name anyway
    Invoke-PM2Command -Command "pm2 stop $AppName" -Description "Stopping PM2 application $AppName"
    Invoke-PM2Command -Command "pm2 delete $AppName" -Description "Deleting PM2 application $AppName"
}

Write-Header "Cleaning Up PM2 Configuration"

# Remove PM2 startup configuration
Write-Host "Removing PM2 startup configuration..."
Invoke-PM2Command -Command "pm2 unstartup" -Description "Removing PM2 startup configuration"

# Clear PM2 save file
Write-Host "Clearing PM2 save file..."
Invoke-PM2Command -Command "pm2 kill" -Description "Killing PM2 daemon"

# Remove PM2 configuration files if they exist
$pm2Dir = "$env:USERPROFILE\.pm2"
if (Test-Path $pm2Dir) {
    if (-not $Force) {
        $confirm = Read-Host "Remove PM2 configuration directory '$pm2Dir'? (y/n)"
        if ($confirm -eq 'y' -or $confirm -eq 'Y') {
            try {
                Remove-Item -Path $pm2Dir -Recurse -Force
                Write-Success "PM2 configuration directory removed"
            } catch {
                Write-Warning "Failed to remove PM2 configuration directory: $($_.Exception.Message)"
            }
        }
    } else {
        try {
            Remove-Item -Path $pm2Dir -Recurse -Force
            Write-Success "PM2 configuration directory removed"
        } catch {
            Write-Warning "Failed to remove PM2 configuration directory: $($_.Exception.Message)"
        }
    }
}

Write-Header "Cleaning Up PM2 Config Files"

# Remove PM2 configuration files from deployment directory
$pm2ConfigPath = "$PSScriptRoot\..\config\pm2-frontend.json"
if (Test-Path $pm2ConfigPath) {
    if (-not $Force) {
        $confirm = Read-Host "Remove PM2 config file '$pm2ConfigPath'? (y/n)"
        if ($confirm -eq 'y' -or $confirm -eq 'Y') {
            Remove-Item -Path $pm2ConfigPath -Force
            Write-Success "PM2 configuration file removed"
        }
    } else {
        Remove-Item -Path $pm2ConfigPath -Force
        Write-Success "PM2 configuration file removed"
    }
}

# Remove old serve-frontend.js wrapper
$oldWrapperPath = "$PSScriptRoot\..\config\serve-frontend.js"
if (Test-Path $oldWrapperPath) {
    if (-not $Force) {
        $confirm = Read-Host "Remove old frontend wrapper '$oldWrapperPath'? (y/n)"
        if ($confirm -eq 'y' -or $confirm -eq 'Y') {
            Remove-Item -Path $oldWrapperPath -Force
            Write-Success "Old frontend wrapper removed"
        }
    } else {
        Remove-Item -Path $oldWrapperPath -Force
        Write-Success "Old frontend wrapper removed"
    }
}

# Optionally uninstall PM2 entirely
if ($RemovePM2) {
    Write-Header "Uninstalling PM2"
    
    if (-not $Force) {
        $confirm = Read-Host "Completely uninstall PM2 from the system? This will affect all PM2 applications. (y/n)"
        if ($confirm -ne 'y' -and $confirm -ne 'Y') {
            Write-Host "Skipping PM2 uninstallation"
        } else {
            Write-Host "Uninstalling PM2..."
            npm uninstall -g pm2
            if ($LASTEXITCODE -eq 0) {
                Write-Success "PM2 uninstalled successfully"
            } else {
                Write-Warning "PM2 uninstallation may have failed"
            }
        }
    } else {
        Write-Host "Uninstalling PM2..."
        npm uninstall -g pm2
        if ($LASTEXITCODE -eq 0) {
            Write-Success "PM2 uninstalled successfully"
        } else {
            Write-Warning "PM2 uninstallation may have failed"
        }
    }
}

Write-Header "PM2 Cleanup Complete"
Write-Success "PM2 cleanup has been completed successfully"
Write-Host ""
Write-Host "Summary of actions taken:"
Write-Host "  - Stopped and removed ExcelAddin PM2 applications"
Write-Host "  - Removed PM2 startup configuration"
Write-Host "  - Cleared PM2 save file and killed daemon"
if ($RemovePM2) {
    Write-Host "  - Uninstalled PM2 from the system"
}
Write-Host ""
Write-Host "The system is now ready for NSSM-based frontend deployment"