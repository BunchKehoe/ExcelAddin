<#
.SYNOPSIS
    Fixes IIS web.config configuration issues
.DESCRIPTION
    This script addresses common IIS web.config issues including:
    - HTTP Error 500.19 due to locked configuration sections (handlers)
    - Duplicate MIME type entries
    - Ensures the latest compatible web.config is deployed
.PARAMETER SiteName
    Name of the IIS site (default: ExcelAddin)
.PARAMETER Force
    Force replacement of existing web.config
.EXAMPLE
    .\fix-iis-config.ps1
    .\fix-iis-config.ps1 -SiteName "MyExcelApp" -Force
#>

param(
    [string]$SiteName = "ExcelAddin",
    [switch]$Force
)

# Set error handling
$ErrorActionPreference = "Stop"

Write-Host "Fixing IIS web.config configuration issues..." -ForegroundColor Green

try {
    # Check if running as Administrator
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "This script must be run as Administrator"
    }

    # Variables
    $WebsiteRoot = "C:\inetpub\wwwroot\$SiteName"
    $currentWebConfig = Join-Path $WebsiteRoot "web.config"
    $newWebConfigSource = Join-Path $PSScriptRoot "..\iis\web.config"
    $backupWebConfig = Join-Path $WebsiteRoot "web.config.backup"

    # Step 1: Check if IIS site exists
    Write-Host "1. Checking IIS site configuration..." -ForegroundColor Cyan
    
    try {
        Import-Module WebAdministration -ErrorAction Stop
        $site = Get-Website -Name $SiteName -ErrorAction SilentlyContinue
        if (-not $site) {
            throw "IIS site '$SiteName' not found"
        }
        Write-Host "   ✓ IIS site '$SiteName' found" -ForegroundColor Green
    } catch {
        throw "IIS configuration error: $($_.Exception.Message)"
    }

    # Step 2: Backup existing web.config
    Write-Host "2. Backing up current web.config..." -ForegroundColor Cyan
    
    if (Test-Path $currentWebConfig) {
        if (Test-Path $backupWebConfig) {
            if ($Force) {
                Remove-Item $backupWebConfig -Force
                Write-Host "   Removed existing backup" -ForegroundColor Yellow
            } else {
                Write-Host "   Backup already exists: $backupWebConfig" -ForegroundColor Green
                Write-Host "   Use -Force to overwrite existing backup" -ForegroundColor Yellow
            }
        }
        
        if (-not (Test-Path $backupWebConfig)) {
            Copy-Item $currentWebConfig $backupWebConfig
            Write-Host "   ✓ Backed up to: $backupWebConfig" -ForegroundColor Green
        }
    } else {
        Write-Host "   No existing web.config found" -ForegroundColor Yellow
    }

    # Step 3: Check for problematic configurations
    Write-Host "3. Checking for configuration issues..." -ForegroundColor Cyan
    
    $hasIssues = $false
    $issueDetails = @()
    
    if (Test-Path $currentWebConfig) {
        $configContent = Get-Content $currentWebConfig -Raw
        
        # Check for handlers section (most common issue)
        if ($configContent -match '<handlers>') {
            $hasIssues = $true
            $issueDetails += "- Contains <handlers> section (causes 500.19 error)"
            Write-Host "   ⚠ Found <handlers> section (locked at server level)" -ForegroundColor Red
        }
        
        # Check for duplicate MIME types
        if ($configContent -match 'fileExtension="\.js"') {
            $hasIssues = $true
            $issueDetails += "- Contains duplicate .js MIME type"
            Write-Host "   ⚠ Found duplicate .js MIME type definition" -ForegroundColor Red
        }
        
        if ($configContent -match 'fileExtension="\.json"') {
            $hasIssues = $true
            $issueDetails += "- Contains duplicate .json MIME type"
            Write-Host "   ⚠ Found duplicate .json MIME type definition" -ForegroundColor Red
        }
        
        if (-not $hasIssues) {
            Write-Host "   ✓ No obvious configuration issues detected" -ForegroundColor Green
        }
    }

    # Step 4: Deploy fixed web.config
    Write-Host "4. Deploying fixed web.config..." -ForegroundColor Cyan
    
    if (-not (Test-Path $newWebConfigSource)) {
        throw "Fixed web.config not found at: $newWebConfigSource"
    }
    
    if ($hasIssues -or $Force -or -not (Test-Path $currentWebConfig)) {
        Copy-Item $newWebConfigSource $currentWebConfig -Force
        Write-Host "   ✓ Deployed fixed web.config" -ForegroundColor Green
        
        # Verify the new configuration
        $newConfigContent = Get-Content $currentWebConfig -Raw
        if ($newConfigContent -notmatch '<handlers>') {
            Write-Host "   ✓ Verified: No problematic <handlers> section" -ForegroundColor Green
        } else {
            Write-Warning "   ⚠ Warning: handlers section still present"
        }
    } else {
        Write-Host "   Current web.config appears to be fine" -ForegroundColor Green
        Write-Host "   Use -Force to replace anyway" -ForegroundColor Yellow
    }

    # Step 5: Test IIS configuration
    Write-Host "5. Testing IIS configuration..." -ForegroundColor Cyan
    
    try {
        # Test configuration syntax
        $appCmd = "$env:SystemRoot\System32\inetsrv\appcmd.exe"
        if (Test-Path $appCmd) {
            $testResult = & $appCmd list site $SiteName 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Host "   ✓ IIS configuration syntax is valid" -ForegroundColor Green
            } else {
                Write-Warning "   IIS configuration may have issues: $testResult"
            }
        }
        
        # Try to restart the site
        Stop-Website -Name $SiteName -ErrorAction SilentlyContinue
        Start-Website -Name $SiteName -ErrorAction SilentlyContinue
        
        $siteState = (Get-Website -Name $SiteName).State
        if ($siteState -eq "Started") {
            Write-Host "   ✓ IIS site restarted successfully" -ForegroundColor Green
        } else {
            Write-Warning "   IIS site state: $siteState"
        }
    } catch {
        Write-Warning "Could not fully test IIS configuration: $($_.Exception.Message)"
    }

    # Summary
    Write-Host "`nConfiguration fix completed!" -ForegroundColor Green -BackgroundColor DarkGreen
    
    if ($hasIssues) {
        Write-Host "`nIssues that were fixed:" -ForegroundColor Yellow
        foreach ($issue in $issueDetails) {
            Write-Host $issue -ForegroundColor White
        }
    }
    
    Write-Host "`nNext steps:" -ForegroundColor Yellow
    Write-Host "1. Test the website: https://server-vs81t.intranet.local:9443/excellence/" -ForegroundColor White
    Write-Host "2. If still having issues, check IIS logs: %SystemDrive%\inetpub\logs\LogFiles\" -ForegroundColor White
    Write-Host "3. Restore backup if needed: Copy '$backupWebConfig' to '$currentWebConfig'" -ForegroundColor White

} catch {
    Write-Error "Configuration fix failed: $($_.Exception.Message)"
    Write-Host "`nTroubleshooting:" -ForegroundColor Yellow
    Write-Host "1. Ensure you have Administrator privileges" -ForegroundColor White
    Write-Host "2. Check that the IIS site exists and is properly configured" -ForegroundColor White
    Write-Host "3. Verify the web.config source file exists in deployment/iis/" -ForegroundColor White
    Write-Host "4. Check Windows Event Logs for more details" -ForegroundColor White
    exit 1
}