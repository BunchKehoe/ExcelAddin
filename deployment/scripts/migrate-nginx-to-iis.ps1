<#
.SYNOPSIS
    Migrates Excel Add-in from nginx to IIS
.DESCRIPTION
    This script stops nginx services and migrates the Excel Add-in to IIS
.EXAMPLE
    .\migrate-nginx-to-iis.ps1 -Force
#>

param(
    [switch]$Force = $false
)

# Check if running as Administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script must be run as Administrator. Right-click PowerShell and select 'Run as Administrator'"
    exit 1
}

Write-Host "Migrating Excel Add-in from nginx to IIS..." -ForegroundColor Green

try {
    # Step 1: Stop nginx services
    Write-Host "1. Stopping nginx services..." -ForegroundColor Cyan
    
    # Stop nginx service if it exists
    $nginxService = Get-Service -Name "nginx" -ErrorAction SilentlyContinue
    if ($nginxService) {
        if ($nginxService.Status -eq 'Running') {
            Stop-Service -Name "nginx" -Force
            Write-Host "   Stopped nginx service" -ForegroundColor Green
        }
        
        if ($Force) {
            # Remove nginx service using NSSM
            try {
                & nssm remove nginx confirm
                Write-Host "   Removed nginx service" -ForegroundColor Green
            } catch {
                Write-Warning "Could not remove nginx service: $($_.Exception.Message)"
            }
        }
    } else {
        Write-Host "   nginx service not found (already removed or not installed)" -ForegroundColor Yellow
    }

    # Step 2: Backup nginx configuration (optional)
    Write-Host "2. Backing up nginx configuration..." -ForegroundColor Cyan
    
    $nginxConfigPath = "C:\nginx\conf\excel-addin.conf"
    if (Test-Path $nginxConfigPath) {
        $backupPath = "C:\nginx\conf\excel-addin.conf.backup.$(Get-Date -Format 'yyyyMMdd-HHmmss')"
        Copy-Item $nginxConfigPath $backupPath
        Write-Host "   nginx config backed up to: $backupPath" -ForegroundColor Green
    } else {
        Write-Host "   nginx config not found, skipping backup" -ForegroundColor Yellow
    }

    # Step 3: Run IIS setup
    Write-Host "3. Setting up IIS..." -ForegroundColor Cyan
    
    $setupScript = Join-Path $PSScriptRoot "setup-iis.ps1"
    if (Test-Path $setupScript) {
        & $setupScript -Force:$Force
        Write-Host "   IIS setup completed" -ForegroundColor Green
    } else {
        Write-Error "IIS setup script not found: $setupScript"
        exit 1
    }

    # Step 4: Copy existing frontend files if they exist in nginx location
    Write-Host "4. Migrating frontend files..." -ForegroundColor Cyan
    
    $nginxDistPath = "C:\inetpub\wwwroot\ExcelAddin\dist"
    $iisDistPath = "C:\inetpub\wwwroot\ExcelAddin\dist"
    
    if (Test-Path $nginxDistPath) {
        Write-Host "   Frontend files already exist in correct location" -ForegroundColor Green
    } else {
        Write-Host "   Please build and deploy frontend files:" -ForegroundColor Yellow
        Write-Host "     npm run build:staging" -ForegroundColor White
        Write-Host "     Copy dist/* to: $iisDistPath" -ForegroundColor White
    }

    # Step 5: Update firewall rules
    Write-Host "5. Updating firewall rules..." -ForegroundColor Cyan
    
    # Remove old nginx rules
    Get-NetFirewallRule -DisplayName "*nginx*" -ErrorAction SilentlyContinue | Remove-NetFirewallRule
    Write-Host "   Removed nginx firewall rules" -ForegroundColor Green
    
    # IIS rules are handled by setup-iis.ps1
    Write-Host "   IIS firewall rules configured by setup script" -ForegroundColor Green

    # Step 6: Test the new setup
    Write-Host "6. Testing IIS setup..." -ForegroundColor Cyan
    
    $testScript = Join-Path $PSScriptRoot "test-iis-simple.ps1"
    if (Test-Path $testScript) {
        Write-Host "   Running IIS validation tests..." -ForegroundColor White
        & $testScript
    } else {
        Write-Warning "IIS test script not found: $testScript"
    }

    Write-Host "`nMigration completed!" -ForegroundColor Green -BackgroundColor DarkGreen
    Write-Host "`nKey differences between nginx and IIS:" -ForegroundColor Yellow
    Write-Host "• Service: nginx service → IIS (W3SVC)" -ForegroundColor White
    Write-Host "• Configuration: nginx.conf → web.config" -ForegroundColor White  
    Write-Host "• Management: NSSM → IIS Manager or PowerShell" -ForegroundColor White
    Write-Host "• Same URL: https://server-vs81t.intranet.local:9443/excellence/" -ForegroundColor White
    Write-Host "`nManaging IIS:" -ForegroundColor Yellow
    Write-Host "• Start: Start-Website -Name ExcelAddin" -ForegroundColor White
    Write-Host "• Stop: Stop-Website -Name ExcelAddin" -ForegroundColor White
    Write-Host "• Status: Get-Website -Name ExcelAddin" -ForegroundColor White
    Write-Host "• IIS Manager: Run 'inetmgr' as Administrator" -ForegroundColor White

} catch {
    Write-Error "Migration failed: $($_.Exception.Message)"
    Write-Host "You may need to:" -ForegroundColor Red
    Write-Host "1. Manually stop nginx processes" -ForegroundColor White
    Write-Host "2. Install IIS and required modules manually" -ForegroundColor White
    Write-Host "3. Run setup-iis.ps1 separately" -ForegroundColor White
    exit 1
}