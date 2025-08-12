# DEPRECATED: This script has been replaced with setup-backend-iis.ps1
# PowerShell script redirecting to new IIS-based backend setup
# Usage: .\setup-backend-iis.ps1 -BackendPath "C:\inetpub\wwwroot\ExcelAddin\backend"

param(
    [Parameter(Mandatory=$false)]
    [string]$BackendPath = "C:\inetpub\wwwroot\ExcelAddin\backend",
    
    [Parameter(Mandatory=$false)]
    [string]$PythonPath = "",
    
    [Parameter(Mandatory=$false)]
    [switch]$Uninstall,
    
    [Parameter(Mandatory=$false)]
    [switch]$Force,
    
    [Parameter(Mandatory=$false)]
    [switch]$Debug
)

Write-Host "DEPRECATED: This script has been replaced" -ForegroundColor Yellow
Write-Host "Please use the new IIS-based backend setup instead:" -ForegroundColor Yellow
Write-Host ".\setup-backend-iis.ps1 -BackendPath `"$BackendPath`"" -ForegroundColor Cyan
Write-Host ""
Write-Host "The new approach hosts the backend directly in IIS," -ForegroundColor Green
Write-Host "eliminating the need for NSSM (Non-Sucking Service Manager)." -ForegroundColor Green
Write-Host ""
Write-Host "Press Enter to run the new script, or Ctrl+C to cancel..." -ForegroundColor Yellow
Read-Host

# Redirect to the new script
$scriptDir = Split-Path $PSCommandPath -Parent
$newScriptPath = Join-Path $scriptDir "setup-backend-iis.ps1"
if (Test-Path $newScriptPath) {
    & $newScriptPath @PSBoundParameters
} else {
    Write-Error "New script not found at: $newScriptPath"
    Write-Host "Please run: .\deployment\scripts\setup-backend-iis.ps1" -ForegroundColor Yellow
}