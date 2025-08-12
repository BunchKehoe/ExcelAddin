param(
    [string]$SiteName = "Default Web Site",
    [string]$ApplicationName = "excellence",
    [switch]$Confirm = $false
)

Write-Host "=== Simple Deployment Cleanup ===" -ForegroundColor Yellow

if (-not $Confirm) {
    $response = Read-Host "This will remove the IIS applications and files. Continue? (y/N)"
    if ($response -ne 'y' -and $response -ne 'Y') {
        Write-Host "Cleanup cancelled." -ForegroundColor Green
        exit 0
    }
}

# Import IIS module
Import-Module WebAdministration -ErrorAction SilentlyContinue
if (-not (Get-Module WebAdministration)) {
    Write-Error "IIS WebAdministration module not available"
    exit 1
}

try {
    # Remove backend application
    $BackendAppName = "$ApplicationName/backend"
    if (Get-WebApplication -Site $SiteName -Name $BackendAppName -ErrorAction SilentlyContinue) {
        Write-Host "Removing backend application: $BackendAppName" -ForegroundColor Yellow
        Remove-WebApplication -Site $SiteName -Name $BackendAppName
        Write-Host "✓ Backend application removed" -ForegroundColor Green
    }
    
    # Remove frontend application
    if (Get-WebApplication -Site $SiteName -Name $ApplicationName -ErrorAction SilentlyContinue) {
        Write-Host "Removing frontend application: $ApplicationName" -ForegroundColor Yellow
        Remove-WebApplication -Site $SiteName -Name $ApplicationName
        Write-Host "✓ Frontend application removed" -ForegroundColor Green
    }
    
    # Remove files
    $IISPath = "C:\inetpub\wwwroot\$ApplicationName"
    if (Test-Path $IISPath) {
        Write-Host "Removing files: $IISPath" -ForegroundColor Yellow
        Remove-Item $IISPath -Recurse -Force
        Write-Host "✓ Files removed" -ForegroundColor Green
    }
    
    Write-Host "`n✓ Cleanup completed successfully!" -ForegroundColor Green
    Write-Host "The Excel add-in deployment has been completely removed." -ForegroundColor Cyan
    
} catch {
    Write-Error "Cleanup failed: $_"
    exit 1
}