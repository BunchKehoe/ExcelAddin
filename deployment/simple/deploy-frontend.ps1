param(
    [string]$SiteName = "Default Web Site",
    [string]$ApplicationName = "excellence",
    [switch]$Force
)

Write-Host "=== Simple Frontend Deployment ===" -ForegroundColor Green

# Variables
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$RepoRoot = Split-Path -Parent (Split-Path -Parent $ScriptPath)
$BuildPath = Join-Path $RepoRoot "dist"
$IISPath = "C:\inetpub\wwwroot\$ApplicationName"

Write-Host "Repository: $RepoRoot"
Write-Host "Build path: $BuildPath"
Write-Host "IIS path: $IISPath"

# Check if build exists
if (-not (Test-Path $BuildPath)) {
    Write-Error "Build not found at $BuildPath. Run 'npm run build:staging' first."
    exit 1
}

Write-Host "Build found: $BuildPath"

# Import IIS module
Import-Module WebAdministration -ErrorAction SilentlyContinue
if (-not (Get-Module WebAdministration)) {
    Write-Error "IIS WebAdministration module not available"
    exit 1
}

try {
    # Remove existing application if Force is specified
    if ($Force -and (Get-WebApplication -Site $SiteName -Name $ApplicationName -ErrorAction SilentlyContinue)) {
        Write-Host "Removing existing application: $ApplicationName" -ForegroundColor Yellow
        Remove-WebApplication -Site $SiteName -Name $ApplicationName
    }
    
    # Create or update IIS directory
    if (Test-Path $IISPath) {
        Write-Host "Cleaning existing IIS directory..." -ForegroundColor Yellow
        Remove-Item $IISPath -Recurse -Force
    }
    
    Write-Host "Creating IIS directory: $IISPath"
    New-Item -ItemType Directory -Path $IISPath -Force | Out-Null
    
    # Copy build files
    Write-Host "Copying build files..."
    Copy-Item "$BuildPath\*" -Destination $IISPath -Recurse -Force
    
    # Create IIS application
    if (-not (Get-WebApplication -Site $SiteName -Name $ApplicationName -ErrorAction SilentlyContinue)) {
        Write-Host "Creating IIS application: $ApplicationName"
        New-WebApplication -Site $SiteName -Name $ApplicationName -PhysicalPath $IISPath
    } else {
        Write-Host "IIS application already exists: $ApplicationName"
    }
    
    Write-Host "Frontend deployment completed successfully!" -ForegroundColor Green
    Write-Host "URL: https://localhost:9443/$ApplicationName/" -ForegroundColor Cyan
    
} catch {
    Write-Error "Frontend deployment failed: $_"
    exit 1
}