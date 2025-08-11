<#
.SYNOPSIS
    Builds the React frontend and deploys to IIS
.DESCRIPTION
    Complete build and deployment script that:
    1. Builds the React application using webpack
    2. Copies files to the correct IIS directory structure
    3. Ensures proper configuration
.PARAMETER SiteName
    Name of the IIS site (default: ExcelAddin)
.EXAMPLE
    .\build-and-deploy-iis.ps1
    .\build-and-deploy-iis.ps1 -SiteName "MyExcelApp"
#>

param(
    [string]$SiteName = "ExcelAddin"
)

# Set error handling
$ErrorActionPreference = "Stop"

Write-Host "Building and deploying Excel Add-in to IIS..." -ForegroundColor Green

# Variables
$ProjectRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$WebsiteRoot = "C:\inetpub\wwwroot\$SiteName"
$ExcellenceDir = Join-Path $WebsiteRoot "excellence"
$DistDir = Join-Path $ProjectRoot "dist"

try {
    # Step 1: Check prerequisites
    Write-Host "1. Checking prerequisites..." -ForegroundColor Cyan
    
    # Check if we're in the correct directory
    if (-not (Test-Path (Join-Path $ProjectRoot "package.json"))) {
        throw "package.json not found. Please run this script from the project root or ensure the project structure is correct."
    }
    
    # Check if npm is available
    try {
        npm --version | Out-Null
        Write-Host "   npm is available" -ForegroundColor Green
    } catch {
        throw "npm is not available. Please install Node.js and npm."
    }
    
    # Check if IIS site exists
    try {
        Import-Module WebAdministration -ErrorAction Stop
        $site = Get-Website -Name $SiteName -ErrorAction SilentlyContinue
        if (-not $site) {
            throw "IIS site '$SiteName' not found. Please run deploy-to-existing-iis.ps1 first."
        }
        Write-Host "   IIS site '$SiteName' found" -ForegroundColor Green
    } catch {
        throw "IIS is not properly configured. Please run deploy-to-existing-iis.ps1 first."
    }

    # Step 2: Build the React application
    Write-Host "2. Building React application..." -ForegroundColor Cyan
    
    Push-Location $ProjectRoot
    try {
        # Clean previous build
        if (Test-Path $DistDir) {
            Remove-Item $DistDir -Recurse -Force
            Write-Host "   Cleaned previous build" -ForegroundColor Green
        }
        
        # Build the application
        Write-Host "   Running npm run build:staging..." -ForegroundColor Yellow
        npm run build:staging
        Write-Host "   Build completed successfully" -ForegroundColor Green
        
        # Verify build output
        if (-not (Test-Path $DistDir)) {
            throw "Build failed - dist directory not created"
        }
        
        $buildFiles = Get-ChildItem $DistDir -File
        if ($buildFiles.Count -eq 0) {
            throw "Build failed - no files in dist directory"
        }
        
        Write-Host "   Build created $($buildFiles.Count) files" -ForegroundColor Green
        
    } finally {
        Pop-Location
    }

    # Step 3: Deploy to IIS directory structure
    Write-Host "3. Deploying to IIS..." -ForegroundColor Cyan
    
    # Ensure target directories exist
    if (-not (Test-Path $WebsiteRoot)) {
        throw "IIS website root directory not found: $WebsiteRoot"
    }
    
    if (-not (Test-Path $ExcellenceDir)) {
        New-Item -ItemType Directory -Path $ExcellenceDir -Force | Out-Null
        Write-Host "   Created excellence directory: $ExcellenceDir" -ForegroundColor Green
    }
    
    # Copy all files from dist to excellence directory
    Write-Host "   Copying files from $DistDir to $ExcellenceDir..." -ForegroundColor Yellow
    Copy-Item -Path "$DistDir\*" -Destination $ExcellenceDir -Recurse -Force
    
    # Verify deployment
    $deployedFiles = Get-ChildItem $ExcellenceDir -Recurse -File
    Write-Host "   Deployed $($deployedFiles.Count) files to IIS" -ForegroundColor Green
    
    # Step 3.1: Update web.config (ensure latest version is deployed)
    Write-Host "   Updating web.config..." -ForegroundColor Yellow
    $webConfigSource = Join-Path $PSScriptRoot "..\iis\web.config"
    $webConfigDest = Join-Path $WebsiteRoot "web.config"
    
    if (Test-Path $webConfigSource) {
        Copy-Item $webConfigSource $webConfigDest -Force
        Write-Host "   [OK] Updated web.config to latest version" -ForegroundColor Green
    } else {
        Write-Warning "   web.config not found at: $webConfigSource"
    }
    
    # Check for key files
    $keyFiles = @("taskpane.html", "commands.html")
    foreach ($file in $keyFiles) {
        $filePath = Join-Path $ExcellenceDir $file
        if (Test-Path $filePath) {
            Write-Host "   [OK] $file deployed successfully" -ForegroundColor Green
        } else {
            Write-Warning "   [WARNING] $file not found in deployment"
        }
    }

    # Step 4: Set proper permissions
    Write-Host "4. Setting directory permissions..." -ForegroundColor Cyan
    
    try {
        $acl = Get-Acl $ExcellenceDir
        $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("IIS_IUSRS", "ReadAndExecute", "ContainerInherit,ObjectInherit", "None", "Allow")
        $acl.SetAccessRule($accessRule)
        $accessRule2 = New-Object System.Security.AccessControl.FileSystemAccessRule("IUSR", "ReadAndExecute", "ContainerInherit,ObjectInherit", "None", "Allow")
        $acl.SetAccessRule($accessRule2)
        Set-Acl -Path $ExcellenceDir -AclObject $acl
        Write-Host "   Directory permissions configured" -ForegroundColor Green
    } catch {
        Write-Warning "Could not set directory permissions: $($_.Exception.Message)"
    }

    # Step 5: Restart IIS site
    Write-Host "5. Restarting IIS site..." -ForegroundColor Cyan
    
    try {
        Stop-Website -Name $SiteName -ErrorAction SilentlyContinue
        Start-Website -Name $SiteName
        Write-Host "   IIS site restarted" -ForegroundColor Green
    } catch {
        Write-Warning "Could not restart IIS site: $($_.Exception.Message)"
    }

    # Success message
    Write-Host "`nDeployment completed successfully!" -ForegroundColor Green -BackgroundColor DarkGreen
    Write-Host "Website URL: https://server-vs81t.intranet.local:9443/excellence/" -ForegroundColor Cyan
    Write-Host "Taskpane: https://server-vs81t.intranet.local:9443/excellence/taskpane.html" -ForegroundColor Cyan
    Write-Host "Health check: https://server-vs81t.intranet.local:9443/health" -ForegroundColor Cyan
    
    Write-Host "`nNext steps:" -ForegroundColor Yellow
    Write-Host "1. Start Flask backend: cd backend && python app.py" -ForegroundColor White
    Write-Host "2. Test the deployment: .\deployment\scripts\test-iis-simple.ps1" -ForegroundColor White
    Write-Host "3. Load the add-in in Excel using the manifest file" -ForegroundColor White

} catch {
    Write-Error "Deployment failed: $($_.Exception.Message)"
    Write-Host "`nTroubleshooting tips:" -ForegroundColor Yellow
    Write-Host "1. Ensure you have Administrator privileges" -ForegroundColor White
    Write-Host "2. Check that IIS is properly configured with deploy-to-existing-iis.ps1" -ForegroundColor White
    Write-Host "3. Verify Node.js and npm are installed and working" -ForegroundColor White
    Write-Host "4. Check that all project files are present" -ForegroundColor White
    exit 1
}