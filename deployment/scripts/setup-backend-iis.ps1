# PowerShell script to set up Excel Add-in Backend with IIS (replacing NSSM)
# Usage: .\setup-backend-iis.ps1 -SiteName "ExcelAddin"

param(
    [Parameter(Mandatory=$false)]
    [string]$SiteName = "ExcelAddin",
    
    [Parameter(Mandatory=$false)]
    [string]$PythonPath = "",
    
    [Parameter(Mandatory=$false)]
    [switch]$Uninstall,
    
    [Parameter(Mandatory=$false)]
    [switch]$Force,
    
    [Parameter(Mandatory=$false)]
    [switch]$Debug
)

# Check if running as administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Administrator)) {
    Write-Error "This script must be run as Administrator"
    exit 1
}

Write-Host "Excel Add-in Backend IIS Setup (replacing NSSM)" -ForegroundColor Cyan
Write-Host "=" * 60

# Variables
$WebsiteRoot = "C:\inetpub\wwwroot\$SiteName"
$ExcellenceDir = Join-Path $WebsiteRoot "excellence"
$BackendPath = Join-Path $ExcellenceDir "backend"
$ProjectRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$SourceBackendPath = Join-Path $ProjectRoot "backend"

# Validate source backend directory
if (-not (Test-Path $SourceBackendPath)) {
    Write-Error "Source backend directory not found: $SourceBackendPath"
    exit 1
}

if (-not (Test-Path "$SourceBackendPath\app.py")) {
    Write-Error "app.py not found in $SourceBackendPath"
    exit 1
}

# Find Python executable
if (-not $PythonPath) {
    Write-Host "Finding Python installation..." -ForegroundColor Yellow
    
    $pythonCandidates = @(
        "C:\Python39\python.exe",
        "C:\Python310\python.exe", 
        "C:\Python311\python.exe",
        "C:\Python312\python.exe",
        "${env:LOCALAPPDATA}\Programs\Python\Python39\python.exe",
        "${env:LOCALAPPDATA}\Programs\Python\Python310\python.exe",
        "${env:LOCALAPPDATA}\Programs\Python\Python311\python.exe",
        "${env:LOCALAPPDATA}\Programs\Python\Python312\python.exe"
    )
    
    # Try candidates first
    foreach ($candidate in $pythonCandidates) {
        if (Test-Path $candidate) {
            $PythonPath = $candidate
            break
        }
    }
    
    # Fallback to PATH search
    if (-not $PythonPath) {
        try {
            $pythonFromPath = Get-Command python -ErrorAction Stop
            $PythonPath = $pythonFromPath.Source
        }
        catch {
            Write-Error "Python not found. Please install Python 3.9+ or specify -PythonPath"
            exit 1
        }
    }
}

if (-not (Test-Path $PythonPath)) {
    Write-Error "Python executable not found: $PythonPath"
    exit 1
}

# Verify Python version
try {
    $pythonVersion = & $PythonPath --version 2>$null
    Write-Host "[OK] Python found: $pythonVersion at $PythonPath" -ForegroundColor Green
}
catch {
    Write-Error "Failed to verify Python installation: $PythonPath"
    exit 1
}

# IIS check removed as requested - assuming IIS is available

# Install/Verify wfastcgi for Python-IIS integration
Write-Host "Setting up Python FastCGI integration..." -ForegroundColor Yellow
try {
    & $PythonPath -m pip install wfastcgi
    Write-Host "[OK] wfastcgi installed/updated" -ForegroundColor Green
}
catch {
    Write-Error "Failed to install wfastcgi. Please check Python and pip installation."
    exit 1
}

# Enable FastCGI in IIS
Write-Host "Configuring IIS FastCGI..." -ForegroundColor Yellow
try {
    # Try the enable command, but don't fail if the module structure has changed
    $fastcgiResult = & $PythonPath -m wfastcgi --help 2>$null
    if ($LASTEXITCODE -eq 0) {
        & $PythonPath -m wfastcgi enable 2>$null
        Write-Host "[OK] FastCGI enabled for Python" -ForegroundColor Green
    } else {
        Write-Host "[INFO] wfastcgi enable command not available - FastCGI may need manual configuration" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "[INFO] wfastcgi enable command failed - FastCGI may need manual configuration" -ForegroundColor Yellow
}

if ($Uninstall) {
    Write-Host "Uninstall mode - removing IIS configuration..." -ForegroundColor Yellow
    
    # Remove the backend application if it exists
    $backendApp = Get-WebApplication -Name "backend" -Site $SiteName -ErrorAction SilentlyContinue
    if ($backendApp) {
        Remove-WebApplication -Name "backend" -Site $SiteName
        Write-Host "[OK] Removed backend web application from site $SiteName" -ForegroundColor Green
    }
    
    # Also remove from Default Web Site if it exists there
    $defaultBackendApp = Get-WebApplication -Name "backend" -Site "Default Web Site" -ErrorAction SilentlyContinue
    if ($defaultBackendApp) {
        Remove-WebApplication -Name "backend" -Site "Default Web Site"
        Write-Host "[OK] Removed backend web application from Default Web Site" -ForegroundColor Green
    }
    
    # Remove backend directory
    if (Test-Path $BackendPath) {
        Remove-Item $BackendPath -Recurse -Force
        Write-Host "[OK] Removed backend directory: $BackendPath" -ForegroundColor Green
    }
    
    Write-Host "Uninstall completed successfully!" -ForegroundColor Green
    exit 0
}

# Create IIS Application for backend
Write-Host "Creating IIS backend application and deploying files..." -ForegroundColor Yellow

# Ensure the website exists
try {
    $website = Get-Website -Name $SiteName -ErrorAction SilentlyContinue
    if (-not $website) {
        Write-Error "IIS website '$SiteName' not found. Please run deploy-to-existing-iis.ps1 first."
        exit 1
    }
    Write-Host "[OK] Found IIS website: $SiteName" -ForegroundColor Green
}
catch {
    Write-Error "Failed to check IIS website: $($_.Exception.Message)"
    exit 1
}

# Ensure excellence directory exists
if (-not (Test-Path $ExcellenceDir)) {
    New-Item -ItemType Directory -Path $ExcellenceDir -Force | Out-Null
    Write-Host "[OK] Created excellence directory: $ExcellenceDir" -ForegroundColor Green
}

# Remove existing backend if Force is used
if ($Force) {
    $existingApp = Get-WebApplication -Name "backend" -Site $SiteName -ErrorAction SilentlyContinue
    if ($existingApp) {
        Remove-WebApplication -Name "backend" -Site $SiteName
        Write-Host "[INFO] Removed existing backend application" -ForegroundColor Yellow
    }
    
    if (Test-Path $BackendPath) {
        Remove-Item $BackendPath -Recurse -Force
        Write-Host "[INFO] Removed existing backend directory" -ForegroundColor Yellow
    }
}

# Create backend directory and copy files
if (-not (Test-Path $BackendPath)) {
    New-Item -ItemType Directory -Path $BackendPath -Force | Out-Null
}

Write-Host "Copying backend files from $SourceBackendPath to $BackendPath..." -ForegroundColor Yellow
try {
    # Copy all backend files except __pycache__ and other temp files
    $excludeItems = @("__pycache__", "*.pyc", ".pytest_cache", "venv", ".venv", "node_modules")
    robocopy $SourceBackendPath $BackendPath /E /XD $excludeItems /XF "*.pyc" /NFL /NDL /NJH /NJS
    
    # Verify key files were copied
    $keyFiles = @("app.py", "wsgi_app.py", "web.config", "pyproject.toml")
    foreach ($file in $keyFiles) {
        $filePath = Join-Path $BackendPath $file
        if (Test-Path $filePath) {
            Write-Host "[OK] $file copied successfully" -ForegroundColor Green
        } elseif ($file -eq "pyproject.toml" -or $file -eq "requirements.txt") {
            # Optional files
            Write-Host "[INFO] $file not found (optional)" -ForegroundColor Yellow
        } else {
            Write-Warning "Required file missing: $file"
        }
    }
}
catch {
    Write-Error "Failed to copy backend files: $($_.Exception.Message)"
    exit 1
}

# Create the IIS web application for backend
try {
    # Check if there's an existing application and remove it if Force is used
    $virtualPath = "/excellence/backend"
    $existingApp = Get-WebApplication -Name "backend" -Site $SiteName -ErrorAction SilentlyContinue
    if ($existingApp) {
        if ($Force) {
            Remove-WebApplication -Name "backend" -Site $SiteName
            Write-Host "[INFO] Removed existing backend application (Force mode)" -ForegroundColor Yellow
        } else {
            Write-Warning "Backend application already exists. Use -Force to overwrite."
            Write-Host "[INFO] Skipping application creation, will only update files and config" -ForegroundColor Yellow
        }
    }
    
    # Create new application if it doesn't exist
    if (-not (Get-WebApplication -Name "backend" -Site $SiteName -ErrorAction SilentlyContinue)) {
        New-WebApplication -Name "backend" -Site $SiteName -PhysicalPath $BackendPath -ApplicationPool "DefaultAppPool"
        Write-Host "[OK] Created IIS application 'backend' at $BackendPath" -ForegroundColor Green
        Write-Host "[OK] Virtual path: $virtualPath" -ForegroundColor Green
    }
}
catch {
    Write-Error "Failed to create IIS application: $($_.Exception.Message)"
    exit 1
}

# Install Python dependencies
Write-Host "Installing Python dependencies..." -ForegroundColor Yellow
try {
    Push-Location $BackendPath
    
    # Install from pyproject.toml if it exists, otherwise from requirements.txt
    if (Test-Path "pyproject.toml") {
        & $PythonPath -m pip install .
        Write-Host "[OK] Installed dependencies from pyproject.toml" -ForegroundColor Green
    }
    elseif (Test-Path "requirements.txt") {
        & $PythonPath -m pip install -r requirements.txt
        Write-Host "[OK] Installed dependencies from requirements.txt" -ForegroundColor Green
    }
    else {
        Write-Warning "No pyproject.toml or requirements.txt found - installing basic dependencies"
        & $PythonPath -m pip install flask flask-cors python-dotenv
    }
}
catch {
    Write-Warning "Failed to install some dependencies: $($_.Exception.Message)"
}
finally {
    Pop-Location
}

# Update web.config with correct Python path
Write-Host "Updating web.config with Python path..." -ForegroundColor Yellow
$webConfigPath = Join-Path $BackendPath "web.config"
if (Test-Path $webConfigPath) {
    try {
        $webConfig = Get-Content $webConfigPath -Raw
        # Update Python path in FastCGI configuration - use proper path escaping
        $webConfig = $webConfig -replace 'C:\\Python39\\python\.exe', $PythonPath
        $pythonLibPath = (Split-Path $PythonPath -Parent) + '\Lib\site-packages\wfastcgi.py'
        $webConfig = $webConfig -replace 'C:\\Python39\\Lib\\site-packages\\wfastcgi\.py', $pythonLibPath
        
        # Update PYTHONPATH to point to the correct backend directory
        $webConfig = $webConfig -replace 'C:\\inetpub\\wwwroot\\ExcelAddin\\backend', $BackendPath.Replace('\', '\\')
        
        $webConfig | Set-Content $webConfigPath
        Write-Host "[OK] Updated web.config with Python path: $PythonPath" -ForegroundColor Green
        Write-Host "[OK] Updated PYTHONPATH to: $BackendPath" -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to update web.config: $($_.Exception.Message)"
    }
} else {
    Write-Warning "web.config not found at: $webConfigPath - backend may not work correctly"
}

# Test the backend application
Write-Host "Testing backend application..." -ForegroundColor Yellow
try {
    Push-Location $BackendPath
    & $PythonPath -c "from app import create_app; app = create_app(); print('Backend app created successfully')"
    Write-Host "[OK] Backend application loads successfully" -ForegroundColor Green
}
catch {
    Write-Warning "Backend application test failed: $($_.Exception.Message)"
}
finally {
    Pop-Location
}

# Set proper permissions for backend directory
Write-Host "Setting directory permissions..." -ForegroundColor Yellow
try {
    $acl = Get-Acl $BackendPath
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("IIS_IUSRS", "ReadAndExecute", "ContainerInherit,ObjectInherit", "None", "Allow")
    $acl.SetAccessRule($accessRule)
    $accessRule2 = New-Object System.Security.AccessControl.FileSystemAccessRule("IUSR", "ReadAndExecute", "ContainerInherit,ObjectInherit", "None", "Allow")
    $acl.SetAccessRule($accessRule2)
    Set-Acl -Path $BackendPath -AclObject $acl
    Write-Host "[OK] Directory permissions configured for IIS" -ForegroundColor Green
} catch {
    Write-Warning "Could not set directory permissions: $($_.Exception.Message)"
}

Write-Host ""
Write-Host "Setup completed successfully!" -ForegroundColor Green
Write-Host "=" * 50
Write-Host "Backend Configuration:" -ForegroundColor Cyan
Write-Host "• IIS Website: $SiteName"
Write-Host "• Backend Application: /excellence/backend"
Write-Host "• Physical Path: $BackendPath"
Write-Host "• Python Path: $PythonPath"
Write-Host "• WSGI Handler: wsgi_app.application"
Write-Host ""
Write-Host "Architecture:" -ForegroundColor Yellow
Write-Host "• Frontend: https://yourserver/$SiteName/excellence/"
Write-Host "• API Calls: https://yourserver/$SiteName/excellence/api/* → /excellence/backend/api/*"
Write-Host "• Health Check: https://yourserver/$SiteName/excellence/backend/api/health"
Write-Host ""
Write-Host "Next Steps:" -ForegroundColor Yellow
Write-Host "1. Run the full deployment: .\build-and-deploy-iis.ps1 -SiteName '$SiteName'"
Write-Host "2. Test the API: Test-WebRequest https://yourserver/$SiteName/excellence/api/health"
Write-Host "3. Load the Excel add-in using the manifest file"
Write-Host ""
Write-Host "Troubleshooting:" -ForegroundColor Yellow
Write-Host "• Check IIS logs at: C:\inetpub\logs\LogFiles\"
Write-Host "• Check application in IIS Manager: Sites → $SiteName → backend"
Write-Host "• Verify routing in main web.config: /excellence/api/* → /excellence/backend/api/*"