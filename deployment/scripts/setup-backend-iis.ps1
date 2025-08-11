# PowerShell script to set up Excel Add-in Backend with IIS (replacing NSSM)
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

# Validate backend directory
if (-not (Test-Path $BackendPath)) {
    Write-Error "Backend directory not found: $BackendPath"
    exit 1
}

if (-not (Test-Path "$BackendPath\app.py")) {
    Write-Error "app.py not found in $BackendPath"
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

# Check if IIS is installed and running
$iisFeature = Get-WindowsFeature -Name IIS-WebServerRole -ErrorAction SilentlyContinue
if (-not $iisFeature -or $iisFeature.InstallState -ne "Installed") {
    Write-Error "IIS is not installed. Please install IIS with ASP.NET support first."
    Write-Host "Run: Enable-WindowsOptionalFeature -Online -FeatureName IIS-WebServerRole,IIS-ASPNET45" -ForegroundColor Yellow
    exit 1
}

Write-Host "[OK] IIS is installed and available" -ForegroundColor Green

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
    & $PythonPath -m wfastcgi.enable
    Write-Host "[OK] FastCGI enabled for Python" -ForegroundColor Green
}
catch {
    Write-Warning "FastCGI configuration may have failed, but continuing..."
}

if ($Uninstall) {
    Write-Host "Uninstall mode - removing IIS configuration..." -ForegroundColor Yellow
    
    # Remove the backend application if it exists
    $backendApp = Get-WebApplication -Name "backend" -Site "Default Web Site" -ErrorAction SilentlyContinue
    if ($backendApp) {
        Remove-WebApplication -Name "backend" -Site "Default Web Site"
        Write-Host "[OK] Removed backend web application" -ForegroundColor Green
    }
    
    Write-Host "Uninstall completed successfully!" -ForegroundColor Green
    exit 0
}

# Create IIS Application for backend
Write-Host "Creating IIS application for backend..." -ForegroundColor Yellow

# Remove existing application if it exists (when Force is used)
if ($Force) {
    $existingApp = Get-WebApplication -Name "backend" -Site "Default Web Site" -ErrorAction SilentlyContinue
    if ($existingApp) {
        Remove-WebApplication -Name "backend" -Site "Default Web Site"
        Write-Host "[INFO] Removed existing backend application" -ForegroundColor Yellow
    }
}

# Create the web application
try {
    New-WebApplication -Name "backend" -Site "Default Web Site" -PhysicalPath $BackendPath -ApplicationPool "DefaultAppPool"
    Write-Host "[OK] Created IIS application 'backend' at $BackendPath" -ForegroundColor Green
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
        # Update Python path in FastCGI configuration
        $webConfig = $webConfig -replace 'C:\\Python39\\python\.exe', $PythonPath
        $webConfig = $webConfig -replace 'C:\\Python39\\Lib\\site-packages\\wfastcgi\.py', "$((Split-Path $PythonPath -Parent)\Lib\site-packages\wfastcgi.py)"
        $webConfig | Set-Content $webConfigPath
        Write-Host "[OK] Updated web.config with Python path: $PythonPath" -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to update web.config: $($_.Exception.Message)"
    }
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

Write-Host ""
Write-Host "Setup completed successfully!" -ForegroundColor Green
Write-Host "=" * 50
Write-Host "Backend Configuration:" -ForegroundColor Cyan
Write-Host "• IIS Application: /backend"
Write-Host "• Physical Path: $BackendPath"
Write-Host "• Python Path: $PythonPath"
Write-Host "• WSGI Handler: wsgi_app.application"
Write-Host ""
Write-Host "Next Steps:" -ForegroundColor Yellow
Write-Host "1. Verify IIS can serve the backend at: http://localhost/backend/api/health"
Write-Host "2. Update the main web.config to route API calls to /backend/"
Write-Host "3. Test the Excel add-in with the new configuration"
Write-Host ""
Write-Host "Troubleshooting:" -ForegroundColor Yellow
Write-Host "• Check IIS logs at: C:\inetpub\logs\LogFiles\"
Write-Host "• Check Python errors in IIS Manager > Application > Failed Request Tracing"
Write-Host "• Test backend directly: python wsgi_app.py"