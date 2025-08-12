param(
    [string]$SiteName = "Default Web Site", 
    [string]$ApplicationName = "excellence",
    [string]$BackendAppName = "backend",
    [switch]$Force
)

Write-Host "=== Simple Backend Deployment ===" -ForegroundColor Green

# Variables
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$RepoRoot = Split-Path -Parent (Split-Path -Parent $ScriptPath)
$BackendPath = Join-Path $RepoRoot "backend"
$IISBackendPath = "C:\inetpub\wwwroot\$ApplicationName\$BackendAppName"
$FullBackendAppName = "$ApplicationName/$BackendAppName"

Write-Host "Repository: $RepoRoot"
Write-Host "Backend source: $BackendPath"  
Write-Host "IIS backend path: $IISBackendPath"
Write-Host "Full app name: $FullBackendAppName"

# Check if backend exists
if (-not (Test-Path $BackendPath)) {
    Write-Error "Backend not found at $BackendPath"
    exit 1
}

# Import IIS module
Import-Module WebAdministration -ErrorAction SilentlyContinue
if (-not (Get-Module WebAdministration)) {
    Write-Error "IIS WebAdministration module not available"
    exit 1
}

try {
    # Find Python
    $PythonPath = $null
    $PythonPaths = @(
        "python",
        "python3",
        "C:\Python*\python.exe",
        "C:\Program Files\Python*\python.exe",
        "C:\pyenv\pyenv-win\shims\python.bat"
    )
    
    foreach ($path in $PythonPaths) {
        try {
            if ($path -like "*\*") {
                $resolved = Get-ChildItem $path -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($resolved) {
                    $testPath = $resolved.FullName
                } else {
                    continue
                }
            } else {
                $testPath = $path
            }
            
            $version = & $testPath --version 2>&1
            if ($LASTEXITCODE -eq 0) {
                $PythonPath = $testPath
                Write-Host "Found Python: $version at $PythonPath" -ForegroundColor Green
                break
            }
        } catch {
            continue
        }
    }
    
    if (-not $PythonPath) {
        Write-Error "Python not found. Install Python 3.8+ and ensure it's in PATH."
        exit 1
    }
    
    # Install backend dependencies
    Write-Host "Installing backend dependencies..."
    Push-Location $BackendPath
    
    # Check if poetry is available
    try {
        & poetry --version 2>$null
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Installing dependencies with Poetry..."
            & poetry install --only=main
        } else {
            throw "Poetry not available"
        }
    } catch {
        Write-Host "Poetry not available, using pip..." -ForegroundColor Yellow
        # Install basic requirements with pip
        & $PythonPath -m pip install Flask==3.1.1 Flask-CORS==6.0.1 python-dotenv==1.1.1 requests==2.32.4 --quiet
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Failed to install Python dependencies"
            exit 1
        }
    }
    
    # Install wfastcgi
    Write-Host "Installing wfastcgi..."
    & $PythonPath -m pip install wfastcgi --quiet
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to install wfastcgi"
        exit 1
    }
    
    Pop-Location
    
    # Remove existing backend application
    if (Get-WebApplication -Site $SiteName -Name $FullBackendAppName -ErrorAction SilentlyContinue) {
        if ($Force) {
            Write-Host "Removing existing backend application: $FullBackendAppName" -ForegroundColor Yellow
            Remove-WebApplication -Site $SiteName -Name $FullBackendAppName
        } else {
            Write-Host "Backend application already exists. Use -Force to recreate." -ForegroundColor Yellow
            return
        }
    }
    
    # Create or update backend directory  
    if (Test-Path $IISBackendPath) {
        Write-Host "Cleaning existing backend directory..." -ForegroundColor Yellow
        Remove-Item $IISBackendPath -Recurse -Force
    }
    
    Write-Host "Creating backend directory: $IISBackendPath"
    New-Item -ItemType Directory -Path $IISBackendPath -Force | Out-Null
    
    # Copy backend files
    Write-Host "Copying backend files..."
    Copy-Item "$BackendPath\*" -Destination $IISBackendPath -Recurse -Force
    
    # Create web.config for backend
    $WebConfigContent = @"
<?xml version="1.0" encoding="UTF-8"?>
<configuration>
  <appSettings>
    <add key="PYTHONPATH" value="$IISBackendPath" />
    <add key="WSGI_HANDLER" value="wsgi_app.application" />
  </appSettings>
  <system.webServer>
    <handlers>
      <add name="PythonHandler" 
           path="*" 
           verb="*" 
           modules="FastCgiModule" 
           scriptProcessor="$PythonPath|$IISBackendPath\wfastcgi.py" 
           resourceType="Unspecified" 
           requireAccess="Script" />
    </handlers>
  </system.webServer>
</configuration>
"@
    
    $WebConfigPath = Join-Path $IISBackendPath "web.config"
    Write-Host "Creating web.config: $WebConfigPath"
    $WebConfigContent | Out-File -FilePath $WebConfigPath -Encoding UTF8
    
    # Create WSGI entry point
    $WSGIContent = @"
import sys
import os

# Add backend directory to Python path
backend_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, backend_dir)

# Set environment to staging
os.environ['FLASK_ENV'] = 'staging'

# Import and create Flask app
from app import create_app
application = create_app()

if __name__ == '__main__':
    application.run()
"@
    
    $WSGIPath = Join-Path $IISBackendPath "wsgi_app.py"
    Write-Host "Creating WSGI entry point: $WSGIPath"
    $WSGIContent | Out-File -FilePath $WSGIPath -Encoding UTF8
    
    # Configure FastCGI
    Write-Host "Configuring FastCGI..."
    $FastCGIPath = "$PythonPath|$IISBackendPath\wfastcgi.py"
    
    # Check if FastCGI application already exists
    $ExistingFastCGI = Get-WebConfiguration -Filter "system.webServer/fastCgi/application[@fullPath='$PythonPath'][@arguments='$IISBackendPath\wfastcgi.py']" -PSPath "MACHINE/WEBROOT/APPHOST"
    
    if (-not $ExistingFastCGI.fullPath) {
        Write-Host "Adding FastCGI application..."
        Add-WebConfiguration -Filter "system.webServer/fastCgi" -Value @{
            fullPath = $PythonPath;
            arguments = "$IISBackendPath\wfastcgi.py";
            maxInstances = 4;
            idleTimeout = 300;
            activityTimeout = 30;
            requestTimeout = 90;
            instanceMaxRequests = 10000;
            protocol = "NamedPipe";
            flushNamedPipe = $false
        }
    } else {
        Write-Host "FastCGI application already configured"
    }
    
    # Create IIS application
    Write-Host "Creating IIS backend application: $FullBackendAppName"
    New-WebApplication -Site $SiteName -Name $FullBackendAppName -PhysicalPath $IISBackendPath
    
    Write-Host "Backend deployment completed successfully!" -ForegroundColor Green
    Write-Host "Backend URL: https://localhost:9443/$ApplicationName/$BackendAppName/" -ForegroundColor Cyan
    Write-Host "Health check: https://localhost:9443/$ApplicationName/$BackendAppName/api/health" -ForegroundColor Cyan
    
} catch {
    Write-Error "Backend deployment failed: $_"
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    exit 1
}