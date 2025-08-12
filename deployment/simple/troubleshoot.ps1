param(
    [string]$ApplicationName = "excellence"
)

Write-Host "=== Simple Deployment Troubleshooting ===" -ForegroundColor Cyan

$IISPath = "C:\inetpub\wwwroot\$ApplicationName"
$BackendPath = "$IISPath\backend"

# Check IIS Applications
Write-Host "`n1. Checking IIS Applications..." -ForegroundColor Yellow
Import-Module WebAdministration -ErrorAction SilentlyContinue

if (Get-Module WebAdministration) {
    $FrontendApp = Get-WebApplication -Site "Default Web Site" -Name $ApplicationName -ErrorAction SilentlyContinue
    $BackendApp = Get-WebApplication -Site "Default Web Site" -Name "$ApplicationName/backend" -ErrorAction SilentlyContinue
    
    if ($FrontendApp) {
        Write-Host "✓ Frontend application exists: $ApplicationName" -ForegroundColor Green
        Write-Host "  Path: $($FrontendApp.PhysicalPath)" -ForegroundColor Gray
    } else {
        Write-Host "✗ Frontend application missing: $ApplicationName" -ForegroundColor Red
    }
    
    if ($BackendApp) {
        Write-Host "✓ Backend application exists: $ApplicationName/backend" -ForegroundColor Green  
        Write-Host "  Path: $($BackendApp.PhysicalPath)" -ForegroundColor Gray
    } else {
        Write-Host "✗ Backend application missing: $ApplicationName/backend" -ForegroundColor Red
    }
} else {
    Write-Host "✗ IIS WebAdministration module not available" -ForegroundColor Red
}

# Check Files
Write-Host "`n2. Checking File Structure..." -ForegroundColor Yellow

if (Test-Path $IISPath) {
    Write-Host "✓ Frontend directory exists: $IISPath" -ForegroundColor Green
    
    $IndexFile = Join-Path $IISPath "taskpane.html"
    if (Test-Path $IndexFile) {
        Write-Host "✓ Frontend built correctly (taskpane.html found)" -ForegroundColor Green
    } else {
        Write-Host "✗ Frontend not built (taskpane.html missing)" -ForegroundColor Red
        Write-Host "  Run: npm run build:staging" -ForegroundColor Yellow
    }
    
    $ManifestFile = Join-Path $IISPath "manifest.xml"
    if (Test-Path $ManifestFile) {
        Write-Host "✓ Manifest file exists" -ForegroundColor Green
    } else {
        Write-Host "✗ Manifest file missing" -ForegroundColor Red
    }
} else {
    Write-Host "✗ Frontend directory missing: $IISPath" -ForegroundColor Red
}

if (Test-Path $BackendPath) {
    Write-Host "✓ Backend directory exists: $BackendPath" -ForegroundColor Green
    
    $AppFile = Join-Path $BackendPath "app.py"
    $WSGIFile = Join-Path $BackendPath "wsgi_app.py"
    $WebConfigFile = Join-Path $BackendPath "web.config"
    
    if (Test-Path $AppFile) {
        Write-Host "✓ Backend app.py exists" -ForegroundColor Green
    } else {
        Write-Host "✗ Backend app.py missing" -ForegroundColor Red
    }
    
    if (Test-Path $WSGIFile) {
        Write-Host "✓ Backend WSGI entry point exists" -ForegroundColor Green
    } else {
        Write-Host "✗ Backend WSGI entry point missing" -ForegroundColor Red
    }
    
    if (Test-Path $WebConfigFile) {
        Write-Host "✓ Backend web.config exists" -ForegroundColor Green
    } else {
        Write-Host "✗ Backend web.config missing" -ForegroundColor Red
    }
} else {
    Write-Host "✗ Backend directory missing: $BackendPath" -ForegroundColor Red
}

# Check Python
Write-Host "`n3. Checking Python Environment..." -ForegroundColor Yellow

$PythonPaths = @("python", "python3", "C:\pyenv\pyenv-win\shims\python.bat")
$PythonFound = $false

foreach ($pythonCmd in $PythonPaths) {
    try {
        $version = & $pythonCmd --version 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ Python found: $version at $pythonCmd" -ForegroundColor Green
            $PythonFound = $true
            
            # Check Flask
            $flaskCheck = & $pythonCmd -c "import flask; print(f'Flask {flask.__version__}')" 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ Flask available: $flaskCheck" -ForegroundColor Green
            } else {
                Write-Host "✗ Flask not available" -ForegroundColor Red
                Write-Host "  Install: pip install Flask Flask-CORS python-dotenv" -ForegroundColor Yellow
            }
            
            # Check wfastcgi
            $wfastcgiCheck = & $pythonCmd -c "import wfastcgi; print('wfastcgi available')" 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ wfastcgi available" -ForegroundColor Green
            } else {
                Write-Host "✗ wfastcgi not available" -ForegroundColor Red  
                Write-Host "  Install: pip install wfastcgi" -ForegroundColor Yellow
            }
            break
        }
    } catch {
        continue
    }
}

if (-not $PythonFound) {
    Write-Host "✗ Python not found" -ForegroundColor Red
    Write-Host "  Install Python 3.8+ and ensure it's in PATH" -ForegroundColor Yellow
}

# Check URLs
Write-Host "`n4. Testing URLs..." -ForegroundColor Yellow

$BaseUrl = "https://localhost:9443/$ApplicationName"
$FrontendUrl = "$BaseUrl/"
$BackendUrl = "$BaseUrl/backend/api/health"

function Test-SimpleUrl {
    param([string]$Url)
    try {
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
        $request = [System.Net.WebRequest]::Create($Url)
        $request.Timeout = 10000
        $response = $request.GetResponse()
        $statusCode = $response.StatusCode
        $response.Close()
        return $statusCode
    } catch {
        return $_.Exception.Message
    }
}

$frontendResult = Test-SimpleUrl -Url $FrontendUrl
if ($frontendResult -eq "OK") {
    Write-Host "✓ Frontend accessible: $FrontendUrl" -ForegroundColor Green
} else {
    Write-Host "✗ Frontend not accessible: $frontendResult" -ForegroundColor Red
}

$backendResult = Test-SimpleUrl -Url $BackendUrl  
if ($backendResult -eq "OK") {
    Write-Host "✓ Backend API accessible: $BackendUrl" -ForegroundColor Green
} else {
    Write-Host "✗ Backend API not accessible: $backendResult" -ForegroundColor Red
}

# Summary
Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host "If issues found above:" -ForegroundColor Yellow
Write-Host "1. Run: .\deployment\simple\deploy-all.ps1 -Force" -ForegroundColor Yellow
Write-Host "2. Check Windows Event Logs (Application)" -ForegroundColor Yellow
Write-Host "3. Check IIS Manager for detailed error messages" -ForegroundColor Yellow
Write-Host "4. Use: .\deployment\simple\uninstall.ps1 to clean up and start over" -ForegroundColor Yellow