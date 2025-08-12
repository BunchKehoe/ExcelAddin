# PowerShell script to test the IIS backend integration
# Usage: .\test-backend-integration.ps1 -SiteName "ExcelAddin"

param(
    [Parameter(Mandatory=$false)]
    [string]$SiteName = "ExcelAddin",
    
    [Parameter(Mandatory=$false)]
    [string]$ServerUrl = "https://localhost"
)

Write-Host "Testing Excel Add-in Backend IIS Integration" -ForegroundColor Cyan
Write-Host "=" * 50

$WebsiteRoot = "C:\inetpub\wwwroot\$SiteName"
$ExcellenceDir = Join-Path $WebsiteRoot "excellence"
$BackendPath = Join-Path $ExcellenceDir "backend"

# Test 1: Check if IIS website exists
Write-Host "1. Checking IIS website..." -ForegroundColor Yellow
try {
    Import-Module WebAdministration -ErrorAction Stop
    $website = Get-Website -Name $SiteName -ErrorAction SilentlyContinue
    if ($website) {
        Write-Host "[OK] IIS website '$SiteName' exists" -ForegroundColor Green
        Write-Host "    Physical Path: $($website.PhysicalPath)" -ForegroundColor Gray
        Write-Host "    State: $($website.State)" -ForegroundColor Gray
    } else {
        Write-Host "[ERROR] IIS website '$SiteName' not found" -ForegroundColor Red
        exit 1
    }
} catch {
    Write-Host "[ERROR] Failed to check IIS website: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Test 2: Check if backend application exists
Write-Host "2. Checking backend application..." -ForegroundColor Yellow
try {
    $backendApp = Get-WebApplication -Name "backend" -Site $SiteName -ErrorAction SilentlyContinue
    if ($backendApp) {
        Write-Host "[OK] Backend application exists" -ForegroundColor Green
        Write-Host "    Virtual Path: $($backendApp.Path)" -ForegroundColor Gray
        Write-Host "    Physical Path: $($backendApp.PhysicalPath)" -ForegroundColor Gray
        Write-Host "    Application Pool: $($backendApp.ApplicationPool)" -ForegroundColor Gray
    } else {
        Write-Host "[ERROR] Backend application not found" -ForegroundColor Red
        Write-Host "    Run: .\setup-backend-iis.ps1 -SiteName '$SiteName'" -ForegroundColor Yellow
        exit 1
    }
} catch {
    Write-Host "[ERROR] Failed to check backend application: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Test 3: Check backend files
Write-Host "3. Checking backend files..." -ForegroundColor Yellow
$requiredFiles = @("app.py", "wsgi_app.py", "web.config")
$missingFiles = @()

foreach ($file in $requiredFiles) {
    $filePath = Join-Path $BackendPath $file
    if (Test-Path $filePath) {
        Write-Host "[OK] $file exists" -ForegroundColor Green
    } else {
        Write-Host "[ERROR] $file missing" -ForegroundColor Red
        $missingFiles += $file
    }
}

if ($missingFiles.Count -gt 0) {
    Write-Host "[ERROR] Missing required files. Run setup-backend-iis.ps1 again" -ForegroundColor Red
    exit 1
}

# Test 4: Check web.config configuration
Write-Host "4. Checking web.config configuration..." -ForegroundColor Yellow
$webConfigPath = Join-Path $BackendPath "web.config"
if (Test-Path $webConfigPath) {
    try {
        $webConfig = Get-Content $webConfigPath -Raw
        if ($webConfig -match 'WSGI_HANDLER.*wsgi_app\.application') {
            Write-Host "[OK] WSGI handler configured correctly" -ForegroundColor Green
        } else {
            Write-Host "[WARNING] WSGI handler may not be configured correctly" -ForegroundColor Yellow
        }
        
        if ($webConfig -match 'PYTHONPATH') {
            Write-Host "[OK] PYTHONPATH configured" -ForegroundColor Green
        } else {
            Write-Host "[WARNING] PYTHONPATH not found in configuration" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "[ERROR] Failed to read web.config: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "[ERROR] web.config not found" -ForegroundColor Red
    exit 1
}

# Test 5: Check main website routing
Write-Host "5. Checking main website routing..." -ForegroundColor Yellow
$mainWebConfigPath = Join-Path $WebsiteRoot "web.config"
if (Test-Path $mainWebConfigPath) {
    try {
        $mainWebConfig = Get-Content $mainWebConfigPath -Raw
        if ($mainWebConfig -match 'excellence/api.*excellence/backend/api') {
            Write-Host "[OK] API routing configured: /excellence/api/* → /excellence/backend/api/*" -ForegroundColor Green
        } else {
            Write-Host "[WARNING] API routing may not be configured correctly" -ForegroundColor Yellow
            Write-Host "    Expected: /excellence/api/* → /excellence/backend/api/*" -ForegroundColor Gray
        }
    } catch {
        Write-Host "[ERROR] Failed to read main web.config: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "[WARNING] Main web.config not found at: $mainWebConfigPath" -ForegroundColor Yellow
}

# Test 6: Check Python dependencies
Write-Host "6. Checking Python dependencies..." -ForegroundColor Yellow
try {
    Push-Location $BackendPath
    
    # Try to import the main app
    $pythonTest = python -c "
try:
    from app import create_app
    app = create_app()
    print('SUCCESS: Backend app imports successfully')
except ImportError as e:
    print(f'ERROR: Import failed: {e}')
    exit(1)
except Exception as e:
    print(f'ERROR: App creation failed: {e}')
    exit(1)
" 2>&1
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "[OK] Python backend imports successfully" -ForegroundColor Green
    } else {
        Write-Host "[ERROR] Python backend test failed:" -ForegroundColor Red
        Write-Host "    $pythonTest" -ForegroundColor Red
    }
} catch {
    Write-Host "[ERROR] Failed to test Python backend: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    Pop-Location
}

Write-Host ""
Write-Host "Integration Test Summary:" -ForegroundColor Cyan
Write-Host "• Architecture: Single IIS site with frontend at /excellence/ and backend at /excellence/backend/" -ForegroundColor White
Write-Host "• API Routing: /excellence/api/* should route to /excellence/backend/api/*" -ForegroundColor White
Write-Host "• No localhost:5000 service - everything runs through IIS" -ForegroundColor White
Write-Host ""
Write-Host "Next Steps:" -ForegroundColor Yellow
Write-Host "1. Test API endpoint: Invoke-WebRequest '$ServerUrl/$SiteName/excellence/api/health' (if website is running)" -ForegroundColor White
Write-Host "2. Check IIS Manager for both applications under '$SiteName'" -ForegroundColor White
Write-Host "3. Review IIS logs if there are any issues" -ForegroundColor White