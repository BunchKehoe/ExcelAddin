<#
.SYNOPSIS
    Diagnose backend IIS configuration issues
.DESCRIPTION
    Detailed diagnosis script to identify backend deployment and routing issues
.PARAMETER SiteName
    Name of the IIS site (default: ExcelAddin)
.EXAMPLE
    .\diagnose-backend-iis.ps1
    .\diagnose-backend-iis.ps1 -SiteName "MyExcelApp"
#>

param(
    [string]$SiteName = "ExcelAddin"
)

Write-Host "Diagnosing Backend IIS Configuration..." -ForegroundColor Green
Write-Host "=" * 60

$WebsiteRoot = "C:\inetpub\wwwroot\$SiteName"
$ExcellenceDir = Join-Path $WebsiteRoot "excellence"
$BackendPath = Join-Path $ExcellenceDir "backend"

# Diagnosis 1: Check directory structure
Write-Host "1. Directory Structure:" -ForegroundColor Cyan
Write-Host "   Website Root: $WebsiteRoot"
if (Test-Path $WebsiteRoot) {
    Write-Host "   [OK] Website root exists" -ForegroundColor Green
} else {
    Write-Host "   [ERROR] Website root does not exist" -ForegroundColor Red
}

Write-Host "   Excellence Dir: $ExcellenceDir"
if (Test-Path $ExcellenceDir) {
    $excellenceFiles = Get-ChildItem $ExcellenceDir -File
    Write-Host "   [OK] Excellence directory exists with $($excellenceFiles.Count) files" -ForegroundColor Green
} else {
    Write-Host "   [ERROR] Excellence directory does not exist" -ForegroundColor Red
}

Write-Host "   Backend Dir: $BackendPath"
if (Test-Path $BackendPath) {
    $backendFiles = Get-ChildItem $BackendPath -File
    Write-Host "   [OK] Backend directory exists with $($backendFiles.Count) files" -ForegroundColor Green
    
    # Check for key backend files
    $keyFiles = @("app.py", "wsgi_app.py", "web.config")
    foreach ($file in $keyFiles) {
        $filePath = Join-Path $BackendPath $file
        if (Test-Path $filePath) {
            Write-Host "   [OK] $file found" -ForegroundColor Green
        } else {
            Write-Host "   [ERROR] $file missing" -ForegroundColor Red
        }
    }
} else {
    Write-Host "   [ERROR] Backend directory does not exist" -ForegroundColor Red
}

# Diagnosis 2: Check IIS applications
Write-Host "`n2. IIS Applications:" -ForegroundColor Cyan
try {
    Import-Module WebAdministration -ErrorAction Stop
    
    $site = Get-Website -Name $SiteName -ErrorAction SilentlyContinue
    if ($site) {
        Write-Host "   [OK] Website '$SiteName' exists (State: $($site.State))" -ForegroundColor Green
        
        # Check for applications
        $apps = Get-WebApplication -Site $SiteName
        Write-Host "   Applications in site:" -ForegroundColor Yellow
        foreach ($app in $apps) {
            Write-Host "     - Name: $($app.Name), Path: $($app.Path), Physical: $($app.PhysicalPath)" -ForegroundColor White
        }
        
        # Specifically check for backend application
        $backendApp = $apps | Where-Object { $_.Path -like "*backend*" -or $_.Name -like "*backend*" }
        if ($backendApp) {
            Write-Host "   [OK] Backend application found: $($backendApp.Name) at $($backendApp.Path)" -ForegroundColor Green
        } else {
            Write-Host "   [ERROR] No backend application found in IIS" -ForegroundColor Red
        }
    } else {
        Write-Host "   [ERROR] Website '$SiteName' not found" -ForegroundColor Red
    }
} catch {
    Write-Host "   [ERROR] Cannot access IIS: $($_.Exception.Message)" -ForegroundColor Red
}

# Diagnosis 3: Check web.config routing
Write-Host "`n3. Web.config Routing:" -ForegroundColor Cyan
$webConfigPath = Join-Path $WebsiteRoot "web.config"
if (Test-Path $webConfigPath) {
    Write-Host "   [OK] Main web.config exists" -ForegroundColor Green
    
    try {
        $webConfigContent = Get-Content $webConfigPath -Raw
        if ($webConfigContent -match 'excellence/api/\(\.\*\)') {
            Write-Host "   [OK] API routing rule found: excellence/api/(.*)" -ForegroundColor Green
        } else {
            Write-Host "   [WARNING] API routing rule not found in web.config" -ForegroundColor Yellow
        }
        
        if ($webConfigContent -match 'excellence/backend/api/') {
            Write-Host "   [OK] Backend routing target found: excellence/backend/api/" -ForegroundColor Green
        } else {
            Write-Host "   [WARNING] Backend routing target not found in web.config" -ForegroundColor Yellow
        }
    } catch {
        Write-Host "   [ERROR] Cannot parse web.config: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "   [ERROR] Main web.config not found at: $webConfigPath" -ForegroundColor Red
}

$backendWebConfigPath = Join-Path $BackendPath "web.config"
if (Test-Path $backendWebConfigPath) {
    Write-Host "   [OK] Backend web.config exists" -ForegroundColor Green
} else {
    Write-Host "   [ERROR] Backend web.config not found at: $backendWebConfigPath" -ForegroundColor Red
}

# Diagnosis 4: Check Python and FastCGI configuration
Write-Host "`n4. Python and FastCGI:" -ForegroundColor Cyan
if (Test-Path $backendWebConfigPath) {
    try {
        $backendConfig = Get-Content $backendWebConfigPath -Raw
        
        # Extract Python path from web.config
        if ($backendConfig -match 'fullPath="([^"]+python\.exe)"') {
            $pythonPath = $matches[1]
            Write-Host "   Python Path in config: $pythonPath" -ForegroundColor White
            
            if (Test-Path $pythonPath) {
                try {
                    $pythonVersion = & $pythonPath --version 2>$null
                    Write-Host "   [OK] Python executable found: $pythonVersion" -ForegroundColor Green
                } catch {
                    Write-Host "   [ERROR] Python executable not working: $($_.Exception.Message)" -ForegroundColor Red
                }
            } else {
                Write-Host "   [ERROR] Python executable not found at: $pythonPath" -ForegroundColor Red
            }
        } else {
            Write-Host "   [WARNING] Cannot find Python path in backend web.config" -ForegroundColor Yellow
        }
        
        # Check wfastcgi path
        if ($backendConfig -match 'arguments="([^"]*wfastcgi\.py[^"]*)"') {
            $wfastcgiPath = $matches[1]
            Write-Host "   wfastcgi Path in config: $wfastcgiPath" -ForegroundColor White
            
            if (Test-Path $wfastcgiPath) {
                Write-Host "   [OK] wfastcgi found" -ForegroundColor Green
            } else {
                Write-Host "   [ERROR] wfastcgi not found at: $wfastcgiPath" -ForegroundColor Red
            }
        }
        
    } catch {
        Write-Host "   [ERROR] Cannot parse backend web.config: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Diagnosis 5: Test Python app directly
Write-Host "`n5. Python App Test:" -ForegroundColor Cyan
if (Test-Path $BackendPath) {
    try {
        Push-Location $BackendPath
        
        # Try to import the Flask app
        $testScript = @"
import sys
import os
sys.path.insert(0, os.getcwd())

try:
    from app import create_app
    app = create_app()
    print("SUCCESS: Flask app created successfully")
    print("Routes available:")
    for rule in app.url_map.iter_rules():
        print(f"  {rule.methods} {rule.rule}")
except Exception as e:
    print(f"ERROR: {str(e)}")
    import traceback
    traceback.print_exc()
"@
        
        $testScriptPath = Join-Path $env:TEMP "test_backend.py"
        $testScript | Set-Content $testScriptPath
        
        # Find Python executable
        $pythonCandidates = @(
            "C:\Python39\python.exe",
            "C:\Python310\python.exe", 
            "C:\Python311\python.exe",
            "C:\Python312\python.exe",
            "python"
        )
        
        $pythonPath = $null
        foreach ($candidate in $pythonCandidates) {
            try {
                & $candidate --version 2>$null | Out-Null
                $pythonPath = $candidate
                break
            } catch {
                continue
            }
        }
        
        if ($pythonPath) {
            Write-Host "   Testing with Python: $pythonPath" -ForegroundColor Yellow
            $result = & $pythonPath $testScriptPath 2>&1
            Write-Host "   Test Result:" -ForegroundColor White
            $result | ForEach-Object { Write-Host "   $_" -ForegroundColor Gray }
        } else {
            Write-Host "   [ERROR] No Python executable found for testing" -ForegroundColor Red
        }
        
        Remove-Item $testScriptPath -ErrorAction SilentlyContinue
        
    } finally {
        Pop-Location
    }
}

# Summary and Recommendations
Write-Host "`n" + "=" * 60
Write-Host "SUMMARY AND RECOMMENDATIONS:" -ForegroundColor Yellow

Write-Host "`nBased on the diagnosis above, check the following:" -ForegroundColor White
Write-Host "1. Ensure backend directory exists with all required files" -ForegroundColor Cyan
Write-Host "2. Verify IIS application is created at the correct virtual path" -ForegroundColor Cyan
Write-Host "3. Check web.config routing rules are properly configured" -ForegroundColor Cyan
Write-Host "4. Ensure Python path in backend web.config is correct" -ForegroundColor Cyan
Write-Host "5. Verify Flask app can be imported and created successfully" -ForegroundColor Cyan

Write-Host "`nTo fix issues, try:" -ForegroundColor White
Write-Host "1. Re-run backend setup: .\setup-backend-iis.ps1 -SiteName $SiteName -Force" -ForegroundColor Green
Write-Host "2. Re-run full deployment: .\build-and-deploy-iis.ps1 -SiteName $SiteName" -ForegroundColor Green
Write-Host "3. Test the result: .\test-iis-simple.ps1" -ForegroundColor Green

Write-Host "`n" + "=" * 60