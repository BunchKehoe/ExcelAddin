<#
.SYNOPSIS
    Tests IIS setup for Excel Add-in
.DESCRIPTION
    Simple validation script to check IIS configuration and connectivity
.EXAMPLE
    .\test-iis-simple.ps1
#>

param(
    [string]$ServerName = "server-vs81t.intranet.local",
    [int]$Port = 9443,
    [switch]$Detailed = $false
)

Write-Host "Testing IIS setup for Excel Add-in..." -ForegroundColor Green
Write-Host "Target: https://$ServerName`:$Port/excellence/" -ForegroundColor Cyan

$testsPassed = 0
$totalTests = 0

function Test-Step {
    param([string]$Description, [scriptblock]$Test)
    $totalTests++
    Write-Host "$totalTests. $Description..." -ForegroundColor Cyan
    try {
        $result = & $Test
        if ($result) {
            Write-Host "   [PASSED] $result" -ForegroundColor Green
            $script:testsPassed++
        } else {
            Write-Host "   [FAILED] Test returned false or null" -ForegroundColor Red
        }
    } catch {
        Write-Host "   [FAILED] $($_.Exception.Message)" -ForegroundColor Red
        if ($Detailed) {
            Write-Host "   Details: $($_.Exception.ToString())" -ForegroundColor Gray
        }
    }
}

# Test 1: Check if IIS is running
Test-Step "Checking IIS service" {
    $service = Get-Service W3SVC -ErrorAction SilentlyContinue
    if ($service -and $service.Status -eq 'Running') {
        return "IIS service is running"
    } else {
        throw "IIS service is not running or not installed"
    }
}

# Test 2: Check if Excel Add-in site exists
Test-Step "Checking Excel Add-in website" {
    Import-Module WebAdministration -ErrorAction Stop
    $site = Get-Website -Name "ExcelAddin" -ErrorAction SilentlyContinue
    if ($site) {
        return "ExcelAddin website exists (State: $($site.State))"
    } else {
        throw "ExcelAddin website not found"
    }
}

# Test 3: Check if port 9443 is listening
Test-Step "Checking port $Port listening" {
    $listening = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
    if ($listening) {
        return "Port $Port is listening"
    } else {
        throw "Port $Port is not listening"
    }
}

# Test 4: Check physical path exists
Test-Step "Checking physical path" {
    $physicalPath = "C:\inetpub\wwwroot\ExcelAddin"
    $excellencePath = "C:\inetpub\wwwroot\ExcelAddin\excellence"
    
    if (Test-Path $physicalPath) {
        if (Test-Path $excellencePath) {
            $fileCount = (Get-ChildItem $excellencePath -Recurse -File).Count
            return "Physical path exists with excellence subdirectory ($fileCount files)"
        } else {
            throw "Physical path exists but excellence subdirectory is missing: $excellencePath"
        }
    } else {
        throw "Physical path does not exist: $physicalPath"
    }
}

# Test 5: Check web.config exists
Test-Step "Checking web.config" {
    $webConfigPath = "C:\inetpub\wwwroot\ExcelAddin\web.config"
    if (Test-Path $webConfigPath) {
        return "web.config exists in root directory"
    } else {
        throw "web.config not found at: $webConfigPath"
    }
}

# Test 6: Check SSL certificate
Test-Step "Checking SSL certificate binding" {
    Import-Module WebAdministration -ErrorAction Stop
    $binding = Get-WebBinding -Name "ExcelAddin" -Protocol "https" -ErrorAction SilentlyContinue
    if ($binding) {
        return "HTTPS binding exists on port $($binding.bindingInformation)"
    } else {
        throw "HTTPS binding not found for ExcelAddin site"
    }
}

# Test 7: Check if required IIS modules are available
Test-Step "Checking IIS modules" {
    $rewriteModule = Get-WindowsFeature -Name "IIS-HttpRedirect" -ErrorAction SilentlyContinue
    if ($rewriteModule -and $rewriteModule.InstallState -eq "Installed") {
        return "IIS modules are available"
    } else {
        throw "Required IIS modules may not be installed"
    }
}

# Test 8: Test health check endpoint
Test-Step "Testing health check endpoint" {
    try {
        # PowerShell version compatibility
        if ($PSVersionTable.PSVersion.Major -ge 6) {
            $response = Invoke-WebRequest -Uri "https://$ServerName`:$Port/health" -SkipCertificateCheck -UseBasicParsing -TimeoutSec 10
        } else {
            # Windows PowerShell 5.1 compatibility
            [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
            $response = Invoke-WebRequest -Uri "https://$ServerName`:$Port/health" -UseBasicParsing -TimeoutSec 10
            [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
        }
        
        if ($response.StatusCode -eq 200) {
            return "Health check responded with status $($response.StatusCode)"
        } else {
            throw "Health check returned status $($response.StatusCode)"
        }
    } catch {
        throw "Health check failed: $($_.Exception.Message)"
    }
}

# Test 9: Test frontend file
Test-Step "Testing frontend access" {
    try {
        # PowerShell version compatibility
        if ($PSVersionTable.PSVersion.Major -ge 6) {
            $response = Invoke-WebRequest -Uri "https://$ServerName`:$Port/excellence/taskpane.html" -SkipCertificateCheck -UseBasicParsing -TimeoutSec 10
        } else {
            # Windows PowerShell 5.1 compatibility
            [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
            $response = Invoke-WebRequest -Uri "https://$ServerName`:$Port/excellence/taskpane.html" -UseBasicParsing -TimeoutSec 10
            [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
        }
        
        if ($response.StatusCode -eq 200) {
            return "Frontend taskpane.html accessible (status $($response.StatusCode))"
        } else {
            throw "Frontend returned status $($response.StatusCode)"
        }
    } catch {
        throw "Frontend access failed: $($_.Exception.Message)"
    }
}

# Test 10: Test backend API through IIS routing
Test-Step "Testing backend API routing" {
    try {
        # Test the backend health endpoint through IIS routing
        # This should route /excellence/api/health to /excellence/backend/api/health
        # PowerShell version compatibility
        if ($PSVersionTable.PSVersion.Major -ge 6) {
            $response = Invoke-WebRequest -Uri "https://$ServerName`:$Port/excellence/api/health" -SkipCertificateCheck -UseBasicParsing -TimeoutSec 10
        } else {
            # Windows PowerShell 5.1 compatibility
            [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
            $response = Invoke-WebRequest -Uri "https://$ServerName`:$Port/excellence/api/health" -UseBasicParsing -TimeoutSec 10
            [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
        }
        
        if ($response.StatusCode -eq 200) {
            $content = $response.Content | ConvertFrom-Json -ErrorAction SilentlyContinue
            if ($content -and $content.status -eq "healthy") {
                return "Backend API accessible via IIS routing (status: $($content.status))"
            } else {
                return "Backend API accessible but response format unexpected"
            }
        } else {
            throw "Backend API returned status $($response.StatusCode)"
        }
    } catch {
        throw "Backend API test failed: $($_.Exception.Message)"
    }
}

# Test 11: Check Windows Firewall
Test-Step "Checking Windows Firewall" {
    $firewallRule = Get-NetFirewallRule | Where-Object { $_.DisplayName -like "*Excel*9443*" -or $_.DisplayName -like "*IIS*9443*" } | Select-Object -First 1
    if ($firewallRule) {
        return "Firewall rule found: $($firewallRule.DisplayName) (Enabled: $($firewallRule.Enabled))"
    } else {
        throw "No firewall rule found for port $Port"
    }
}

# Summary
Write-Host "`n" + "="*50 -ForegroundColor Yellow
Write-Host "Test Results: $testsPassed/$totalTests tests passed" -ForegroundColor $(if ($testsPassed -eq $totalTests) { "Green" } else { "Red" })

if ($testsPassed -eq $totalTests) {
    Write-Host "`n[OK] IIS setup appears to be working correctly!" -ForegroundColor Green -BackgroundColor DarkGreen
    Write-Host "URLs to test:" -ForegroundColor Cyan
    Write-Host "  Main site: https://$ServerName`:$Port/excellence/" -ForegroundColor White
    Write-Host "  Health check: https://$ServerName`:$Port/health" -ForegroundColor White
    Write-Host "  Taskpane: https://$ServerName`:$Port/excellence/taskpane.html" -ForegroundColor White
    Write-Host "  Backend API: https://$ServerName`:$Port/excellence/api/health" -ForegroundColor White
} else {
    Write-Host "`n✗ Some tests failed. Please check the issues above." -ForegroundColor Red
    Write-Host "Common solutions:" -ForegroundColor Yellow
    Write-Host "  1. Ensure IIS is installed with URL Rewrite and ARR modules" -ForegroundColor White
    Write-Host "  2. Run .\setup-iis.ps1 -Force to reconfigure" -ForegroundColor White
    Write-Host "  3. Check that frontend files are built and deployed" -ForegroundColor White
    Write-Host "  4. Verify SSL certificate is properly configured" -ForegroundColor White
    Write-Host "  5. Run backend setup manually: .\deployment\scripts\setup-backend-iis.ps1" -ForegroundColor White
}

Write-Host "="*50 -ForegroundColor Yellow