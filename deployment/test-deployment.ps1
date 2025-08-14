# ExcelAddin Deployment Testing Script
# Comprehensive tests for all deployed services

param(
    [switch]$Verbose,
    [switch]$SkipExternal
)

# Import common functions
. (Join-Path $PSScriptRoot "scripts" | Join-Path -ChildPath "common.ps1")

$TestResults = @()

function Add-TestResult {
    param(
        [string]$TestName,
        [bool]$Passed,
        [string]$Message = "",
        [string]$Details = ""
    )
    
    $TestResults += [PSCustomObject]@{
        Test = $TestName
        Passed = $Passed
        Message = $Message
        Details = $Details
    }
    
    if ($Passed) {
        Write-Success "PASS: $TestName - $Message"
    } else {
        Write-Error "FAIL: $TestName - $Message"
    }
    
    if ($Verbose -and $Details) {
        Write-Host "  Details: $Details" -ForegroundColor Gray
    }
}

Write-Header "ExcelAddin Deployment Testing"

try {
    # Test 1: Service Status Tests
    Write-Header "Service Status Tests"
    
    # Test Backend NSSM Service
    Write-Host "Testing backend service status..."
    $backendService = Get-Service -Name "ExcelAddin-Backend" -ErrorAction SilentlyContinue
    if ($backendService) {
        Add-TestResult -TestName "Backend Service Exists" -Passed $true -Message "Service found"
        Add-TestResult -TestName "Backend Service Running" -Passed ($backendService.Status -eq "Running") -Message "Status: $($backendService.Status)"
    } else {
        Add-TestResult -TestName "Backend Service Exists" -Passed $false -Message "Service not found"
        Add-TestResult -TestName "Backend Service Running" -Passed $false -Message "Service not found"
    }
    
    # Test Frontend NSSM Service
    Write-Host "Testing frontend service status..."
    $frontendService = Get-Service -Name "ExcelAddin-Frontend" -ErrorAction SilentlyContinue
    if ($frontendService) {
        Add-TestResult -TestName "Frontend Service Exists" -Passed $true -Message "Service found"
        Add-TestResult -TestName "Frontend Service Running" -Passed ($frontendService.Status -eq "Running") -Message "Status: $($frontendService.Status)"
    } else {
        Add-TestResult -TestName "Frontend Service Exists" -Passed $false -Message "Service not found"
        Add-TestResult -TestName "Frontend Service Running" -Passed $false -Message "Service not found"
    }
    
    # Test IIS Site
    Write-Host "Testing IIS site status..."
    $iisSite = Get-IISSite -Name "ExcelAddin" -ErrorAction SilentlyContinue
    if ($iisSite) {
        Add-TestResult -TestName "IIS Site Exists" -Passed $true -Message "Site found"
        Add-TestResult -TestName "IIS Site Running" -Passed ($iisSite.State -eq "Started") -Message "State: $($iisSite.State)"
    } else {
        Add-TestResult -TestName "IIS Site Exists" -Passed $false -Message "Site not found"
        Add-TestResult -TestName "IIS Site Running" -Passed $false -Message "Site not found"
    }
    
    # Test 2: Port Availability Tests
    Write-Header "Port Connectivity Tests"
    
    # Test Backend Port
    Write-Host "Testing backend port connectivity..."
    $backendPortOpen = Test-Port -Port 5000
    Add-TestResult -TestName "Backend Port 5000" -Passed $backendPortOpen -Message $(if ($backendPortOpen) { "Port accessible" } else { "Port not accessible" })
    
    # Test Frontend Port
    Write-Host "Testing frontend port connectivity..."
    $frontendPortOpen = Test-Port -Port 3000
    Add-TestResult -TestName "Frontend Port 3000" -Passed $frontendPortOpen -Message $(if ($frontendPortOpen) { "Port accessible" } else { "Port not accessible" })
    
    # Test IIS Port
    Write-Host "Testing IIS port connectivity..."
    $iisPortOpen = Test-Port -Port 9443
    Add-TestResult -TestName "IIS Port 9443" -Passed $iisPortOpen -Message $(if ($iisPortOpen) { "Port accessible" } else { "Port not accessible" })
    
    # Test 3: HTTP Health Check Tests
    Write-Header "HTTP Health Check Tests"
    
    # Test Backend Health Endpoint
    Write-Host "Testing backend health endpoint..."
    try {
        $backendHealth = Invoke-RestMethod -Uri "http://127.0.0.1:5000/api/health" -TimeoutSec 10
        $healthPassed = $backendHealth.status -eq "healthy"
        Add-TestResult -TestName "Backend Health Check" -Passed $healthPassed -Message "Response: $($backendHealth.status)" -Details $backendHealth.message
    } catch {
        Add-TestResult -TestName "Backend Health Check" -Passed $false -Message "Health check failed" -Details $_.Exception.Message
    }
    
    # Test Frontend HTTP Response
    Write-Host "Testing frontend HTTP response..."
    try {
        # First check if port is listening
        $portListening = Test-NetConnection -ComputerName "127.0.0.1" -Port 3000 -InformationLevel Quiet -ErrorAction SilentlyContinue
        if (-not $portListening) {
            Add-TestResult -TestName "Frontend HTTP Response" -Passed $false -Message "Port 3000 not listening" -Details "Frontend service may not be running or may have failed to bind to port"
        } else {
            $frontendResponse = Invoke-WebRequest -Uri "http://127.0.0.1:3000" -TimeoutSec 10
            $frontendPassed = $frontendResponse.StatusCode -eq 200
            Add-TestResult -TestName "Frontend HTTP Response" -Passed $frontendPassed -Message "HTTP Status: $($frontendResponse.StatusCode)"
        }
    } catch {
        Add-TestResult -TestName "Frontend HTTP Response" -Passed $false -Message "HTTP request failed" -Details $_.Exception.Message
        
        # Additional diagnostics for frontend failures
        Write-Host "Performing additional frontend diagnostics..." -ForegroundColor Yellow
        
        # Check service status
        $frontendService = Get-Service -Name "ExcelAddin-Frontend" -ErrorAction SilentlyContinue
        if ($frontendService) {
            Write-Host "  Frontend service status: $($frontendService.Status)" -ForegroundColor Gray
        } else {
            Write-Host "  Frontend service not found" -ForegroundColor Gray
        }
        
        # Check for port conflicts
        try {
            $portConflict = Get-NetTCPConnection -LocalPort 3000 -ErrorAction SilentlyContinue
            if ($portConflict) {
                $conflictProcess = Get-Process -Id $portConflict.OwningProcess -ErrorAction SilentlyContinue
                if ($conflictProcess) {
                    Write-Host "  Port 3000 in use by: $($conflictProcess.ProcessName) (PID: $($conflictProcess.Id))" -ForegroundColor Gray
                }
            } else {
                Write-Host "  Port 3000 is not in use" -ForegroundColor Gray
            }
        } catch {
            Write-Host "  Could not check port usage" -ForegroundColor Gray
        }
        
        # Check recent logs
        $logFile = "C:\Logs\ExcelAddin\frontend-stderr.log"
        if (Test-Path $logFile) {
            Write-Host "  Recent error log entries:" -ForegroundColor Gray
            Get-Content $logFile -Tail 3 | ForEach-Object { Write-Host "    $_" -ForegroundColor Gray }
        }
    }
    
    # Test 4: External Access Tests (if not skipped)
    if (-not $SkipExternal) {
        Write-Header "External Access Tests"
        
        # Test HTTPS External Access
        Write-Host "Testing external HTTPS access..."
        try {
            $externalResponse = Invoke-WebRequest -Uri "https://server-vs81t.intranet.local:9443" -TimeoutSec 20 -SkipCertificateCheck
            $externalPassed = $externalResponse.StatusCode -eq 200
            Add-TestResult -TestName "External HTTPS Access" -Passed $externalPassed -Message "HTTP Status: $($externalResponse.StatusCode)"
        } catch {
            Add-TestResult -TestName "External HTTPS Access" -Passed $false -Message "External access failed" -Details $_.Exception.Message
        }
        
        # Test API through proxy
        Write-Host "Testing API through IIS proxy..."
        try {
            $proxyApiResponse = Invoke-RestMethod -Uri "https://server-vs81t.intranet.local:9443/api/health" -TimeoutSec 20 -SkipCertificateCheck
            $proxyApiPassed = $proxyApiResponse.status -eq "healthy"
            Add-TestResult -TestName "API Through Proxy" -Passed $proxyApiPassed -Message "Response: $($proxyApiResponse.status)"
        } catch {
            Add-TestResult -TestName "API Through Proxy" -Passed $false -Message "Proxy API test failed" -Details $_.Exception.Message
        }
    } else {
        Write-Warning "External access tests skipped"
    }
    
    # Test 5: Configuration Tests
    Write-Header "Configuration Tests"
    
    # Test web.config exists
    $webConfigPath = "C:\inetpub\wwwroot\ExcelAddin\web.config"
    $webConfigExists = Test-Path $webConfigPath
    Add-TestResult -TestName "IIS web.config" -Passed $webConfigExists -Message $(if ($webConfigExists) { "File exists" } else { "File missing" })
    
    # Test backend environment configuration
    $backendPath = Get-BackendPath
    $backendEnvFile = Join-Path $backendPath ".env"
    $backendEnvExists = Test-Path $backendEnvFile
    Add-TestResult -TestName "Backend Environment" -Passed $backendEnvExists -Message $(if ($backendEnvExists) { "Configuration exists" } else { "Configuration missing" })
    
    # Test frontend build output
    $frontendDistPath = Join-Path (Get-ProjectRoot) "dist"
    $frontendBuiltExists = Test-Path $frontendDistPath
    Add-TestResult -TestName "Frontend Build Output" -Passed $frontendBuiltExists -Message $(if ($frontendBuiltExists) { "Build output exists" } else { "Build output missing" })
    
    # Test 6: Security Tests
    Write-Header "Security Tests"
    
    # Test SSL Certificate
    Write-Host "Testing SSL certificate..."
    try {
        $certificates = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { 
            $_.Subject -like "*server-vs81t.intranet.local*" -or 
            $_.DnsNameList -like "*server-vs81t.intranet.local*" 
        }
        $sslCertExists = $certificates.Count -gt 0
        Add-TestResult -TestName "SSL Certificate" -Passed $sslCertExists -Message $(if ($sslCertExists) { "Certificate found" } else { "Certificate missing" })
    } catch {
        Add-TestResult -TestName "SSL Certificate" -Passed $false -Message "Certificate check failed" -Details $_.Exception.Message
    }
    
    # Test Firewall Rule
    Write-Host "Testing firewall configuration..."
    try {
        $firewallRule = Get-NetFirewallRule -DisplayName "ExcelAddin HTTPS" -ErrorAction SilentlyContinue
        $firewallConfigured = $null -ne $firewallRule
        Add-TestResult -TestName "Firewall Configuration" -Passed $firewallConfigured -Message $(if ($firewallConfigured) { "Rule configured" } else { "Rule missing" })
    } catch {
        Add-TestResult -TestName "Firewall Configuration" -Passed $false -Message "Firewall check failed" -Details $_.Exception.Message
    }
    
    # Generate Test Report
    Write-Header "Test Results Summary"
    
    $totalTests = $TestResults.Count
    $passedTests = ($TestResults | Where-Object { $_.Passed }).Count
    $failedTests = $totalTests - $passedTests
    $passRate = if ($totalTests -gt 0) { [math]::Round(($passedTests / $totalTests) * 100, 1) } else { 0 }
    
    Write-Host ""
    Write-Host "Overall Results:" -ForegroundColor Cyan
    Write-Host "  Total Tests: $totalTests"
    Write-Host "  Passed: $passedTests" -ForegroundColor Green
    Write-Host "  Failed: $failedTests" -ForegroundColor Red
    Write-Host "  Pass Rate: $passRate%" -ForegroundColor $(if ($passRate -ge 80) { "Green" } elseif ($passRate -ge 60) { "Yellow" } else { "Red" })
    Write-Host ""
    
    # Show failed tests in detail
    $failedTestsList = $TestResults | Where-Object { -not $_.Passed }
    if ($failedTestsList.Count -gt 0) {
        Write-Host "Failed Tests:" -ForegroundColor Red
        foreach ($test in $failedTestsList) {
            Write-Host "  - $($test.Test): $($test.Message)" -ForegroundColor Red
            if ($test.Details) {
                Write-Host "    Details: $($test.Details)" -ForegroundColor Gray
            }
        }
        Write-Host ""
    }
    
    # Service Management Information
    Write-Header "Service Management Information"
    Write-Host "Backend Service Commands:"
    Write-Host "  Status: Get-Service -Name 'ExcelAddin-Backend'"
    Write-Host "  Start: Start-Service -Name 'ExcelAddin-Backend'"
    Write-Host "  Stop: Stop-Service -Name 'ExcelAddin-Backend'"
    Write-Host "  Logs: Check NSSM service logs in C:\Logs\ExcelAddin\"
    Write-Host ""
    Write-Host "Frontend Service Commands:"
    Write-Host "  Status: pm2 status exceladdin-frontend"
    Write-Host "  Start: pm2 start exceladdin-frontend"
    Write-Host "  Stop: pm2 stop exceladdin-frontend"
    Write-Host "  Logs: pm2 logs exceladdin-frontend"
    Write-Host ""
    Write-Host "IIS Management:"
    Write-Host "  Status: Get-IISSite -Name 'ExcelAddin'"
    Write-Host "  Start: Start-IISSite -Name 'ExcelAddin'"
    Write-Host "  Stop: Stop-IISSite -Name 'ExcelAddin'"
    Write-Host ""
    
    # Exit with appropriate code
    if ($failedTests -eq 0) {
        Write-Success "All tests passed! Deployment is healthy."
        exit 0
    } else {
        Write-Error "$failedTests test(s) failed. Please check the deployment."
        exit 1
    }
    
} catch {
    Write-Error "Testing failed: $($_.Exception.Message)"
    Write-Host $_.ScriptStackTrace
    exit 1
}