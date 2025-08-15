# Comprehensive Connectivity Debug Script
# Tests all communication paths: Frontend, Backend, NiFi, and IIS Proxy

param(
    [string]$Environment = "staging",
    [string]$FrontendPort = "3000",
    [string]$BackendPort = "5000",
    [string]$ProxyPort = "9443"
)

$ErrorActionPreference = "Continue"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  EXCELADDIN CONNECTIVITY DEBUG" -ForegroundColor Cyan
Write-Host "  Environment: $Environment" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Test results collection
$results = @()

# Helper function to test URL
function Test-Url {
    param(
        [string]$Name,
        [string]$Url,
        [int]$TimeoutSec = 10
    )
    
    Write-Host "Testing $Name..." -ForegroundColor Yellow -NoNewline
    
    try {
        $response = Invoke-WebRequest -Uri $Url -TimeoutSec $TimeoutSec -UseBasicParsing -ErrorAction Stop
        Write-Host " ✅ SUCCESS (HTTP $($response.StatusCode))" -ForegroundColor Green
        return @{
            Name = $Name
            Url = $Url
            Status = "SUCCESS"
            StatusCode = $response.StatusCode
            Error = $null
        }
    } catch {
        Write-Host " ❌ FAILED" -ForegroundColor Red
        Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red
        return @{
            Name = $Name
            Url = $Url
            Status = "FAILED"
            StatusCode = $null
            Error = $_.Exception.Message
        }
    }
}

# 1. TEST FRONTEND SERVICE
Write-Host "1. TESTING FRONTEND SERVICE" -ForegroundColor Yellow
Write-Host "==============================" -ForegroundColor Yellow

$frontendUrls = @(
    @{ Name = "Frontend Health Check"; Url = "http://localhost:$FrontendPort/health" },
    @{ Name = "Frontend Debug Assets"; Url = "http://localhost:$FrontendPort/debug/assets" },
    @{ Name = "Frontend Taskpane"; Url = "http://localhost:$FrontendPort/excellence/taskpane.html" },
    @{ Name = "Frontend Commands"; Url = "http://localhost:$FrontendPort/excellence/commands.html" },
    @{ Name = "Frontend Assets"; Url = "http://localhost:$FrontendPort/assets/PCAG_white_trans.png" },
    @{ Name = "Frontend Excellence Assets"; Url = "http://localhost:$FrontendPort/excellence/assets/PCAG_white_trans.png" }
)

foreach ($test in $frontendUrls) {
    $results += Test-Url -Name $test.Name -Url $test.Url
}

Write-Host ""

# 2. TEST BACKEND SERVICE
Write-Host "2. TESTING BACKEND SERVICE" -ForegroundColor Yellow
Write-Host "=============================" -ForegroundColor Yellow

$backendUrls = @(
    @{ Name = "Backend Health Check"; Url = "http://localhost:$BackendPort/api/health" },
    @{ Name = "Backend Debug Info"; Url = "http://localhost:$BackendPort/api/debug" },
    @{ Name = "Backend Root"; Url = "http://localhost:$BackendPort/" },
    @{ Name = "Raw Data Categories"; Url = "http://localhost:$BackendPort/api/raw-data/categories" }
)

foreach ($test in $backendUrls) {
    $results += Test-Url -Name $test.Name -Url $test.Url
}

Write-Host ""

# 3. TEST IIS PROXY (if running)
Write-Host "3. TESTING IIS PROXY SERVICE" -ForegroundColor Yellow
Write-Host "==============================" -ForegroundColor Yellow

$proxyUrls = @(
    @{ Name = "Proxy Status Page"; Url = "http://localhost:$ProxyPort/" },
    @{ Name = "Proxy Health (Frontend)"; Url = "http://localhost:$ProxyPort/excellence/taskpane.html" },
    @{ Name = "Proxy API Health (Backend)"; Url = "http://localhost:$ProxyPort/api/health" },
    @{ Name = "Proxy Assets"; Url = "http://localhost:$ProxyPort/excellence/assets/PCAG_white_trans.png" }
)

foreach ($test in $proxyUrls) {
    $results += Test-Url -Name $test.Name -Url $test.Url
}

Write-Host ""

# 4. TEST NIFI CONNECTIVITY
Write-Host "4. TESTING NIFI CONNECTIVITY" -ForegroundColor Yellow
Write-Host "==============================" -ForegroundColor Yellow

# Read NiFi endpoint from backend env file
$envFile = ".env.$Environment"
$backendEnvPath = "../backend/$envFile"

if (Test-Path $backendEnvPath) {
    $envContent = Get-Content $backendEnvPath
    $nifiEndpoint = ($envContent | Where-Object { $_ -match "NIFI_ENDPOINT=" }) -replace "NIFI_ENDPOINT=", ""
    
    if ($nifiEndpoint) {
        Write-Host "Found NiFi endpoint: $nifiEndpoint" -ForegroundColor Gray
        $results += Test-Url -Name "NiFi Endpoint" -Url $nifiEndpoint -TimeoutSec 15
    } else {
        Write-Host "⚠️  NiFi endpoint not found in $backendEnvPath" -ForegroundColor Yellow
    }
} else {
    Write-Host "⚠️  Backend environment file not found: $backendEnvPath" -ForegroundColor Yellow
}

Write-Host ""

# 5. CHECK WINDOWS SERVICES
Write-Host "5. CHECKING WINDOWS SERVICES" -ForegroundColor Yellow
Write-Host "==============================" -ForegroundColor Yellow

$services = @("ExcelAddin-Frontend", "ExcelAddin-Backend")
foreach ($serviceName in $services) {
    try {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if ($service) {
            $status = if ($service.Status -eq "Running") { "✅ RUNNING" } else { "❌ $($service.Status)" }
            Write-Host "Service '$serviceName': $status" -ForegroundColor $(if ($service.Status -eq "Running") { "Green" } else { "Red" })
        } else {
            Write-Host "Service '$serviceName': ❌ NOT FOUND" -ForegroundColor Red
        }
    } catch {
        Write-Host "Service '$serviceName': ❌ ERROR - $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ""

# 6. CHECK PORT USAGE
Write-Host "6. CHECKING PORT USAGE" -ForegroundColor Yellow
Write-Host "========================" -ForegroundColor Yellow

$ports = @($FrontendPort, $BackendPort, $ProxyPort)
foreach ($port in $ports) {
    try {
        $connection = Get-NetTCPConnection -LocalPort $port -ErrorAction SilentlyContinue
        if ($connection) {
            Write-Host "Port $port`: ✅ IN USE by PID $($connection[0].OwningProcess)" -ForegroundColor Green
            
            # Try to get process name
            try {
                $process = Get-Process -Id $connection[0].OwningProcess -ErrorAction SilentlyContinue
                if ($process) {
                    Write-Host "  Process: $($process.ProcessName) ($($process.Id))" -ForegroundColor Gray
                }
            } catch {
                # Ignore process lookup errors
            }
        } else {
            Write-Host "Port $port`: ❌ NOT IN USE" -ForegroundColor Red
        }
    } catch {
        Write-Host "Port $port`: ❌ ERROR - $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ""

# 7. SUMMARY
Write-Host "7. SUMMARY AND RECOMMENDATIONS" -ForegroundColor Yellow
Write-Host "================================" -ForegroundColor Yellow

$successCount = ($results | Where-Object { $_.Status -eq "SUCCESS" }).Count
$failCount = ($results | Where-Object { $_.Status -eq "FAILED" }).Count

Write-Host "Total Tests: $($results.Count)" -ForegroundColor White
Write-Host "Successful: $successCount" -ForegroundColor Green
Write-Host "Failed: $failCount" -ForegroundColor Red
Write-Host ""

if ($failCount -gt 0) {
    Write-Host "FAILED TESTS:" -ForegroundColor Red
    $failedTests = $results | Where-Object { $_.Status -eq "FAILED" }
    foreach ($test in $failedTests) {
        Write-Host "  ❌ $($test.Name): $($test.Url)" -ForegroundColor Red
        Write-Host "     Error: $($test.Error)" -ForegroundColor Gray
    }
    Write-Host ""
    
    Write-Host "RECOMMENDATIONS:" -ForegroundColor Yellow
    
    # Check if frontend is down
    $frontendDown = $results | Where-Object { $_.Name -like "*Frontend*" -and $_.Status -eq "FAILED" }
    if ($frontendDown) {
        Write-Host "  🔧 Frontend service appears to be down:" -ForegroundColor Yellow
        Write-Host "     - Check if the frontend service is running: Get-Service -Name 'ExcelAddin-Frontend'" -ForegroundColor Gray
        Write-Host "     - Try restarting: Restart-Service -Name 'ExcelAddin-Frontend'" -ForegroundColor Gray
        Write-Host "     - Check if the dist folder exists and has been built: npm run build:staging" -ForegroundColor Gray
    }
    
    # Check if backend is down
    $backendDown = $results | Where-Object { $_.Name -like "*Backend*" -and $_.Status -eq "FAILED" }
    if ($backendDown) {
        Write-Host "  🔧 Backend service appears to be down:" -ForegroundColor Yellow
        Write-Host "     - Check if the backend service is running: Get-Service -Name 'ExcelAddin-Backend'" -ForegroundColor Gray
        Write-Host "     - Try restarting: Restart-Service -Name 'ExcelAddin-Backend'" -ForegroundColor Gray
        Write-Host "     - Check backend logs for errors" -ForegroundColor Gray
    }
    
    # Check if assets are missing
    $assetsFailed = $results | Where-Object { $_.Name -like "*Assets*" -and $_.Status -eq "FAILED" }
    if ($assetsFailed) {
        Write-Host "  🔧 Assets are not being served properly:" -ForegroundColor Yellow
        Write-Host "     - Ensure the project has been built: npm run build:staging" -ForegroundColor Gray
        Write-Host "     - Check if dist/assets/ directory exists and contains PNG files" -ForegroundColor Gray
        Write-Host "     - Verify Vite configuration is correct for asset paths" -ForegroundColor Gray
    }
    
    # Check if NiFi is down
    $nifiDown = $results | Where-Object { $_.Name -like "*NiFi*" -and $_.Status -eq "FAILED" }
    if ($nifiDown) {
        Write-Host "  🔧 NiFi connectivity issues:" -ForegroundColor Yellow
        Write-Host "     - Verify NiFi is running on the specified port" -ForegroundColor Gray
        Write-Host "     - Check if the NiFi endpoint URL is correct in the .env file" -ForegroundColor Gray
        Write-Host "     - Verify SSL certificates if using HTTPS" -ForegroundColor Gray
        Write-Host "     - Check network connectivity and firewall settings" -ForegroundColor Gray
    }
    
} else {
    Write-Host "🎉 ALL TESTS PASSED! The infrastructure appears to be working correctly." -ForegroundColor Green
}

Write-Host ""
Write-Host "Debug completed at $(Get-Date)" -ForegroundColor Gray