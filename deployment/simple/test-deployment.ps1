param(
    [string]$ApplicationName = "excellence",
    [int]$TimeoutSeconds = 30
)

Write-Host "=== Testing Simple Deployment ===" -ForegroundColor Cyan

$BaseUrl = "https://localhost:9443/$ApplicationName"
$FrontendUrl = "$BaseUrl/"
$BackendHealthUrl = "$BaseUrl/backend/api/health"

# Function to test URL with timeout
function Test-Url {
    param(
        [string]$Url,
        [int]$TimeoutSeconds = 30
    )
    
    try {
        Write-Host "Testing: $Url" -ForegroundColor Yellow
        
        # Skip SSL certificate validation for localhost testing
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
        
        $request = [System.Net.WebRequest]::Create($Url)
        $request.Timeout = $TimeoutSeconds * 1000
        $request.Method = "GET"
        
        $response = $request.GetResponse()
        $statusCode = $response.StatusCode
        $response.Close()
        
        Write-Host "✓ $Url - Status: $statusCode" -ForegroundColor Green
        return $true
        
    } catch [System.Net.WebException] {
        $statusCode = $_.Exception.Response.StatusCode
        Write-Host "✗ $Url - Error: $statusCode" -ForegroundColor Red
        Write-Host "  Details: $($_.Exception.Message)" -ForegroundColor Red
        return $false
        
    } catch {
        Write-Host "✗ $Url - Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to test API endpoint with JSON response
function Test-ApiEndpoint {
    param(
        [string]$Url,
        [int]$TimeoutSeconds = 30
    )
    
    try {
        Write-Host "Testing API: $Url" -ForegroundColor Yellow
        
        # Skip SSL certificate validation
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
        
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("Accept", "application/json")
        
        $jsonResponse = $webClient.DownloadString($Url)
        $response = ConvertFrom-Json $jsonResponse
        
        Write-Host "✓ $Url - API Response:" -ForegroundColor Green
        Write-Host "  Status: $($response.status)" -ForegroundColor Green
        Write-Host "  Message: $($response.message)" -ForegroundColor Green
        return $true
        
    } catch [System.Net.WebException] {
        $statusCode = $_.Exception.Response.StatusCode
        Write-Host "✗ $Url - HTTP Error: $statusCode" -ForegroundColor Red
        Write-Host "  Details: $($_.Exception.Message)" -ForegroundColor Red
        return $false
        
    } catch {
        Write-Host "✗ $Url - Error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    } finally {
        if ($webClient) { $webClient.Dispose() }
    }
}

$TestResults = @()

# Test frontend
Write-Host "`n1. Testing Frontend..." -ForegroundColor Yellow
$FrontendTest = Test-Url -Url $FrontendUrl -TimeoutSeconds $TimeoutSeconds
$TestResults += @{ Name = "Frontend"; Url = $FrontendUrl; Success = $FrontendTest }

# Test backend health endpoint
Write-Host "`n2. Testing Backend API..." -ForegroundColor Yellow  
$BackendTest = Test-ApiEndpoint -Url $BackendHealthUrl -TimeoutSeconds $TimeoutSeconds
$TestResults += @{ Name = "Backend API"; Url = $BackendHealthUrl; Success = $BackendTest }

# Summary
Write-Host "`n=== Test Results Summary ===" -ForegroundColor Cyan

$AllPassed = $true
foreach ($result in $TestResults) {
    $status = if ($result.Success) { "✓ PASS" } else { "✗ FAIL"; $AllPassed = $false }
    $color = if ($result.Success) { "Green" } else { "Red" }
    Write-Host "$status - $($result.Name): $($result.Url)" -ForegroundColor $color
}

if ($AllPassed) {
    Write-Host "`n🎉 All tests passed! Deployment is working correctly." -ForegroundColor Green
    Write-Host "You can now use the Excel add-in with:" -ForegroundColor Cyan
    Write-Host "  Frontend: $FrontendUrl" -ForegroundColor Cyan
    Write-Host "  Backend API: $BaseUrl/backend/api/" -ForegroundColor Cyan
    exit 0
} else {
    Write-Host "`n❌ Some tests failed. Check IIS configuration and logs." -ForegroundColor Red
    Write-Host "Troubleshooting tips:" -ForegroundColor Yellow
    Write-Host "  1. Check IIS Manager for application status" -ForegroundColor Yellow
    Write-Host "  2. Review Windows Event Logs (Application)" -ForegroundColor Yellow  
    Write-Host "  3. Verify Python and wfastcgi installation" -ForegroundColor Yellow
    Write-Host "  4. Check file permissions on IIS directories" -ForegroundColor Yellow
    exit 1
}