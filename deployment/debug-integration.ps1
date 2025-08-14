# ExcelAddin Integration Debug Script
# Comprehensive connectivity and deployment troubleshooting for Windows Server 10

param(
    [switch]$Detailed,
    [switch]$FixIssues,
    [string]$TestUrl = "https://server-vs81t.intranet.local:9443"
)

$ErrorActionPreference = "Continue"

# Configuration
$BackendServiceName = "ExcelAddin-Backend"  
$FrontendServiceName = "ExcelAddin-Frontend"
$BackendPort = 5000
$FrontendPort = 3000
$IISPort = 9443
$LogDir = "C:\Logs\ExcelAddin"

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  ExcelAddin Integration Debug & Test" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# Helper functions
function Test-Port {
    param($Computer, $Port, $Timeout = 5)
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $connect = $tcp.BeginConnect($Computer, $Port, $null, $null)
        $wait = $connect.AsyncWaitHandle.WaitOne($Timeout * 1000, $false)
        if ($wait) {
            $tcp.EndConnect($connect)
            $tcp.Close()
            return $true
        }
        $tcp.Close()
        return $false
    } catch {
        return $false
    }
}

function Get-ServiceDetails {
    param($ServiceName)
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($service) {
        return @{
            Name = $service.Name
            Status = $service.Status
            StartType = $service.StartType
            DisplayName = $service.DisplayName
        }
    }
    return $null
}

function Test-HttpEndpoint {
    param($Url, $ExpectedContent = $null, $Timeout = 10)
    try {
        $response = Invoke-WebRequest -Uri $Url -TimeoutSec $Timeout -ErrorAction SilentlyContinue
        $result = @{
            Success = $true
            StatusCode = $response.StatusCode
            ContentLength = $response.Content.Length
            Headers = $response.Headers
            Content = if ($response.Content.Length -lt 500) { $response.Content } else { $response.Content.Substring(0, 500) + "..." }
        }
        
        if ($ExpectedContent -and $response.Content -notlike "*$ExpectedContent*") {
            $result.Success = $false
            $result.Error = "Expected content not found: $ExpectedContent"
        }
        
        return $result
    } catch {
        return @{
            Success = $false
            Error = $_.Exception.Message
            StatusCode = $null
        }
    }
}

# 1. System Information
Write-Host "1. SYSTEM INFORMATION" -ForegroundColor Yellow
Write-Host "=" * 50

$osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
Write-Host "  OS: $($osInfo.Caption) $($osInfo.Version)" -ForegroundColor Green
Write-Host "  PowerShell: $($PSVersionTable.PSVersion)" -ForegroundColor Green

$hostname = $env:COMPUTERNAME
Write-Host "  Hostname: $hostname" -ForegroundColor Green
Write-Host ""

# 2. Service Status Check
Write-Host "2. SERVICE STATUS" -ForegroundColor Yellow  
Write-Host "=" * 50

$backendService = Get-ServiceDetails $BackendServiceName
if ($backendService) {
    Write-Host "  Backend Service:" -ForegroundColor Green
    Write-Host "    Name: $($backendService.Name)"
    Write-Host "    Status: $($backendService.Status)"
    Write-Host "    Start Type: $($backendService.StartType)"
    
    if ($backendService.Status -ne "Running") {
        Write-Warning "    ⚠️  Backend service is not running!"
        if ($FixIssues) {
            Write-Host "    Attempting to start backend service..." -ForegroundColor Yellow
            Start-Service -Name $BackendServiceName -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 3
            $backendService = Get-ServiceDetails $BackendServiceName
            Write-Host "    New Status: $($backendService.Status)"
        }
    }
} else {
    Write-Error "  Backend service '$BackendServiceName' not found!"
}

$frontendService = Get-ServiceDetails $FrontendServiceName
if ($frontendService) {
    Write-Host "  Frontend Service:" -ForegroundColor Green
    Write-Host "    Name: $($frontendService.Name)"
    Write-Host "    Status: $($frontendService.Status)"
    Write-Host "    Start Type: $($frontendService.StartType)"
    
    if ($frontendService.Status -ne "Running") {
        Write-Warning "    ⚠️  Frontend service is not running!"
        if ($FixIssues) {
            Write-Host "    Attempting to start frontend service..." -ForegroundColor Yellow
            Start-Service -Name $FrontendServiceName -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 3
            $frontendService = Get-ServiceDetails $FrontendServiceName
            Write-Host "    New Status: $($frontendService.Status)"
        }
    }
} else {
    Write-Error "  Frontend service '$FrontendServiceName' not found!"
}

Write-Host ""

# 3. Port Connectivity
Write-Host "3. PORT CONNECTIVITY" -ForegroundColor Yellow
Write-Host "=" * 50

$ports = @(
    @{ Port = $BackendPort; Service = "Backend" },
    @{ Port = $FrontendPort; Service = "Frontend" },
    @{ Port = $IISPort; Service = "IIS Proxy" }
)

foreach ($portTest in $ports) {
    $isOpen = Test-Port "localhost" $portTest.Port
    $status = if ($isOpen) { "OPEN" } else { "CLOSED" }
    $color = if ($isOpen) { "Green" } else { "Red" }
    Write-Host "  Port $($portTest.Port) ($($portTest.Service)): $status" -ForegroundColor $color
    
    if ($isOpen) {
        # Get process information
        $connection = Get-NetTCPConnection -LocalPort $portTest.Port -ErrorAction SilentlyContinue | 
            Where-Object { $_.State -eq "Listen" }
        if ($connection) {
            $processId = $connection.OwningProcess
            $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
            if ($process) {
                Write-Host "    Process: $($process.ProcessName) (PID: $processId)" -ForegroundColor Gray
            }
        }
    }
}

Write-Host ""

# 4. Backend API Tests (Excel Add-in specific endpoints)
Write-Host "4. BACKEND API TESTS" -ForegroundColor Yellow
Write-Host "=" * 50

if (Test-Port "localhost" $BackendPort) {
    $backendEndpoints = @(
        @{ Path = "/api/health"; Name = "Health Check"; Expected = "healthy" },
        @{ Path = "/api"; Name = "API Root"; Expected = $null }
    )
    
    foreach ($endpoint in $backendEndpoints) {
        $url = "http://localhost:$BackendPort$($endpoint.Path)"
        $result = Test-HttpEndpoint $url $endpoint.Expected
        
        $status = if ($result.Success) { "PASS" } else { "FAIL" }
        $color = if ($result.Success) { "Green" } else { "Red" }
        Write-Host "  $($endpoint.Name): $status" -ForegroundColor $color
        
        if ($Detailed) {
            Write-Host "    URL: $url" -ForegroundColor Gray
            if ($result.Success) {
                Write-Host "    Status Code: $($result.StatusCode)" -ForegroundColor Gray
                Write-Host "    Content Length: $($result.ContentLength)" -ForegroundColor Gray
            } else {
                Write-Host "    Error: $($result.Error)" -ForegroundColor Gray
            }
        }
    }
} else {
    Write-Warning "  Backend port $BackendPort is not accessible - skipping API tests"
}

Write-Host ""

# 5. Frontend Tests (Excel Add-in specific files)
Write-Host "5. FRONTEND TESTS" -ForegroundColor Yellow
Write-Host "=" * 50

if (Test-Port "localhost" $FrontendPort) {
    # Test the exact endpoints Excel will access
    $frontendEndpoints = @(
        @{ Path = "/health"; Name = "Health Check"; Expected = "healthy" },
        @{ Path = "/excellence/taskpane.html"; Name = "Taskpane HTML"; Expected = "PrimeExcelence" },
        @{ Path = "/excellence/commands.html"; Name = "Commands HTML"; Expected = "Commands" },
        @{ Path = "/functions.json"; Name = "Functions Manifest"; Expected = "functions" },
        @{ Path = "/manifest.xml"; Name = "Excel Manifest"; Expected = "OfficeApp" }
    )
    
    foreach ($endpoint in $frontendEndpoints) {
        $url = "http://localhost:$FrontendPort$($endpoint.Path)"
        $result = Test-HttpEndpoint $url $endpoint.Expected
        
        $status = if ($result.Success) { "PASS" } else { "FAIL" }
        $color = if ($result.Success) { "Green" } else { "Red" }
        Write-Host "  $($endpoint.Name): $status" -ForegroundColor $color
        
        if ($result.Success -and $endpoint.Path -like "*.html") {
            # Check if HTML contains required Excel Add-in elements
            if ($result.Content -like "*office.js*") {
                Write-Host "    ✓ Office.js reference found" -ForegroundColor Green
            } else {
                Write-Warning "    ⚠️  Office.js reference missing!"
            }
        }
        
        if ($Detailed) {
            Write-Host "    URL: $url" -ForegroundColor Gray
            if ($result.Success) {
                Write-Host "    Status Code: $($result.StatusCode)" -ForegroundColor Gray
                Write-Host "    Content Length: $($result.ContentLength)" -ForegroundColor Gray
                if ($result.Content.Length -lt 200) {
                    Write-Host "    Content Preview: $($result.Content)" -ForegroundColor Gray
                }
            } else {
                Write-Host "    Error: $($result.Error)" -ForegroundColor Gray
            }
        }
    }
} else {
    Write-Warning "  Frontend port $FrontendPort is not accessible - skipping frontend tests"
}

Write-Host ""

# 6. IIS Integration Test
Write-Host "6. IIS INTEGRATION TEST" -ForegroundColor Yellow
Write-Host "=" * 50

if (Test-Port "localhost" $IISPort) {
    Write-Host "  IIS proxy port $IISPort is accessible" -ForegroundColor Green
    
    # Test through IIS proxy (the actual URLs Excel will use)
    $excelEndpoints = @(
        @{ Path = "/excellence/taskpane.html"; Name = "Excel Taskpane" },
        @{ Path = "/excellence/commands.html"; Name = "Excel Commands" },
        @{ Path = "/api/health"; Name = "API via Proxy" }
    )
    
    foreach ($endpoint in $excelEndpoints) {
        $url = "http://localhost:$IISPort$($endpoint.Path)"
        $result = Test-HttpEndpoint $url
        
        $status = if ($result.Success) { "PASS" } else { "FAIL" }
        $color = if ($result.Success) { "Green" } else { "Red" }
        Write-Host "  $($endpoint.Name): $status" -ForegroundColor $color
        
        if ($Detailed -and -not $result.Success) {
            Write-Host "    Error: $($result.Error)" -ForegroundColor Gray
        }
    }
} else {
    Write-Warning "  IIS proxy port $IISPort is not accessible"
    Write-Host "  Check if IIS is configured and running" -ForegroundColor Yellow
}

Write-Host ""

# 7. External URL Test (if provided)
if ($TestUrl -ne "https://server-vs81t.intranet.local:9443") {
    Write-Host "7. EXTERNAL URL TEST" -ForegroundColor Yellow
    Write-Host "=" * 50
    
    Write-Host "  Testing external URL: $TestUrl" -ForegroundColor Cyan
    
    $externalEndpoints = @(
        @{ Path = "/excellence/taskpane.html"; Name = "External Taskpane" },
        @{ Path = "/api/health"; Name = "External API" }
    )
    
    foreach ($endpoint in $externalEndpoints) {
        $url = "$TestUrl$($endpoint.Path)"
        $result = Test-HttpEndpoint $url
        
        $status = if ($result.Success) { "PASS" } else { "FAIL" }
        $color = if ($result.Success) { "Green" } else { "Red" }
        Write-Host "  $($endpoint.Name): $status" -ForegroundColor $color
        
        if ($Detailed) {
            Write-Host "    URL: $url" -ForegroundColor Gray
            if (-not $result.Success) {
                Write-Host "    Error: $($result.Error)" -ForegroundColor Gray
            }
        }
    }
    
    Write-Host ""
}

# 8. Log File Analysis
Write-Host "8. LOG FILE ANALYSIS" -ForegroundColor Yellow
Write-Host "=" * 50

if (Test-Path $LogDir) {
    Write-Host "  Log directory: $LogDir" -ForegroundColor Green
    
    $logFiles = @(
        "backend-stderr.log",
        "backend-stdout.log", 
        "frontend-stderr.log",
        "frontend-stdout.log"
    )
    
    foreach ($logFile in $logFiles) {
        $logPath = Join-Path $LogDir $logFile
        if (Test-Path $logPath) {
            $logInfo = Get-Item $logPath
            Write-Host "  $logFile: $($logInfo.Length) bytes, modified $($logInfo.LastWriteTime)" -ForegroundColor Green
            
            if ($Detailed) {
                $recentLogs = Get-Content $logPath -Tail 3 -ErrorAction SilentlyContinue
                if ($recentLogs) {
                    Write-Host "    Recent entries:" -ForegroundColor Gray
                    foreach ($line in $recentLogs) {
                        Write-Host "      $line" -ForegroundColor Gray
                    }
                }
            }
        } else {
            Write-Warning "  $logFile: Not found"
        }
    }
} else {
    Write-Warning "  Log directory not found: $LogDir"
}

Write-Host ""

# 9. Summary and Recommendations
Write-Host "9. SUMMARY & RECOMMENDATIONS" -ForegroundColor Yellow
Write-Host "=" * 50

$issues = @()
$recommendations = @()

# Check service status
if ($backendService -and $backendService.Status -ne "Running") {
    $issues += "Backend service is not running"
    $recommendations += "Start backend service: Start-Service -Name '$BackendServiceName'"
}

if ($frontendService -and $frontendService.Status -ne "Running") {
    $issues += "Frontend service is not running"
    $recommendations += "Start frontend service: Start-Service -Name '$FrontendServiceName'"
}

# Check port accessibility
if (-not (Test-Port "localhost" $BackendPort)) {
    $issues += "Backend port $BackendPort is not accessible"
    $recommendations += "Check backend service and firewall rules"
}

if (-not (Test-Port "localhost" $FrontendPort)) {
    $issues += "Frontend port $FrontendPort is not accessible"
    $recommendations += "Check frontend service and firewall rules"
}

if (-not (Test-Port "localhost" $IISPort)) {
    $issues += "IIS proxy port $IISPort is not accessible"
    $recommendations += "Configure IIS site and check bindings"
}

if ($issues.Count -eq 0) {
    Write-Host "  ✅ No critical issues detected!" -ForegroundColor Green
    Write-Host "  System appears to be functioning correctly" -ForegroundColor Green
} else {
    Write-Host "  ⚠️  Issues detected:" -ForegroundColor Red
    foreach ($issue in $issues) {
        Write-Host "    • $issue" -ForegroundColor Red
    }
    
    Write-Host ""
    Write-Host "  Recommendations:" -ForegroundColor Yellow
    foreach ($rec in $recommendations) {
        Write-Host "    • $rec" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  Integration Debug Complete" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

if ($issues.Count -gt 0) {
    Write-Host "To fix issues automatically, run with -FixIssues parameter" -ForegroundColor Yellow
    Write-Host "For detailed output, run with -Detailed parameter" -ForegroundColor Yellow
}

Write-Host "Useful Commands:" -ForegroundColor Cyan
Write-Host "  Check services: Get-Service '$BackendServiceName', '$FrontendServiceName'" 
Write-Host "  View logs: Get-Content '$LogDir\\*-stderr.log' -Tail 20"
Write-Host "  Test endpoints: Invoke-WebRequest 'http://localhost:3000/health'"