# ExcelAddin Comprehensive Troubleshooting Script
# Consolidates all debugging functionality for backend, frontend, and IIS proxy

param(
    [switch]$TestAll,
    [switch]$TestServices,
    [switch]$TestConnectivity,
    [switch]$TestIIS,
    [switch]$TestSSL,
    [switch]$FixCommonIssues,
    [switch]$RestartServices,
    [switch]$ClearLogs,
    [switch]$CheckPorts,
    [switch]$ResetIIS,
    [switch]$Verbose,
    [switch]$Debug,
    [string]$Environment = "staging",
    [string]$TestUrl = "https://server-vs81t.intranet.local:9443",
    [int]$Duration = 0,
    [string]$ServiceName = "",
    [array]$Ports = @(3000, 5000, 9443),
    [int]$OlderThan = 0,
    [string]$Reason = "Troubleshooting operation"
)

$ErrorActionPreference = "Continue"

# Configuration
$BackendServiceName = "ExcelAddin-Backend"  
$FrontendServiceName = "ExcelAddin Frontend"
$BackendPort = 5000
$FrontendPort = 3000
$IISPort = 9443
$LogDir = "C:\Logs\ExcelAddin"

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  ExcelAddin Comprehensive Troubleshooting" -ForegroundColor Cyan
Write-Host "  Environment: $Environment" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

# Helper Functions
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

function Test-Url {
    param(
        [string]$Name,
        [string]$Url,
        [int]$TimeoutSec = 10
    )
    
    if ($Verbose) { Write-Host "Testing $Name..." -ForegroundColor Yellow -NoNewline }
    
    try {
        $response = Invoke-WebRequest -Uri $Url -TimeoutSec $TimeoutSec -UseBasicParsing -ErrorAction Stop
        if ($Verbose) { Write-Host " ✅ SUCCESS (HTTP $($response.StatusCode))" -ForegroundColor Green }
        return @{
            Name = $Name
            Url = $Url
            Status = "SUCCESS"
            StatusCode = $response.StatusCode
            Error = $null
        }
    } catch {
        if ($Verbose) { 
            Write-Host " ❌ FAILED" -ForegroundColor Red
            Write-Host "   Error: $($_.Exception.Message)" -ForegroundColor Red
        }
        return @{
            Name = $Name
            Url = $Url
            Status = "FAILED"
            StatusCode = $null
            Error = $_.Exception.Message
        }
    }
}

function Get-ServiceDetails {
    param($ServiceName)
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($service) {
        $isActuallyRunning = $false
        if ($ServiceName -eq $FrontendServiceName) {
            $runningProcess = Get-NetTCPConnection -LocalPort $FrontendPort -ErrorAction SilentlyContinue | 
                Where-Object { $_.State -eq "Listen" }
            if ($runningProcess) {
                $processId = $runningProcess.OwningProcess
                $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
                if ($process -and $process.ProcessName -eq "node") {
                    $isActuallyRunning = $true
                }
            }
        } else {
            $isActuallyRunning = ($service.Status -eq "Running")
        }
        
        return @{
            Name = $service.Name
            Status = $service.Status
            StartType = $service.StartType
            DisplayName = $service.DisplayName
            IsActuallyRunning = $isActuallyRunning
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

function Write-TestResult {
    param($TestName, $Success, $Details = "")
    $status = if ($Success) { "✅ PASS" } else { "❌ FAIL" }
    $color = if ($Success) { "Green" } else { "Red" }
    Write-Host "  $TestName`: $status" -ForegroundColor $color
    if ($Details -and ($Debug -or -not $Success)) {
        Write-Host "    $Details" -ForegroundColor Gray
    }
}

# Main Test Functions
function Test-SystemInfo {
    Write-Host "SYSTEM INFORMATION" -ForegroundColor Yellow
    Write-Host "=" * 50
    
    $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
    Write-Host "  OS: $($osInfo.Caption) $($osInfo.Version)" -ForegroundColor Green
    Write-Host "  PowerShell: $($PSVersionTable.PSVersion)" -ForegroundColor Green
    Write-Host "  Hostname: $($env:COMPUTERNAME)" -ForegroundColor Green
    Write-Host ""
}

function Test-WindowsServices {
    Write-Host "WINDOWS SERVICES STATUS" -ForegroundColor Yellow
    Write-Host "=" * 50
    
    # Backend Service
    $backendService = Get-ServiceDetails $BackendServiceName
    if ($backendService) {
        Write-TestResult "Backend Service Exists" $true
        Write-TestResult "Backend Service Running" ($backendService.Status -eq "Running") "Status: $($backendService.Status)"
        
        if ($backendService.Status -ne "Running" -and $FixCommonIssues) {
            Write-Host "  Attempting to start backend service..." -ForegroundColor Yellow
            Start-Service -Name $BackendServiceName -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 5
            $backendService = Get-ServiceDetails $BackendServiceName
            Write-TestResult "Backend Service Started" ($backendService.Status -eq "Running") "New status: $($backendService.Status)"
        }
    } else {
        Write-TestResult "Backend Service Exists" $false "Service '$BackendServiceName' not found"
    }
    
    # Frontend Service
    $frontendService = Get-ServiceDetails $FrontendServiceName
    if ($frontendService) {
        Write-TestResult "Frontend Service Exists" $true
        Write-TestResult "Frontend Service Running" ($frontendService.Status -eq "Running") "Status: $($frontendService.Status)"
        Write-TestResult "Frontend Process Active" $frontendService.IsActuallyRunning "Node.js process listening on port $FrontendPort"
        
        if (-not $frontendService.IsActuallyRunning -and $FixCommonIssues) {
            Write-Host "  Attempting to restart frontend service..." -ForegroundColor Yellow
            Restart-Service -Name $FrontendServiceName -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 5
            $frontendService = Get-ServiceDetails $FrontendServiceName
            Write-TestResult "Frontend Service Restarted" $frontendService.IsActuallyRunning
        }
    } else {
        Write-TestResult "Frontend Service Exists" $false "Service '$FrontendServiceName' not found"
    }
    Write-Host ""
}

function Test-NetworkConnectivity {
    Write-Host "NETWORK CONNECTIVITY" -ForegroundColor Yellow  
    Write-Host "=" * 50
    
    # Port connectivity tests
    Write-TestResult "Backend Port ($BackendPort)" (Test-Port "localhost" $BackendPort)
    Write-TestResult "Frontend Port ($FrontendPort)" (Test-Port "localhost" $FrontendPort) 
    Write-TestResult "IIS Port ($IISPort)" (Test-Port "localhost" $IISPort)
    
    # URL endpoint tests
    $results = @()
    
    # Frontend endpoints
    $frontendUrls = @(
        @{ Name = "Frontend Health"; Url = "http://localhost:$FrontendPort/health" },
        @{ Name = "Frontend Taskpane"; Url = "http://localhost:$FrontendPort/excellence/taskpane.html" },
        @{ Name = "Frontend Assets"; Url = "http://localhost:$FrontendPort/assets/icon-16.png" }
    )
    
    # Backend endpoints  
    $backendUrls = @(
        @{ Name = "Backend Health"; Url = "http://localhost:$BackendPort/api/health" },
        @{ Name = "Backend Debug"; Url = "http://localhost:$BackendPort/api/debug" }
    )
    
    # IIS Proxy endpoints
    $proxyUrls = @(
        @{ Name = "IIS Proxy Root"; Url = "http://localhost:$IISPort/" },
        @{ Name = "IIS Proxy API"; Url = "http://localhost:$IISPort/api/health" },
        @{ Name = "IIS Proxy Frontend"; Url = "http://localhost:$IISPort/excellence/taskpane.html" }
    )
    
    Write-Host "  Frontend Endpoints:" -ForegroundColor Cyan
    foreach ($test in $frontendUrls) {
        $result = Test-Url -Name $test.Name -Url $test.Url -TimeoutSec 5
        Write-TestResult $test.Name ($result.Status -eq "SUCCESS") $result.Error
    }
    
    Write-Host "  Backend Endpoints:" -ForegroundColor Cyan  
    foreach ($test in $backendUrls) {
        $result = Test-Url -Name $test.Name -Url $test.Url -TimeoutSec 5
        Write-TestResult $test.Name ($result.Status -eq "SUCCESS") $result.Error
    }
    
    Write-Host "  IIS Proxy Endpoints:" -ForegroundColor Cyan
    foreach ($test in $proxyUrls) {
        $result = Test-Url -Name $test.Name -Url $test.Url -TimeoutSec 5
        Write-TestResult $test.Name ($result.Status -eq "SUCCESS") $result.Error
    }
    Write-Host ""
}

function Test-IISConfiguration {
    Write-Host "IIS CONFIGURATION" -ForegroundColor Yellow
    Write-Host "=" * 50
    
    try {
        Import-Module WebAdministration -ErrorAction SilentlyContinue
        
        # Check if IIS is running
        $iisService = Get-Service -Name "W3SVC" -ErrorAction SilentlyContinue
        Write-TestResult "IIS Service Running" ($iisService -and $iisService.Status -eq "Running") 
        
        # Check for ExcelAddin website/application
        $website = Get-Website | Where-Object { $_.Name -like "*ExcelAddin*" -or $_.bindings -like "*$IISPort*" }
        Write-TestResult "ExcelAddin Website Configured" ($website -ne $null)
        
        if ($website) {
            Write-TestResult "Website Running" ($website.State -eq "Started") "State: $($website.State)"
            Write-Host "    Website: $($website.Name)" -ForegroundColor Gray
            Write-Host "    Path: $($website.physicalPath)" -ForegroundColor Gray
            Write-Host "    Bindings: $($website.bindings.Collection.bindingInformation -join ', ')" -ForegroundColor Gray
        }
        
        # Check application pools
        $appPool = Get-WebAppPoolState | Where-Object { $_.Name -like "*ExcelAddin*" }
        if ($appPool) {
            Write-TestResult "Application Pool Running" ($appPool.Value -eq "Started") "State: $($appPool.Value)"
        }
        
    } catch {
        Write-TestResult "IIS Module Available" $false "WebAdministration module not available: $($_.Exception.Message)"
    }
    Write-Host ""
}

function Test-SSLConfiguration {
    Write-Host "SSL CONFIGURATION" -ForegroundColor Yellow  
    Write-Host "=" * 50
    
    try {
        # Check SSL certificate binding
        $sslBinding = netsh http show sslcert ipport=0.0.0.0:$IISPort 2>$null
        Write-TestResult "SSL Certificate Bound" ($LASTEXITCODE -eq 0) "Port $IISPort"
        
        # Test SSL endpoint if TestUrl provided
        if ($TestUrl) {
            $sslTest = Test-Url -Name "SSL Endpoint" -Url $TestUrl -TimeoutSec 10
            Write-TestResult "SSL Endpoint Accessible" ($sslTest.Status -eq "SUCCESS") $sslTest.Error
        }
        
        # Check certificate store for valid certificates
        $cert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { 
            $_.Subject -like "*intranet.local*" -and $_.NotAfter -gt (Get-Date) 
        } | Select-Object -First 1
        Write-TestResult "Valid SSL Certificate Available" ($cert -ne $null) 
        if ($cert) {
            Write-Host "    Subject: $($cert.Subject)" -ForegroundColor Gray
            Write-Host "    Expires: $($cert.NotAfter)" -ForegroundColor Gray
        }
        
    } catch {
        Write-TestResult "SSL Check Failed" $false $_.Exception.Message
    }
    Write-Host ""
}

function Test-LogsAndDiagnostics {
    Write-Host "LOGS AND DIAGNOSTICS" -ForegroundColor Yellow
    Write-Host "=" * 50
    
    # Check log directories
    Write-TestResult "Log Directory Exists" (Test-Path $LogDir)
    
    if (Test-Path $LogDir) {
        $logFiles = Get-ChildItem $LogDir -Recurse -File | Sort-Object LastWriteTime -Descending
        Write-TestResult "Log Files Present" ($logFiles.Count -gt 0) "$($logFiles.Count) files found"
        
        if ($logFiles.Count -gt 0) {
            Write-Host "    Recent log files:" -ForegroundColor Gray
            $logFiles | Select-Object -First 5 | ForEach-Object {
                Write-Host "      $($_.Name) - $($_.LastWriteTime)" -ForegroundColor Gray
            }
        }
    }
    
    # Check Windows Event Logs  
    try {
        $recentErrors = Get-WinEvent -LogName Application -FilterXPath "*[System[(Level=2 or Level=3) and TimeCreated[timediff(@SystemTime) <= 86400000]]]" -MaxEvents 5 -ErrorAction SilentlyContinue | 
            Where-Object { $_.ProviderName -like "*ExcelAddin*" }
        Write-TestResult "Recent Service Errors" ($recentErrors.Count -eq 0) "$($recentErrors.Count) errors in last 24h"
    } catch {
        Write-TestResult "Event Log Check" $false "Unable to check event logs: $($_.Exception.Message)"
    }
    Write-Host ""
}

# Action Functions
function Restart-AllServices {
    param([string]$Reason = "Manual restart")
    
    Write-Host "RESTARTING SERVICES" -ForegroundColor Yellow
    Write-Host "=" * 50
    Write-Host "Reason: $Reason" -ForegroundColor Gray
    Write-Host ""
    
    # Restart backend service
    try {
        Write-Host "Restarting backend service..." -ForegroundColor Yellow
        Restart-Service -Name $BackendServiceName -Force -ErrorAction Stop
        Start-Sleep -Seconds 3
        $backendService = Get-Service -Name $BackendServiceName
        Write-TestResult "Backend Service Restarted" ($backendService.Status -eq "Running")
    } catch {
        Write-TestResult "Backend Service Restart" $false $_.Exception.Message
    }
    
    # Restart frontend service  
    try {
        Write-Host "Restarting frontend service..." -ForegroundColor Yellow
        Restart-Service -Name $FrontendServiceName -Force -ErrorAction Stop
        Start-Sleep -Seconds 5
        $frontendService = Get-ServiceDetails $FrontendServiceName
        Write-TestResult "Frontend Service Restarted" $frontendService.IsActuallyRunning
    } catch {
        Write-TestResult "Frontend Service Restart" $false $_.Exception.Message
    }
    Write-Host ""
}

function Clear-OldLogs {
    param([int]$DaysOld = 30)
    
    Write-Host "CLEARING OLD LOG FILES" -ForegroundColor Yellow
    Write-Host "=" * 50
    Write-Host "Removing files older than $DaysOld days" -ForegroundColor Gray
    Write-Host ""
    
    if (Test-Path $LogDir) {
        $cutoffDate = (Get-Date).AddDays(-$DaysOld)
        $oldFiles = Get-ChildItem $LogDir -Recurse -File | Where-Object { $_.LastWriteTime -lt $cutoffDate }
        
        if ($oldFiles.Count -gt 0) {
            Write-Host "Found $($oldFiles.Count) old files to remove:" -ForegroundColor Yellow
            $oldFiles | ForEach-Object {
                Write-Host "  Removing: $($_.FullName)" -ForegroundColor Gray
                Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
            }
            Write-Host "✅ Log cleanup completed" -ForegroundColor Green
        } else {
            Write-Host "✅ No old log files found" -ForegroundColor Green
        }
    } else {
        Write-Host "⚠️  Log directory not found: $LogDir" -ForegroundColor Yellow
    }
    Write-Host ""
}

function Check-PortUsage {
    param([array]$PortsToCheck)
    
    Write-Host "PORT USAGE CHECK" -ForegroundColor Yellow
    Write-Host "=" * 50
    
    foreach ($port in $PortsToCheck) {
        try {
            $connection = Get-NetTCPConnection -LocalPort $port -ErrorAction SilentlyContinue | 
                Where-Object { $_.State -eq "Listen" }
            
            if ($connection) {
                $process = Get-Process -Id $connection.OwningProcess -ErrorAction SilentlyContinue
                Write-TestResult "Port $port Available" $false "Used by $($process.ProcessName) (PID: $($process.Id))"
            } else {
                Write-TestResult "Port $port Available" $true "Port is free"
            }
        } catch {
            Write-TestResult "Port $port Check" $false "Error checking port: $($_.Exception.Message)"
        }
    }
    Write-Host ""
}

function Reset-IISConfiguration {
    Write-Host "RESETTING IIS CONFIGURATION" -ForegroundColor Yellow
    Write-Host "=" * 50
    Write-Host "⚠️  This will reset IIS configuration for ExcelAddin" -ForegroundColor Yellow
    
    try {
        Import-Module WebAdministration -ErrorAction Stop
        
        # Stop website if it exists
        $website = Get-Website | Where-Object { $_.Name -like "*ExcelAddin*" }
        if ($website) {
            Write-Host "Stopping website: $($website.Name)" -ForegroundColor Yellow
            Stop-Website -Name $website.Name -ErrorAction SilentlyContinue
        }
        
        # Stop application pool
        $appPool = Get-WebAppPoolState | Where-Object { $_.Name -like "*ExcelAddin*" }
        if ($appPool) {
            Write-Host "Stopping application pool: $($appPool.Name)" -ForegroundColor Yellow  
            Stop-WebAppPool -Name $appPool.Name -ErrorAction SilentlyContinue
        }
        
        Write-Host "✅ IIS reset completed. Re-run deploy-iis.ps1 to reconfigure." -ForegroundColor Green
        
    } catch {
        Write-TestResult "IIS Reset" $false $_.Exception.Message
    }
    Write-Host ""
}

function Apply-CommonFixes {
    Write-Host "APPLYING COMMON FIXES" -ForegroundColor Yellow
    Write-Host "=" * 50
    
    # Fix 1: Ensure services are running
    Write-Host "1. Checking and starting services..." -ForegroundColor Yellow
    $backendService = Get-Service -Name $BackendServiceName -ErrorAction SilentlyContinue
    if ($backendService -and $backendService.Status -ne "Running") {
        Start-Service -Name $BackendServiceName -ErrorAction SilentlyContinue
        Write-Host "   ✅ Started backend service" -ForegroundColor Green
    }
    
    $frontendService = Get-Service -Name $FrontendServiceName -ErrorAction SilentlyContinue
    if ($frontendService -and $frontendService.Status -ne "Running") {
        Start-Service -Name $FrontendServiceName -ErrorAction SilentlyContinue
        Write-Host "   ✅ Started frontend service" -ForegroundColor Green
    }
    
    # Fix 2: Clear temporary files
    Write-Host "2. Clearing temporary files..." -ForegroundColor Yellow
    $tempDirs = @(
        "$env:TEMP\ExcelAddin*",
        "C:\Windows\Temp\ExcelAddin*"
    )
    foreach ($pattern in $tempDirs) {
        Get-ChildItem $pattern -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }
    Write-Host "   ✅ Temporary files cleared" -ForegroundColor Green
    
    # Fix 3: Reset IIS if needed
    Write-Host "3. Checking IIS configuration..." -ForegroundColor Yellow
    try {
        $iisService = Get-Service -Name "W3SVC" -ErrorAction SilentlyContinue
        if ($iisService -and $iisService.Status -ne "Running") {
            Start-Service -Name "W3SVC" -ErrorAction SilentlyContinue
            Write-Host "   ✅ Started IIS service" -ForegroundColor Green
        }
    } catch {
        Write-Host "   ⚠️  Could not check IIS: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    Write-Host "✅ Common fixes applied" -ForegroundColor Green
    Write-Host ""
}

# Main Execution Logic
if ($TestAll) {
    Test-SystemInfo
    Test-WindowsServices
    Test-NetworkConnectivity
    Test-IISConfiguration
    Test-SSLConfiguration
    Test-LogsAndDiagnostics
} 

if ($TestServices) {
    Test-SystemInfo
    Test-WindowsServices
}

if ($TestConnectivity) {
    Test-NetworkConnectivity
}

if ($TestIIS) {
    Test-IISConfiguration  
}

if ($TestSSL) {
    Test-SSLConfiguration
}

if ($CheckPorts) {
    Check-PortUsage -PortsToCheck $Ports
}

if ($FixCommonIssues) {
    Apply-CommonFixes
}

if ($RestartServices) {
    Restart-AllServices -Reason $Reason
}

if ($ClearLogs) {
    if ($OlderThan -gt 0) {
        Clear-OldLogs -DaysOld $OlderThan
    } else {
        Clear-OldLogs
    }
}

if ($ResetIIS) {
    Reset-IISConfiguration
}

# If no specific test was requested, show help
if (-not ($TestAll -or $TestServices -or $TestConnectivity -or $TestIIS -or $TestSSL -or $CheckPorts -or $FixCommonIssues -or $RestartServices -or $ClearLogs -or $ResetIIS)) {
    Write-Host "ExcelAddin Troubleshooting Script Usage:" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Test Commands:" -ForegroundColor Yellow
    Write-Host "  -TestAll              Run all diagnostic tests"
    Write-Host "  -TestServices         Test Windows services status"
    Write-Host "  -TestConnectivity     Test network connectivity and endpoints"
    Write-Host "  -TestIIS              Test IIS configuration"
    Write-Host "  -TestSSL              Test SSL certificate configuration"
    Write-Host "  -CheckPorts           Check if ports are available"
    Write-Host ""
    Write-Host "Action Commands:" -ForegroundColor Yellow
    Write-Host "  -FixCommonIssues      Apply automated fixes for common problems"
    Write-Host "  -RestartServices      Restart all ExcelAddin services"  
    Write-Host "  -ClearLogs            Clear old log files (use -OlderThan days)"
    Write-Host "  -ResetIIS             Reset IIS configuration (requires re-deployment)"
    Write-Host ""
    Write-Host "Options:" -ForegroundColor Yellow
    Write-Host "  -Environment          Environment (staging|production) - default: staging"
    Write-Host "  -TestUrl              URL to test for SSL connectivity"
    Write-Host "  -Verbose              Show detailed test output"
    Write-Host "  -Debug                Show debug information"
    Write-Host "  -ServiceName          Specific service name for targeted operations"
    Write-Host "  -Ports                Array of ports to check (default: 3000,5000,9443)"
    Write-Host "  -OlderThan            Days old for log cleanup (default: 30)"
    Write-Host "  -Reason               Reason for service restart operations"
    Write-Host ""
    Write-Host "Examples:" -ForegroundColor Green
    Write-Host "  .\troubleshooting.ps1 -TestAll -Verbose"
    Write-Host "  .\troubleshooting.ps1 -FixCommonIssues"
    Write-Host "  .\troubleshooting.ps1 -TestConnectivity -Environment production"
    Write-Host "  .\troubleshooting.ps1 -ClearLogs -OlderThan 7"
    Write-Host "  .\troubleshooting.ps1 -CheckPorts -Ports @(3000,5000,9443)"
    Write-Host ""
}

Write-Host "Troubleshooting completed at $(Get-Date)" -ForegroundColor Green