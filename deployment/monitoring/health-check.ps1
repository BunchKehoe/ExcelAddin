# Health Check and Monitoring Script for Excel Add-in
# This PowerShell script monitors the health of both frontend and backend services

param(
    [string]$DomainName = "your-staging-domain.com",
    [string]$BackendPort = "5000",
    [string]$ServiceName = "ExcelAddinBackend",
    [switch]$Detailed,
    [switch]$SendAlerts,
    [string]$LogPath = "C:\Logs\ExcelAddin\health-check.log"
)

# Ensure log directory exists
$logDir = Split-Path $LogPath -Parent
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Write-Host $logMessage
    Add-Content -Path $LogPath -Value $logMessage
}

function Test-BackendService {
    Write-Log "Checking backend service status..."
    
    try {
        $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if ($service) {
            $status = $service.Status
            Write-Log "Service $ServiceName status: $status"
            
            if ($status -eq 'Running') {
                return @{ Status = "Healthy"; Message = "Service is running"; Details = $service }
            } else {
                return @{ Status = "Unhealthy"; Message = "Service is not running: $status"; Details = $service }
            }
        } else {
            return @{ Status = "Unhealthy"; Message = "Service not found"; Details = $null }
        }
    } catch {
        return @{ Status = "Error"; Message = "Error checking service: $_"; Details = $null }
    }
}

function Test-BackendAPI {
    Write-Log "Testing backend API endpoint..."
    
    try {
        $uri = "http://127.0.0.1:$BackendPort/api/health"
        $response = Invoke-WebRequest -Uri $uri -TimeoutSec 10 -UseBasicParsing
        
        if ($response.StatusCode -eq 200) {
            $content = $response.Content | ConvertFrom-Json -ErrorAction SilentlyContinue
            Write-Log "Backend API responded with status 200"
            return @{ 
                Status = "Healthy"; 
                Message = "API is responding"; 
                Details = @{
                    StatusCode = $response.StatusCode
                    ResponseTime = $response.Headers.'X-Response-Time'
                    Content = $content
                }
            }
        } else {
            return @{ Status = "Unhealthy"; Message = "API returned status: $($response.StatusCode)"; Details = $response }
        }
    } catch {
        return @{ Status = "Unhealthy"; Message = "API not responding: $_"; Details = $null }
    }
}

function Test-FrontendHTTPS {
    Write-Log "Testing frontend HTTPS endpoint..."
    
    try {
        $uri = "https://$DomainName/health"
        $response = Invoke-WebRequest -Uri $uri -TimeoutSec 10 -UseBasicParsing -SkipCertificateCheck
        
        if ($response.StatusCode -eq 200) {
            Write-Log "Frontend HTTPS endpoint responded with status 200"
            return @{ 
                Status = "Healthy"; 
                Message = "HTTPS endpoint is responding"; 
                Details = @{
                    StatusCode = $response.StatusCode
                    Headers = $response.Headers
                }
            }
        } else {
            return @{ Status = "Unhealthy"; Message = "Frontend returned status: $($response.StatusCode)"; Details = $response }
        }
    } catch {
        return @{ Status = "Unhealthy"; Message = "Frontend not responding: $_"; Details = $null }
    }
}

function Test-SSLCertificate {
    Write-Log "Checking SSL certificate..."
    
    try {
        $request = [System.Net.WebRequest]::Create("https://$DomainName")
        $request.Method = "HEAD"
        $request.Timeout = 10000
        
        $response = $request.GetResponse()
        $cert = $request.ServicePoint.Certificate
        
        if ($cert) {
            $expiryDate = [DateTime]::Parse($cert.GetExpirationDateString())
            $daysUntilExpiry = ($expiryDate - (Get-Date)).Days
            
            Write-Log "SSL certificate expires on: $expiryDate ($daysUntilExpiry days from now)"
            
            if ($daysUntilExpiry -lt 30) {
                return @{ 
                    Status = "Warning"; 
                    Message = "Certificate expires in $daysUntilExpiry days"; 
                    Details = @{
                        Expiry = $expiryDate
                        DaysUntilExpiry = $daysUntilExpiry
                        Subject = $cert.Subject
                    }
                }
            } else {
                return @{ 
                    Status = "Healthy"; 
                    Message = "Certificate is valid"; 
                    Details = @{
                        Expiry = $expiryDate
                        DaysUntilExpiry = $daysUntilExpiry
                        Subject = $cert.Subject
                    }
                }
            }
        } else {
            return @{ Status = "Unhealthy"; Message = "No SSL certificate found"; Details = $null }
        }
    } catch {
        return @{ Status = "Error"; Message = "Error checking SSL certificate: $_"; Details = $null }
    }
}

function Test-NginxProcess {
    Write-Log "Checking nginx process..."
    
    try {
        $nginxProcesses = Get-Process -Name "nginx" -ErrorAction SilentlyContinue
        
        if ($nginxProcesses) {
            $processCount = $nginxProcesses.Count
            Write-Log "Found $processCount nginx process(es)"
            return @{ 
                Status = "Healthy"; 
                Message = "nginx is running ($processCount processes)"; 
                Details = $nginxProcesses
            }
        } else {
            return @{ Status = "Unhealthy"; Message = "nginx is not running"; Details = $null }
        }
    } catch {
        return @{ Status = "Error"; Message = "Error checking nginx: $_"; Details = $null }
    }
}

function Test-DiskSpace {
    Write-Log "Checking disk space..."
    
    try {
        $drives = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
        $warnings = @()
        $errors = @()
        
        foreach ($drive in $drives) {
            $freeSpaceGB = [math]::Round($drive.FreeSpace / 1GB, 2)
            $totalSpaceGB = [math]::Round($drive.Size / 1GB, 2)
            $freeSpacePercent = [math]::Round(($drive.FreeSpace / $drive.Size) * 100, 2)
            
            Write-Log "Drive $($drive.DeviceID) - Free: $freeSpaceGB GB ($freeSpacePercent%) of $totalSpaceGB GB"
            
            if ($freeSpacePercent -lt 10) {
                $errors += "Drive $($drive.DeviceID) is critically low on space: $freeSpacePercent% free"
            } elseif ($freeSpacePercent -lt 20) {
                $warnings += "Drive $($drive.DeviceID) is running low on space: $freeSpacePercent% free"
            }
        }
        
        if ($errors.Count -gt 0) {
            return @{ Status = "Unhealthy"; Message = $errors -join "; "; Details = $drives }
        } elseif ($warnings.Count -gt 0) {
            return @{ Status = "Warning"; Message = $warnings -join "; "; Details = $drives }
        } else {
            return @{ Status = "Healthy"; Message = "Disk space is adequate"; Details = $drives }
        }
    } catch {
        return @{ Status = "Error"; Message = "Error checking disk space: $_"; Details = $null }
    }
}

function Send-Alert {
    param(
        [string]$Subject,
        [string]$Body
    )
    
    # Implement alert mechanism (email, webhook, etc.)
    # This is a placeholder - customize based on your alerting system
    Write-Log "ALERT: $Subject - $Body" "ALERT"
    
    # Example: Send to Windows Event Log
    try {
        Write-EventLog -LogName Application -Source "ExcelAddin" -EventID 1001 -EntryType Warning -Message "$Subject`n$Body"
    } catch {
        # Event source might not exist, create it or log differently
        Write-Log "Could not write to event log: $_" "WARNING"
    }
}

# Main health check execution
Write-Log "Starting health check for Excel Add-in"

$healthResults = @{
    BackendService = Test-BackendService
    BackendAPI = Test-BackendAPI
    FrontendHTTPS = Test-FrontendHTTPS
    SSLCertificate = Test-SSLCertificate
    NginxProcess = Test-NginxProcess
    DiskSpace = Test-DiskSpace
}

# Overall health assessment
$overallStatus = "Healthy"
$issues = @()
$warnings = @()

foreach ($check in $healthResults.Keys) {
    $result = $healthResults[$check]
    
    switch ($result.Status) {
        "Unhealthy" {
            $overallStatus = "Unhealthy"
            $issues += "$check: $($result.Message)"
        }
        "Error" {
            $overallStatus = "Unhealthy"
            $issues += "$check: $($result.Message)"
        }
        "Warning" {
            if ($overallStatus -eq "Healthy") { $overallStatus = "Warning" }
            $warnings += "$check: $($result.Message)"
        }
    }
}

# Display results
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "Excel Add-in Health Check Results" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan

foreach ($check in $healthResults.Keys) {
    $result = $healthResults[$check]
    $color = switch ($result.Status) {
        "Healthy" { "Green" }
        "Warning" { "Yellow" }
        "Unhealthy" { "Red" }
        "Error" { "Red" }
        default { "White" }
    }
    
    Write-Host "[$($result.Status.ToUpper().PadRight(9))] $check" -ForegroundColor $color
    Write-Host "  $($result.Message)" -ForegroundColor Gray
    
    if ($Detailed -and $result.Details) {
        Write-Host "  Details: $($result.Details | ConvertTo-Json -Compress)" -ForegroundColor Gray
    }
}

Write-Host ""
Write-Host "Overall Status: $overallStatus" -ForegroundColor $(
    switch ($overallStatus) {
        "Healthy" { "Green" }
        "Warning" { "Yellow" }
        "Unhealthy" { "Red" }
        default { "White" }
    }
)

Write-Log "Health check completed. Overall status: $overallStatus"

# Send alerts if enabled and there are issues
if ($SendAlerts) {
    if ($overallStatus -eq "Unhealthy" -and $issues.Count -gt 0) {
        $alertSubject = "Excel Add-in Health Check - CRITICAL"
        $alertBody = "The following critical issues were detected:`n`n" + ($issues -join "`n")
        Send-Alert -Subject $alertSubject -Body $alertBody
    } elseif ($overallStatus -eq "Warning" -and $warnings.Count -gt 0) {
        $alertSubject = "Excel Add-in Health Check - WARNING"
        $alertBody = "The following warnings were detected:`n`n" + ($warnings -join "`n")
        Send-Alert -Subject $alertSubject -Body $alertBody
    }
}

# Exit with appropriate code
switch ($overallStatus) {
    "Healthy" { exit 0 }
    "Warning" { exit 1 }
    "Unhealthy" { exit 2 }
    default { exit 3 }
}