# IIS Proxy Management Script for ExcelAddin
# Provides easy management commands for the IIS proxy service

param(
    [ValidateSet("start", "stop", "restart", "status", "test", "logs")]
    [Parameter(Mandatory=$true)]
    [string]$Action,
    
    [string]$SiteName = "ExcelAddin-Proxy",
    [string]$AppPoolName = "ExcelAddin-Proxy",
    [int]$Port = 9443,
    [string]$ServerFQDN = "server-vs81t.intranet.local"
)

$ErrorActionPreference = "Stop"

Write-Host "IIS Proxy Management - $($Action.ToUpper())" -ForegroundColor Cyan
Write-Host "Site: $SiteName" -ForegroundColor Gray
Write-Host ""

try {
    # Import WebAdministration module if available
    if (Get-Module -ListAvailable -Name WebAdministration) {
        Import-Module WebAdministration -ErrorAction SilentlyContinue
    }

    switch ($Action) {
        "start" {
            Write-Host "Starting IIS proxy..." -ForegroundColor Yellow
            
            # Start app pool first
            $pool = Get-WebAppPool -Name $AppPoolName -ErrorAction SilentlyContinue
            if ($pool) {
                if ($pool.State -ne "Started") {
                    Start-WebAppPool -Name $AppPoolName
                    Write-Host "  ✅ Application pool started" -ForegroundColor Green
                } else {
                    Write-Host "  ✅ Application pool already running" -ForegroundColor Green
                }
            } else {
                Write-Warning "  Application pool '$AppPoolName' not found"
            }
            
            # Start website
            $site = Get-Website -Name $SiteName -ErrorAction SilentlyContinue
            if ($site) {
                if ($site.State -ne "Started") {
                    Start-Website -Name $SiteName
                    Write-Host "  ✅ Website started" -ForegroundColor Green
                } else {
                    Write-Host "  ✅ Website already running" -ForegroundColor Green
                }
            } else {
                Write-Warning "  Website '$SiteName' not found"
            }
        }
        
        "stop" {
            Write-Host "Stopping IIS proxy..." -ForegroundColor Yellow
            
            # Stop website
            $site = Get-Website -Name $SiteName -ErrorAction SilentlyContinue
            if ($site -and $site.State -eq "Started") {
                Stop-Website -Name $SiteName
                Write-Host "  ✅ Website stopped" -ForegroundColor Green
            } else {
                Write-Host "  ✅ Website already stopped" -ForegroundColor Green
            }
            
            # Stop app pool
            $pool = Get-WebAppPool -Name $AppPoolName -ErrorAction SilentlyContinue
            if ($pool -and $pool.State -eq "Started") {
                Stop-WebAppPool -Name $AppPoolName
                Write-Host "  ✅ Application pool stopped" -ForegroundColor Green
            } else {
                Write-Host "  ✅ Application pool already stopped" -ForegroundColor Green
            }
        }
        
        "restart" {
            Write-Host "Restarting IIS proxy..." -ForegroundColor Yellow
            
            # Stop first
            & $PSCommandPath -Action stop -SiteName $SiteName -AppPoolName $AppPoolName -Port $Port -ServerFQDN $ServerFQDN
            
            Start-Sleep -Seconds 2
            
            # Then start
            & $PSCommandPath -Action start -SiteName $SiteName -AppPoolName $AppPoolName -Port $Port -ServerFQDN $ServerFQDN
        }
        
        "status" {
            Write-Host "Checking IIS proxy status..." -ForegroundColor Yellow
            Write-Host ""
            
            # Check application pool
            $pool = Get-WebAppPool -Name $AppPoolName -ErrorAction SilentlyContinue
            if ($pool) {
                $poolColor = if ($pool.State -eq "Started") { "Green" } else { "Red" }
                Write-Host "Application Pool:" -ForegroundColor Cyan
                Write-Host "  Name: $($pool.Name)"
                Write-Host "  State: $($pool.State)" -ForegroundColor $poolColor
                Write-Host "  Runtime: $($pool.ManagedRuntimeVersion)"
                Write-Host "  Identity: $($pool.ProcessModel.IdentityType)"
                Write-Host ""
            } else {
                Write-Host "Application Pool: NOT FOUND" -ForegroundColor Red
                Write-Host ""
            }
            
            # Check website
            $site = Get-Website -Name $SiteName -ErrorAction SilentlyContinue
            if ($site) {
                $siteColor = if ($site.State -eq "Started") { "Green" } else { "Red" }
                Write-Host "Website:" -ForegroundColor Cyan
                Write-Host "  Name: $($site.Name)"
                Write-Host "  State: $($site.State)" -ForegroundColor $siteColor
                Write-Host "  Physical Path: $($site.PhysicalPath)"
                Write-Host "  App Pool: $($site.ApplicationPool)"
                Write-Host ""
                
                # Check bindings
                $bindings = Get-WebBinding -Name $SiteName
                Write-Host "Bindings:" -ForegroundColor Cyan
                foreach ($binding in $bindings) {
                    $bindingInfo = "$($binding.Protocol)://*:$($binding.BindingInformation.Split(':')[1])"
                    if ($binding.CertificateHash) {
                        $bindingInfo += " (SSL: $($binding.CertificateHash.Substring(0,8))...)"
                    }
                    Write-Host "  $bindingInfo"
                }
                Write-Host ""
            } else {
                Write-Host "Website: NOT FOUND" -ForegroundColor Red
                Write-Host ""
            }
            
            # Check port
            Write-Host "Port Check:" -ForegroundColor Cyan
            $portCheck = Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue
            if ($portCheck) {
                Write-Host "  Port $Port is in use" -ForegroundColor Green
                $process = Get-Process -Id $portCheck.OwningProcess -ErrorAction SilentlyContinue
                if ($process) {
                    Write-Host "  Process: $($process.ProcessName) (PID: $($process.Id))"
                }
            } else {
                Write-Host "  Port $Port is not in use" -ForegroundColor Red
            }
        }
        
        "test" {
            Write-Host "Testing IIS proxy..." -ForegroundColor Yellow
            Write-Host ""
            
            $protocol = if ($Port -eq 443 -or $Port -eq 9443) { "https" } else { "http" }
            $testUrls = @(
                "${protocol}://localhost:${Port}/",
                "${protocol}://localhost:${Port}/excellence/taskpane.html",
                "${protocol}://localhost:${Port}/api/health"
            )
            
            foreach ($url in $testUrls) {
                try {
                    Write-Host "Testing: $url" -ForegroundColor Cyan
                    $response = Invoke-WebRequest -Uri $url -TimeoutSec 10 -UseBasicParsing -ErrorAction Stop
                    Write-Host "  ✅ HTTP $($response.StatusCode) - $($response.Content.Length) bytes" -ForegroundColor Green
                } catch {
                    Write-Host "  ❌ Failed: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            
            Write-Host ""
            Write-Host "External URLs:" -ForegroundColor Cyan
            Write-Host "  Status Page: ${protocol}://${ServerFQDN}:${Port}/"
            Write-Host "  Excel Taskpane: ${protocol}://${ServerFQDN}:${Port}/excellence/taskpane.html"
            Write-Host "  API Health: ${protocol}://${ServerFQDN}:${Port}/api/health"
        }
        
        "logs" {
            Write-Host "Checking IIS logs..." -ForegroundColor Yellow
            Write-Host ""
            
            # IIS logs
            $iisLogPath = "C:\inetpub\logs\LogFiles\W3SVC*"
            $logDirs = Get-ChildItem $iisLogPath -Directory -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
            
            if ($logDirs) {
                Write-Host "Recent IIS log files:" -ForegroundColor Cyan
                foreach ($logDir in $logDirs | Select-Object -First 3) {
                    $logFiles = Get-ChildItem $logDir.FullName -Filter "*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 2
                    foreach ($logFile in $logFiles) {
                        Write-Host "  $($logFile.FullName) ($($logFile.Length) bytes, modified $($logFile.LastWriteTime))"
                        
                        # Show last few lines
                        $content = Get-Content $logFile.FullName -Tail 5 -ErrorAction SilentlyContinue
                        if ($content) {
                            $content | ForEach-Object { Write-Host "    $_" -ForegroundColor Gray }
                        }
                        Write-Host ""
                    }
                }
            } else {
                Write-Host "No IIS log files found at $iisLogPath" -ForegroundColor Yellow
            }
            
            # Windows Event Log
            Write-Host "Recent Windows Event Log entries:" -ForegroundColor Cyan
            try {
                $events = Get-WinEvent -LogName System -FilterXPath "*[System[Provider[@Name='Microsoft-Windows-IIS-IISReset'] or Provider[@Name='Microsoft-Windows-WAS']]]" -MaxEvents 5 -ErrorAction SilentlyContinue
                if ($events) {
                    foreach ($event in $events) {
                        Write-Host "  [$($event.TimeCreated)] $($event.LevelDisplayName): $($event.Message.Split("`n")[0])" -ForegroundColor Gray
                    }
                } else {
                    Write-Host "  No recent IIS-related events found" -ForegroundColor Gray
                }
            } catch {
                Write-Host "  Could not access Windows Event Log" -ForegroundColor Yellow
            }
        }
    }
    
    Write-Host ""
    Write-Host "Operation completed successfully!" -ForegroundColor Green
    
} catch {
    Write-Host ""
    Write-Host "Operation failed: $($_.Exception.Message)" -ForegroundColor Red
    
    if ($_.Exception.Message -like "*WebAdministration*") {
        Write-Host ""
        Write-Host "IIS management tools may not be installed." -ForegroundColor Yellow
        Write-Host "Install using: Enable-WindowsOptionalFeature -Online -FeatureName IIS-ManagementConsole" -ForegroundColor Yellow
    }
    
    exit 1
}