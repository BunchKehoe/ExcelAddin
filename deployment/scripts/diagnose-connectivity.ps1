# Diagnose connectivity issues with nginx and Excel add-in
# Usage: .\diagnose-connectivity.ps1 [-DomainName "localhost:8443"] [-Detailed]

param(
    [string]$DomainName = "localhost:8443",
    [switch]$Detailed = $false
)

Write-Host "🔍 Excel Add-in Connectivity Diagnostics" -ForegroundColor Cyan
Write-Host "Testing connectivity to: $DomainName" -ForegroundColor Yellow
Write-Host ""

# Test 1: Check if nginx is running
Write-Host "1. Checking nginx service status..." -ForegroundColor Green
try {
    $nginxService = Get-Service nginx -ErrorAction Stop
    Write-Host "✅ nginx service status: $($nginxService.Status)" -ForegroundColor Green
    if ($nginxService.Status -ne "Running") {
        Write-Host "❌ nginx service is not running!" -ForegroundColor Red
        Write-Host "💡 Try: Start-Service nginx" -ForegroundColor Yellow
    }
} catch {
    Write-Host "❌ nginx service not found or not installed" -ForegroundColor Red
    Write-Host "💡 Check if nginx is installed as a Windows service" -ForegroundColor Yellow
}

# Test 2: Check if ports are listening
Write-Host ""
Write-Host "2. Checking if nginx is listening on required ports..." -ForegroundColor Green
$listening = netstat -an | findstr ":8443"
if ($listening) {
    Write-Host "✅ Port 8443 is listening:" -ForegroundColor Green
    Write-Host $listening -ForegroundColor White
} else {
    Write-Host "❌ Port 8443 is not listening" -ForegroundColor Red
    Write-Host "💡 nginx may not be configured to listen on port 8443" -ForegroundColor Yellow
}

# Test 3: Check nginx configuration
Write-Host ""
Write-Host "3. Testing nginx configuration syntax..." -ForegroundColor Green
if (Test-Path "C:\nginx\nginx.exe") {
    try {
        $configTest = & "C:\nginx\nginx.exe" -t 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✅ nginx configuration syntax is valid" -ForegroundColor Green
        } else {
            Write-Host "❌ nginx configuration has errors:" -ForegroundColor Red
            Write-Host $configTest -ForegroundColor Red
        }
    } catch {
        Write-Host "❌ Error running nginx config test: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "❌ nginx.exe not found at C:\nginx\nginx.exe" -ForegroundColor Red
}

# Test 4: Check certificate files
Write-Host ""
Write-Host "4. Checking SSL certificate files..." -ForegroundColor Green
$certFiles = @("C:\Cert\server.crt", "C:\Cert\server.key")
foreach ($file in $certFiles) {
    if (Test-Path $file) {
        Write-Host "✅ Found: $file" -ForegroundColor Green
    } else {
        Write-Host "❌ Missing: $file" -ForegroundColor Red
        Write-Host "💡 SSL certificates may not be properly deployed" -ForegroundColor Yellow
    }
}

# Test 5: Check frontend files deployment
Write-Host ""
Write-Host "5. Checking frontend files deployment..." -ForegroundColor Green
$requiredFiles = @(
    "C:\inetpub\wwwroot\ExcelAddin\dist\taskpane.html",
    "C:\inetpub\wwwroot\ExcelAddin\dist\commands.html",
    "C:\inetpub\wwwroot\ExcelAddin\dist\manifest.xml"
)
$missingFiles = @()
foreach ($file in $requiredFiles) {
    if (Test-Path $file) {
        Write-Host "✅ Found: $(Split-Path $file -Leaf)" -ForegroundColor Green
    } else {
        Write-Host "❌ Missing: $file" -ForegroundColor Red
        $missingFiles += $file
    }
}

if ($missingFiles.Count -gt 0) {
    Write-Host "💡 Frontend files may not be deployed. Run: npm run build:staging" -ForegroundColor Yellow
}

# Test 6: Attempt connectivity tests
Write-Host ""
Write-Host "6. Testing HTTP/HTTPS connectivity..." -ForegroundColor Green

# Determine PowerShell version for SSL handling
$psVersion = $PSVersionTable.PSVersion.Major

Write-Host "PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Cyan

# Test backend directly (should always work if backend is running)
Write-Host ""
Write-Host "Testing backend API directly (HTTP)..." -ForegroundColor Yellow
try {
    $backendResponse = Invoke-WebRequest -Uri "http://localhost:5000/api/health" -UseBasicParsing -TimeoutSec 10
    Write-Host "✅ Backend API accessible: HTTP $($backendResponse.StatusCode)" -ForegroundColor Green
} catch {
    Write-Host "❌ Backend API not accessible: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "💡 Check if backend service is running: Get-Service ExcelAddinBackend" -ForegroundColor Yellow
}

# Test HTTPS endpoints
Write-Host ""
Write-Host "Testing HTTPS endpoints through nginx..." -ForegroundColor Yellow

$testUrls = @(
    "https://$DomainName/excellence/api/health",
    "https://$DomainName/excellence/taskpane.html"
)

foreach ($url in $testUrls) {
    Write-Host "Testing: $url" -ForegroundColor Cyan
    
    try {
        if ($psVersion -ge 6) {
            # PowerShell 6.0+ with -SkipCertificateCheck
            $response = Invoke-WebRequest -Uri $url -SkipCertificateCheck -UseBasicParsing -TimeoutSec 10
            Write-Host "✅ Success: HTTP $($response.StatusCode)" -ForegroundColor Green
        } else {
            # Windows PowerShell 5.1 - ignore SSL certs temporarily
            [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
            
            $response = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 10
            Write-Host "✅ Success: HTTP $($response.StatusCode)" -ForegroundColor Green
            
            # Reset certificate validation
            [System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
        }
    } catch {
        Write-Host "❌ Failed: $($_.Exception.Message)" -ForegroundColor Red
        
        # Provide specific guidance based on error
        if ($_.Exception.Message -contains "Unable to connect") {
            Write-Host "💡 Connection refused - nginx may not be listening on port 8443" -ForegroundColor Yellow
        } elseif ($_.Exception.Message -contains "SSL") {
            Write-Host "💡 SSL certificate issue - check certificate files and configuration" -ForegroundColor Yellow
        } elseif ($_.Exception.Message -contains "404") {
            Write-Host "💡 nginx is running but files not found - check frontend deployment" -ForegroundColor Yellow
        }
    }
}

# Test 7: Check nginx error logs
Write-Host ""
Write-Host "7. Recent nginx error logs (last 10 lines)..." -ForegroundColor Green
$errorLogPath = "C:\nginx\logs\error.log"
if (Test-Path $errorLogPath) {
    try {
        $errorLines = Get-Content $errorLogPath -Tail 10 -ErrorAction Stop
        if ($errorLines) {
            Write-Host "Recent errors found:" -ForegroundColor Yellow
            $errorLines | ForEach-Object { Write-Host $_ -ForegroundColor Red }
        } else {
            Write-Host "✅ No recent errors in nginx log" -ForegroundColor Green
        }
    } catch {
        Write-Host "❌ Could not read nginx error log: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "❌ nginx error log not found at $errorLogPath" -ForegroundColor Red
}

# Test 8: Windows Firewall check
Write-Host ""
Write-Host "8. Checking Windows Firewall for port 8443..." -ForegroundColor Green
try {
    $firewallRule = Get-NetFirewallRule | Where-Object {$_.DisplayName -like "*8443*" -or $_.DisplayName -like "*nginx*"} | Select-Object -First 1
    if ($firewallRule) {
        Write-Host "✅ Found firewall rule: $($firewallRule.DisplayName)" -ForegroundColor Green
    } else {
        Write-Host "❌ No firewall rule found for port 8443" -ForegroundColor Red
        Write-Host "💡 You may need to add a firewall rule:" -ForegroundColor Yellow
        Write-Host "New-NetFirewallRule -DisplayName 'nginx HTTPS' -Direction Inbound -Protocol TCP -LocalPort 8443 -Action Allow" -ForegroundColor Cyan
    }
} catch {
    Write-Host "❌ Could not check firewall rules: $($_.Exception.Message)" -ForegroundColor Red
}

# Summary and recommendations
Write-Host ""
Write-Host "🎯 SUMMARY AND RECOMMENDATIONS:" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan

if ($missingFiles.Count -gt 0) {
    Write-Host "🔧 Frontend deployment issue detected:" -ForegroundColor Yellow
    Write-Host "   Run: npm run build:staging && copy dist\* C:\inetpub\wwwroot\ExcelAddin\dist\" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "📋 Quick fixes to try:" -ForegroundColor Green
Write-Host "1. Restart nginx: Stop-Service nginx; Start-Service nginx" -ForegroundColor White
Write-Host "2. Check backend: Get-Service ExcelAddinBackend" -ForegroundColor White
Write-Host "3. Test config: C:\nginx\nginx.exe -t" -ForegroundColor White
Write-Host "4. View logs: Get-Content C:\nginx\logs\error.log -Tail 20" -ForegroundColor White

Write-Host ""
Write-Host "🔍 Diagnostics complete!" -ForegroundColor Cyan