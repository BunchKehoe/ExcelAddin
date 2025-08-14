# IIS Proxy Deployment Script for ExcelAddin
# Deploys and configures IIS reverse proxy to forward requests to frontend (port 3000) and backend (port 5000)
# Automatically removes ALL existing ExcelAddin and ExcelAddin-Proxy instances from IIS before deployment

param(
    [string]$SiteName = "ExcelAddin-Proxy",
    [string]$AppPoolName = "ExcelAddin-Proxy",
    [int]$Port = 9443,
    [string]$FrontendUrl = "http://localhost:3000",
    [string]$BackendUrl = "http://localhost:5000",
    [string]$ServerFQDN = "server-vs81t.intranet.local",
    [switch]$Force,
    [switch]$Debug
)

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Green
Write-Host "  IIS Proxy Deployment for ExcelAddin" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

$startTime = Get-Date

# Check if running as administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script must be run as Administrator"
}

Write-Host "Configuration:" -ForegroundColor Cyan
Write-Host "  Site Name: $SiteName"
Write-Host "  Application Pool: $AppPoolName"
Write-Host "  Port: $Port"
Write-Host "  Frontend URL: $FrontendUrl"
Write-Host "  Backend URL: $BackendUrl"
Write-Host "  Server FQDN: $ServerFQDN"
Write-Host ""

try {
    # Check if IIS is installed
    Write-Host "Checking IIS installation..." -ForegroundColor Yellow
    $iisFeatures = Get-WindowsOptionalFeature -Online -FeatureName "IIS-*" | Where-Object { $_.State -eq "Enabled" }
    if (-not $iisFeatures) {
        Write-Error "IIS is not installed. Please install IIS with the following features: Web-Server, Web-Common-Http, Web-Mgmt-Tools"
    }
    Write-Host "  ✅ IIS is installed" -ForegroundColor Green

    # Import WebAdministration module
    Write-Host "Loading IIS PowerShell module..." -ForegroundColor Yellow
    Import-Module WebAdministration -ErrorAction Stop
    Write-Host "  ✅ WebAdministration module loaded" -ForegroundColor Green

    # Check for URL Rewrite module
    Write-Host "Checking URL Rewrite module..." -ForegroundColor Yellow
    $urlRewriteModule = Get-WebConfigurationProperty -Filter "system.webServer/modules/add[@name='RewriteModule']" -Name "name" -ErrorAction SilentlyContinue
    if (-not $urlRewriteModule) {
        Write-Warning "  ⚠️  URL Rewrite module is not installed"
        Write-Warning "     Download from: https://www.iis.net/downloads/microsoft/url-rewrite"
        Write-Warning "     Proxy will be configured but rewrite rules may not work properly"
    } else {
        Write-Host "  ✅ URL Rewrite module is available" -ForegroundColor Green
    }

    # Remove ALL existing ExcelAddin sites and app pools
    Write-Host "Cleaning up any existing ExcelAddin instances in IIS..." -ForegroundColor Yellow
    
    # Find and remove existing websites
    $existingSites = Get-Website | Where-Object { $_.Name -like "*ExcelAddin*" }
    if ($existingSites) {
        Write-Host "  Found $($existingSites.Count) existing ExcelAddin website(s) to remove:" -ForegroundColor Yellow
        foreach ($site in $existingSites) {
            Write-Host "    • $($site.Name) (State: $($site.State))" -ForegroundColor Gray
            try {
                if ($site.State -eq "Started") {
                    Stop-Website -Name $site.Name -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 2
                }
                Remove-Website -Name $site.Name -ErrorAction Stop
                Write-Host "      ✅ Removed website: $($site.Name)" -ForegroundColor Green
            } catch {
                Write-Warning "      ⚠️  Failed to remove website '$($site.Name)': $($_.Exception.Message)"
            }
        }
    } else {
        Write-Host "  ✅ No existing ExcelAddin websites found" -ForegroundColor Green
    }
    
    # Find and remove existing application pools
    $existingPools = Get-IISAppPool | Where-Object { $_.Name -like "*ExcelAddin*" }
    if ($existingPools) {
        Write-Host "  Found $($existingPools.Count) existing ExcelAddin application pool(s) to remove:" -ForegroundColor Yellow
        foreach ($pool in $existingPools) {
            Write-Host "    • $($pool.Name) (State: $($pool.State))" -ForegroundColor Gray
            try {
                if ($pool.State -eq "Started") {
                    Stop-WebAppPool -Name $pool.Name -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 2
                }
                Remove-WebAppPool -Name $pool.Name -ErrorAction Stop
                Write-Host "      ✅ Removed application pool: $($pool.Name)" -ForegroundColor Green
            } catch {
                Write-Warning "      ⚠️  Failed to remove application pool '$($pool.Name)': $($_.Exception.Message)"
            }
        }
    } else {
        Write-Host "  ✅ No existing ExcelAddin application pools found" -ForegroundColor Green
    }

    Write-Host "  ✅ ExcelAddin cleanup completed" -ForegroundColor Green
    Write-Host ""

    # Verify cleanup was successful (legacy check - should be covered by cleanup above)
    $existingSite = Get-Website -Name $SiteName -ErrorAction SilentlyContinue
    $existingPool = Get-IISAppPool -Name $AppPoolName -ErrorAction SilentlyContinue
    
    if ($existingSite -or $existingPool) {
        if (-not $Force) {
            Write-Error "Site '$SiteName' or Application Pool '$AppPoolName' still exists after cleanup. Use -Force to override any remaining conflicts."
        } else {
            # Force cleanup of any remaining instances
            if ($existingSite) {
                Write-Host "Force removing remaining site '$SiteName'..." -ForegroundColor Yellow
                Stop-Website -Name $SiteName -ErrorAction SilentlyContinue
                Remove-Website -Name $SiteName -ErrorAction SilentlyContinue
            }
            if ($existingPool) {
                Write-Host "Force removing remaining application pool '$AppPoolName'..." -ForegroundColor Yellow
                Stop-WebAppPool -Name $AppPoolName -ErrorAction SilentlyContinue
                Remove-WebAppPool -Name $AppPoolName -ErrorAction SilentlyContinue
            }
        }
    }

    # Create Application Pool
    Write-Host "Creating application pool '$AppPoolName'..." -ForegroundColor Yellow
    New-WebAppPool -Name $AppPoolName
    Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name "processModel.identityType" -Value "ApplicationPoolIdentity"
    Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name "enable32BitAppOnWin64" -Value $false
    Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name "managedRuntimeVersion" -Value ""  # No managed code needed for proxy
    Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name "recycling.periodicRestart.time" -Value "00:00:00"  # Disable periodic restart
    Write-Host "  ✅ Application pool created and configured" -ForegroundColor Green

    # Create physical directory
    Write-Host "Creating site directory..." -ForegroundColor Yellow
    $sitePath = "C:\inetpub\wwwroot\$SiteName"
    if (-not (Test-Path $sitePath)) {
        New-Item -ItemType Directory -Path $sitePath -Force | Out-Null
    }
    
    # Create status page
    $statusPageContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Prime Capital Excel Add-in Proxy</title>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            margin: 40px; 
            background-color: #f8f9fa;
            color: #333;
        }
        .container {
            max-width: 800px;
            margin: 0 auto;
            background: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .header {
            text-align: center;
            margin-bottom: 30px;
            padding-bottom: 20px;
            border-bottom: 2px solid #0078d4;
        }
        .status { 
            background: #d4edda; 
            padding: 20px; 
            border-radius: 8px; 
            border: 1px solid #c3e6cb; 
            margin: 20px 0;
        }
        .status h3 {
            color: #155724;
            margin-top: 0;
        }
        .endpoints {
            background: #f8f9fa;
            padding: 20px;
            border-radius: 8px;
            border: 1px solid #dee2e6;
        }
        .endpoints ul {
            list-style: none;
            padding: 0;
        }
        .endpoints li {
            padding: 8px 0;
            border-bottom: 1px solid #eee;
        }
        .endpoints a {
            color: #0078d4;
            text-decoration: none;
            font-weight: 500;
        }
        .endpoints a:hover {
            text-decoration: underline;
        }
        .info {
            background: #cce5ff;
            padding: 15px;
            border-radius: 6px;
            border: 1px solid #99ccff;
            margin: 15px 0;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Prime Capital Excel Add-in</h1>
            <h2>IIS Reverse Proxy Server</h2>
        </div>
        
        <div class="status">
            <h3>✅ Proxy Server Active</h3>
            <p>This server is acting as a reverse proxy for the Prime Capital Excel Add-in services.</p>
        </div>

        <div class="info">
            <h4>📋 Service Architecture</h4>
            <ul>
                <li><strong>Frontend Service:</strong> $FrontendUrl (Excel Add-in UI)</li>
                <li><strong>Backend Service:</strong> $BackendUrl (API and data processing)</li>
                <li><strong>Proxy Port:</strong> $Port (This server)</li>
            </ul>
        </div>

        <div class="endpoints">
            <h3>🔗 Available Endpoints</h3>
            <ul>
                <li><a href="/excellence/taskpane.html">Excel Taskpane Interface</a> → Frontend</li>
                <li><a href="/excellence/commands.html">Excel Commands Interface</a> → Frontend</li>
                <li><a href="/manifest.xml">Excel Manifest (Local)</a> → Frontend</li>
                <li><a href="/manifest-staging.xml">Excel Manifest (Staging)</a> → Frontend</li>
                <li><a href="/manifest-prod.xml">Excel Manifest (Production)</a> → Frontend</li>
                <li><a href="/functions.json">Custom Functions</a> → Frontend</li>
                <li><a href="/api/health">API Health Check</a> → Backend</li>
                <li><a href="/api/status">API Status</a> → Backend</li>
            </ul>
        </div>
    </div>
    
    <script>
        // Simple health check display
        document.addEventListener('DOMContentLoaded', function() {
            console.log('Prime Capital Excel Add-in Proxy Server');
            console.log('Deployed at: ' + new Date().toISOString());
        });
    </script>
</body>
</html>
"@

    $statusPageContent | Out-File -FilePath (Join-Path $sitePath "default.htm") -Encoding UTF8
    Write-Host "  ✅ Site directory and status page created" -ForegroundColor Green

    # Create Website
    Write-Host "Creating IIS website '$SiteName'..." -ForegroundColor Yellow
    New-Website -Name $SiteName -Port $Port -PhysicalPath $sitePath -ApplicationPool $AppPoolName
    Write-Host "  ✅ IIS website created" -ForegroundColor Green

    # Configure URL rewrite rules
    if ($urlRewriteModule) {
        Write-Host "Configuring URL rewrite rules..." -ForegroundColor Yellow
        
        $webConfigContent = @"
<?xml version="1.0" encoding="UTF-8"?>
<configuration>
  <system.webServer>
    <rewrite>
      <rules>
        <!-- Excel Add-in Frontend Routes -->
        <rule name="Excel Frontend Assets" stopProcessing="true">
          <match url="^excellence/(.*)" />
          <conditions>
            <add input="{REQUEST_METHOD}" pattern="OPTIONS" negate="true" />
          </conditions>
          <action type="Rewrite" url="$FrontendUrl/excellence/{R:1}" />
          <serverVariables>
            <set name="HTTP_X_FORWARDED_HOST" value="{HTTP_HOST}" />
            <set name="HTTP_X_FORWARDED_PROTO" value="{HTTPS}" />
          </serverVariables>
        </rule>
        
        <!-- Root Excel Files (manifests, functions.json) -->
        <rule name="Excel Root Files" stopProcessing="true">
          <match url="^(manifest.*\.xml|functions\.json)$" />
          <conditions>
            <add input="{REQUEST_METHOD}" pattern="OPTIONS" negate="true" />
          </conditions>
          <action type="Rewrite" url="$FrontendUrl/{R:1}" />
          <serverVariables>
            <set name="HTTP_X_FORWARDED_HOST" value="{HTTP_HOST}" />
            <set name="HTTP_X_FORWARDED_PROTO" value="{HTTPS}" />
          </serverVariables>
        </rule>
        
        <!-- Backend API Routes -->
        <rule name="Backend API" stopProcessing="true">
          <match url="^api/(.*)" />
          <conditions>
            <add input="{REQUEST_METHOD}" pattern="OPTIONS" negate="true" />
          </conditions>
          <action type="Rewrite" url="$BackendUrl/api/{R:1}" />
          <serverVariables>
            <set name="HTTP_X_FORWARDED_HOST" value="{HTTP_HOST}" />
            <set name="HTTP_X_FORWARDED_PROTO" value="{HTTPS}" />
          </serverVariables>
        </rule>
        
        <!-- Handle CORS preflight OPTIONS requests -->
        <rule name="CORS Preflight" stopProcessing="true">
          <match url=".*" />
          <conditions>
            <add input="{REQUEST_METHOD}" pattern="OPTIONS" />
          </conditions>
          <action type="CustomResponse" statusCode="200" statusReason="OK" statusDescription="OK" />
        </rule>
      </rules>
      
      <!-- Outbound rules to modify response headers -->
      <outboundRules>
        <rule name="Add CORS Headers" preCondition="IsHTML">
          <match filterByTags="None" pattern=".*" />
          <action type="None" />
          <serverVariables>
            <set name="RESPONSE_Access_Control_Allow_Origin" value="*" />
            <set name="RESPONSE_Access_Control_Allow_Methods" value="GET, POST, PUT, DELETE, PATCH, OPTIONS" />
            <set name="RESPONSE_Access_Control_Allow_Headers" value="Content-Type, Authorization, X-Requested-With" />
          </serverVariables>
        </rule>
        <preConditions>
          <preCondition name="IsHTML">
            <add input="{RESPONSE_CONTENT_TYPE}" pattern="^text/html" />
          </preCondition>
        </preConditions>
      </outboundRules>
    </rewrite>
    
    <!-- CORS headers for Excel Add-in compatibility -->
    <httpProtocol>
      <customHeaders>
        <add name="Access-Control-Allow-Origin" value="*" />
        <add name="Access-Control-Allow-Methods" value="GET, POST, PUT, DELETE, PATCH, OPTIONS" />
        <add name="Access-Control-Allow-Headers" value="Content-Type, Authorization, X-Requested-With, Accept" />
        <add name="Access-Control-Max-Age" value="86400" />
      </customHeaders>
    </httpProtocol>
    
    <!-- Default documents -->
    <defaultDocument>
      <files>
        <clear />
        <add value="default.htm" />
      </files>
    </defaultDocument>
    
    <!-- Static content compression -->
    <urlCompression doDynamicCompression="true" doStaticCompression="true" />
    
    <!-- Security headers -->
    <httpProtocol>
      <customHeaders>
        <add name="X-Content-Type-Options" value="nosniff" />
        <add name="X-Frame-Options" value="SAMEORIGIN" />
        <add name="X-XSS-Protection" value="1; mode=block" />
      </customHeaders>
    </httpProtocol>
  </system.webServer>
</configuration>
"@
        
        $webConfigContent | Out-File -FilePath (Join-Path $sitePath "web.config") -Encoding UTF8
        Write-Host "  ✅ URL rewrite rules configured" -ForegroundColor Green
    }

    # Configure HTTPS binding if port suggests SSL
    if ($Port -eq 443 -or $Port -eq 9443) {
        Write-Host "Configuring HTTPS binding..." -ForegroundColor Yellow
        
        # Look for available SSL certificates
        $certs = Get-ChildItem Cert:\LocalMachine\My | Where-Object { 
            $_.Subject -like "*$ServerFQDN*" -or 
            $_.Subject -like "*localhost*" -or
            $_.DnsNameList -contains $ServerFQDN
        } | Sort-Object NotAfter -Descending

        if ($certs.Count -gt 0) {
            $cert = $certs[0]
            Write-Host "  Using SSL certificate: $($cert.Subject)" -ForegroundColor Green
            Write-Host "  Certificate expires: $($cert.NotAfter)" -ForegroundColor Gray
            
            # Remove existing binding if it exists
            $existingBinding = Get-WebBinding -Name $SiteName -Protocol "https" -ErrorAction SilentlyContinue
            if ($existingBinding) {
                Remove-WebBinding -Name $SiteName -Protocol "https" -Port $Port
            }
            
            # Create HTTPS binding
            New-WebBinding -Name $SiteName -Protocol "https" -Port $Port -SslFlags 1 -Thumbprint $cert.Thumbprint
            Write-Host "  ✅ HTTPS binding configured" -ForegroundColor Green
        } else {
            Write-Warning "  ⚠️  No suitable SSL certificate found for $ServerFQDN"
            Write-Warning "     HTTPS binding not configured - proxy will only work over HTTP"
            Write-Host "  To add SSL certificate:" -ForegroundColor Yellow
            Write-Host "    1. Import certificate to Local Machine Personal store" -ForegroundColor Yellow
            Write-Host "    2. Run: New-WebBinding -Name '$SiteName' -Protocol 'https' -Port $Port -SslFlags 1 -Thumbprint <thumbprint>" -ForegroundColor Yellow
        }
    }

    # Start the website and app pool
    Write-Host "Starting application pool and website..." -ForegroundColor Yellow
    Start-WebAppPool -Name $AppPoolName
    Start-Website -Name $SiteName
    
    # Wait a moment for services to start
    Start-Sleep -Seconds 3
    
    # Verify site is running
    $site = Get-Website -Name $SiteName
    $pool = Get-WebAppPool -Name $AppPoolName
    
    if ($site.State -eq "Started" -and $pool.State -eq "Started") {
        Write-Host "  ✅ Website and application pool are running" -ForegroundColor Green
    } else {
        Write-Warning "  ⚠️  Site State: $($site.State), Pool State: $($pool.State)"
    }

    # Test the proxy
    Write-Host "Testing IIS proxy..." -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    
    $testResults = @()
    
    # Test main site
    try {
        $protocol = if ($Port -eq 443 -or $Port -eq 9443) { "https" } else { "http" }
        $testUrl = "${protocol}://localhost:${Port}"
        $response = Invoke-WebRequest -Uri $testUrl -TimeoutSec 10 -UseBasicParsing -ErrorAction SilentlyContinue
        if ($response.StatusCode -eq 200) {
            $testResults += "✅ Main site responding (HTTP $($response.StatusCode))"
        } else {
            $testResults += "⚠️  Main site returned HTTP $($response.StatusCode)"
        }
    } catch {
        $testResults += "❌ Main site test failed: $($_.Exception.Message)"
    }
    
    # Test if backend and frontend services are running (for proxy functionality)
    try {
        $frontendTest = Invoke-WebRequest -Uri "$FrontendUrl/health" -TimeoutSec 5 -UseBasicParsing -ErrorAction SilentlyContinue
        $testResults += "✅ Frontend service responding (HTTP $($frontendTest.StatusCode))"
    } catch {
        $testResults += "⚠️  Frontend service not responding - proxy forwarding may fail"
    }
    
    try {
        $backendTest = Invoke-WebRequest -Uri "$BackendUrl/api/health" -TimeoutSec 5 -UseBasicParsing -ErrorAction SilentlyContinue
        $testResults += "✅ Backend service responding (HTTP $($backendTest.StatusCode))"
    } catch {
        $testResults += "⚠️  Backend service not responding - API proxy forwarding may fail"
    }

    $endTime = Get-Date
    $duration = $endTime - $startTime

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "  IIS PROXY DEPLOYMENT COMPLETE!" -ForegroundColor Green
    Write-Host "  Duration: $($duration.ToString('mm\:ss'))" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""

    Write-Host "Deployment Summary:" -ForegroundColor Cyan
    Write-Host "  Site Name: $SiteName"
    Write-Host "  Application Pool: $AppPoolName"
    Write-Host "  Port: $Port"
    Write-Host "  Physical Path: $sitePath"
    Write-Host "  State: $($site.State)"
    Write-Host ""

    Write-Host "Test Results:" -ForegroundColor Cyan
    foreach ($result in $testResults) {
        Write-Host "  $result"
    }
    Write-Host ""

    Write-Host "Access URLs:" -ForegroundColor Cyan
    $protocol = if ($Port -eq 443 -or $Port -eq 9443) { "https" } else { "http" }
    Write-Host "  Proxy Status: ${protocol}://localhost:${Port}/"
    Write-Host "  External URL: ${protocol}://${ServerFQDN}:${Port}/"
    Write-Host ""
    Write-Host "  Excel Add-in Endpoints (through proxy):" -ForegroundColor Cyan
    Write-Host "    Taskpane: ${protocol}://${ServerFQDN}:${Port}/excellence/taskpane.html"
    Write-Host "    Commands: ${protocol}://${ServerFQDN}:${Port}/excellence/commands.html"
    Write-Host "    Manifest: ${protocol}://${ServerFQDN}:${Port}/manifest-staging.xml"
    Write-Host "    API Health: ${protocol}://${ServerFQDN}:${Port}/api/health"
    Write-Host ""

    Write-Host "Management Commands:" -ForegroundColor Cyan
    Write-Host "  Start Site: Start-Website -Name '$SiteName'"
    Write-Host "  Stop Site: Stop-Website -Name '$SiteName'"
    Write-Host "  Check Status: Get-Website -Name '$SiteName' | Select Name, State, PhysicalPath"
    Write-Host "  View Bindings: Get-WebBinding -Name '$SiteName'"
    Write-Host ""

    if (-not $urlRewriteModule) {
        Write-Host "NEXT STEPS:" -ForegroundColor Red
        Write-Host "  1. Install URL Rewrite module: https://www.iis.net/downloads/microsoft/url-rewrite" -ForegroundColor Yellow
        Write-Host "  2. Restart this deployment script to configure rewrite rules" -ForegroundColor Yellow
        Write-Host "  3. Ensure frontend and backend services are running for full functionality" -ForegroundColor Yellow
    }

    if ($Debug) {
        Write-Host ""
        Write-Host "DEBUG INFORMATION:" -ForegroundColor Magenta
        Write-Host "  IIS Site Object:" -ForegroundColor Magenta
        $site | Format-List | Out-Host
        Write-Host "  Application Pool Object:" -ForegroundColor Magenta  
        $pool | Format-List | Out-Host
    }

} catch {
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "  IIS PROXY DEPLOYMENT FAILED!" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Duration: $($duration.ToString('mm\:ss'))" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    
    Write-Host "Troubleshooting:" -ForegroundColor Yellow
    Write-Host "  1. Ensure you're running as Administrator" -ForegroundColor Yellow
    Write-Host "  2. Verify IIS is installed with required features" -ForegroundColor Yellow
    Write-Host "  3. Check if ports are available: netstat -an | findstr :$Port" -ForegroundColor Yellow
    Write-Host "  4. Install URL Rewrite module if needed" -ForegroundColor Yellow
    Write-Host "  5. Check Windows Event Logs for IIS errors" -ForegroundColor Yellow
    
    if ($Debug) {
        Write-Host ""
        Write-Host "DETAILED ERROR:" -ForegroundColor Red
        Write-Host $_.Exception.ToString() -ForegroundColor Red
        Write-Host ""
        Write-Host "STACK TRACE:" -ForegroundColor Red  
        Write-Host $_.ScriptStackTrace -ForegroundColor Red
    }
    
    exit 1
}