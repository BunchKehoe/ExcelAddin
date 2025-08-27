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

# Import WebAdministration module early for helper functions
Import-Module WebAdministration -ErrorAction Stop

# Helper function to safely remove IIS application pool
function Remove-IISAppPoolSafely {
    param([string]$PoolName)
    
    try {
        # Try to stop the app pool first (if it exists and is running)
        Stop-WebAppPool -Name $PoolName -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        
        # Try to remove the app pool
        Remove-WebAppPool -Name $PoolName -ErrorAction SilentlyContinue
        
        Write-Host "      ✅ Application pool '$PoolName' removed (if it existed)" -ForegroundColor Green
        return $true
    } catch {
        Write-Warning "      ⚠️  Error during application pool removal for '$PoolName': $($_.Exception.Message)"
        return $false
    }
}

# Helper function to safely remove IIS website
function Remove-IISWebsiteSafely {
    param([string]$SiteName)
    
    try {
        # Try to stop the website first (if it exists and is running)
        Stop-Website -Name $SiteName -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        
        # Try to remove the website
        Remove-Website -Name $SiteName -ErrorAction SilentlyContinue
        
        Write-Host "      ✅ Website '$SiteName' removed (if it existed)" -ForegroundColor Green
        return $true
    } catch {
        Write-Warning "      ⚠️  Error during website removal for '$SiteName': $($_.Exception.Message)"
        return $false
    }
}

# Helper function to clean up port conflicts
function Clear-PortConflicts {
    param([int]$Port)
    
    Write-Host "Checking for port conflicts on port $Port..." -ForegroundColor Yellow
    
    try {
        # Check for any IIS bindings on this port
        Write-Host "  Checking IIS bindings..." -ForegroundColor Gray
        
        # Get all sites and their bindings
        $allSites = @()
        try {
            # Try to enumerate sites using configuration approach since Get-Website may not work
            $sitesConfig = Get-WebConfigurationProperty -Filter "system.webServer/sites/site" -Name "*" 2>$null
            foreach ($siteConfig in $sitesConfig) {
                $allSites += $siteConfig.name
            }
        } catch {
            # If that fails, try to get a list of site names from IIS paths
            try {
                $sitePaths = Get-ChildItem "IIS:\Sites" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name
                $allSites = $sitePaths
            } catch {
                Write-Host "    Unable to enumerate sites, will attempt direct cleanup" -ForegroundColor Gray
            }
        }
        
        # Remove bindings on the target port from any found sites
        foreach ($siteName in $allSites) {
            try {
                $bindings = Get-WebBinding -Name $siteName -ErrorAction SilentlyContinue
                foreach ($binding in $bindings) {
                    if ($binding.bindingInformation -like "*:${Port}:*") {
                        Write-Host "    Removing conflicting binding from '$siteName': $($binding.bindingInformation)" -ForegroundColor Yellow
                        Remove-WebBinding -Name $siteName -Port $Port -Protocol $binding.protocol -ErrorAction SilentlyContinue
                    }
                }
            } catch {
                # Continue if we can't check this site
            }
        }
        
        # Also try to clean up any orphaned SSL certificate bindings
        Write-Host "  Checking SSL certificate bindings..." -ForegroundColor Gray
        try {
            $sslBindings = netsh http show sslcert | Select-String ":$Port"
            if ($sslBindings) {
                Write-Host "    Found SSL bindings on port $Port, attempting cleanup..." -ForegroundColor Yellow
                # Use netsh to remove SSL bindings
                & netsh http delete sslcert ipport=0.0.0.0:$Port 2>$null
                & netsh http delete sslcert ipport=[::]:$Port 2>$null
            }
        } catch {
            # SSL cleanup failed, continue
        }
        
        Write-Host "  ✅ Port conflict cleanup completed" -ForegroundColor Green
        
    } catch {
        Write-Warning "  ⚠️  Port conflict cleanup failed: $($_.Exception.Message)"
    }
}

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

    # Verify WebAdministration module is loaded (already imported at script start)
    Write-Host "Verifying IIS PowerShell module..." -ForegroundColor Yellow
    if (Get-Module -Name WebAdministration) {
        Write-Host "  ✅ WebAdministration module loaded" -ForegroundColor Green
    } else {
        Write-Error "WebAdministration module not available"
    }

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
    
    # Remove potential existing websites using try/catch approach
    $potentialSites = @("ExcelAddin", "ExcelAddin-Proxy", $SiteName)
    $removedSites = 0
    foreach ($siteName in $potentialSites) {
        try {
            $removed = Remove-IISWebsiteSafely -SiteName $siteName
            if ($removed) {
                $removedSites++
            }
        } catch {
            # Silently continue - site probably doesn't exist
        }
    }
    
    # Remove potential existing application pools using try/catch approach  
    $potentialPools = @("ExcelAddin", "ExcelAddin-Proxy", $AppPoolName)
    $removedPools = 0
    foreach ($poolName in $potentialPools) {
        try {
            $removed = Remove-IISAppPoolSafely -PoolName $poolName
            if ($removed) {
                $removedPools++
            }
        } catch {
            # Silently continue - pool probably doesn't exist
        }
    }
    
    if ($removedSites -gt 0 -or $removedPools -gt 0) {
        Write-Host "  ✅ Cleaned up $removedSites website(s) and $removedPools application pool(s)" -ForegroundColor Green
    } else {
        Write-Host "  ✅ No existing ExcelAddin instances found to remove" -ForegroundColor Green
    }

    Write-Host "  ✅ ExcelAddin cleanup completed" -ForegroundColor Green
    Write-Host ""

    # Clean up any port conflicts before proceeding
    Clear-PortConflicts -Port $Port

    # Create Application Pool
    Write-Host "Creating application pool '$AppPoolName'..." -ForegroundColor Yellow
    try {
        New-WebAppPool -Name $AppPoolName -ErrorAction Stop
        Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name "processModel.identityType" -Value "ApplicationPoolIdentity"
        Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name "enable32BitAppOnWin64" -Value $false
        Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name "managedRuntimeVersion" -Value ""  # No managed code needed for proxy
        Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name "recycling.periodicRestart.time" -Value "00:00:00"  # Disable periodic restart
        Write-Host "  ✅ Application pool created and configured" -ForegroundColor Green
    } catch {
        if ($_.Exception.Message -like "*already exists*") {
            Write-Host "  ✅ Application pool '$AppPoolName' already exists (continuing)" -ForegroundColor Green
            # Update existing pool settings
            Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name "processModel.identityType" -Value "ApplicationPoolIdentity" -ErrorAction SilentlyContinue
            Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name "enable32BitAppOnWin64" -Value $false -ErrorAction SilentlyContinue
            Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name "managedRuntimeVersion" -Value "" -ErrorAction SilentlyContinue
            Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name "recycling.periodicRestart.time" -Value "00:00:00" -ErrorAction SilentlyContinue
        } else {
            Write-Error "Failed to create application pool: $($_.Exception.Message)"
        }
    }

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
    try {
        New-Website -Name $SiteName -Port $Port -PhysicalPath $sitePath -ApplicationPool $AppPoolName -ErrorAction Stop
        Write-Host "  ✅ IIS website created" -ForegroundColor Green
    } catch {
        if ($_.Exception.Message -like "*already exists*") {
            Write-Host "  ✅ IIS website '$SiteName' already exists (continuing)" -ForegroundColor Green
            # Update existing website settings
            Set-ItemProperty -Path "IIS:\Sites\$SiteName" -Name "physicalPath" -Value $sitePath -ErrorAction SilentlyContinue
            Set-ItemProperty -Path "IIS:\Sites\$SiteName" -Name "applicationPool" -Value $AppPoolName -ErrorAction SilentlyContinue
        } else {
            Write-Error "Failed to create website: $($_.Exception.Message)"
        }
    }

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
        
        <!-- Backend API Routes (direct) -->
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
        
        <!-- Backend API Routes (via excellence path - for backward compatibility) -->
        <rule name="Backend API Excellence" stopProcessing="true">
          <match url="^excellence/backend/api/(.*)" />
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
                Remove-WebBinding -Name $SiteName -Protocol "https" -Port $Port -HostHeader $ServerFQDN -ErrorAction SilentlyContinue
            }
            
            # Create HTTPS binding with hostname for SNI support
            New-WebBinding -Name $SiteName -Protocol "https" -Port $Port -HostHeader $ServerFQDN -SslFlags 1
            
            # Then bind the SSL certificate to the binding
            try {
                $binding = Get-WebBinding -Name $SiteName -Protocol "https" -Port $Port -HostHeader $ServerFQDN
                $binding.AddSslCertificate($cert.Thumbprint, "my")
                Write-Host "  ✅ HTTPS binding configured with SSL certificate" -ForegroundColor Green
            } catch {
                Write-Warning "  ⚠️  Failed to bind SSL certificate: $($_.Exception.Message)"
                Write-Warning "     HTTPS binding created but SSL certificate not bound"
                Write-Host "  Manual certificate binding command:" -ForegroundColor Yellow
                Write-Host "    netsh http add sslcert ipport=0.0.0.0:$Port certhash=$($cert.Thumbprint) appid={$([System.Guid]::NewGuid().ToString())}" -ForegroundColor Yellow
            }
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
    try {
        Start-WebAppPool -Name $AppPoolName -ErrorAction Stop
        Write-Host "  ✅ Application pool started" -ForegroundColor Green
    } catch {
        Write-Warning "  ⚠️  Failed to start application pool: $($_.Exception.Message)"
    }
    
    try {
        Start-Website -Name $SiteName -ErrorAction Stop
        Write-Host "  ✅ Website started" -ForegroundColor Green
    } catch {
        Write-Warning "  ⚠️  Failed to start website: $($_.Exception.Message)"
        
        # If website start fails, check for port conflicts
        Write-Host "Checking for port conflicts..." -ForegroundColor Yellow
        try {
            $portCheck = netstat -an | findstr ":$Port "
            if ($portCheck) {
                Write-Warning "  ⚠️  Port $Port appears to be in use:"
                Write-Host "    $portCheck" -ForegroundColor Gray
            }
        } catch {
            # netstat command failed, continue
        }
        
        # Try to identify conflicting bindings
        Write-Host "Checking IIS bindings for port conflicts..." -ForegroundColor Yellow
        try {
            $allBindings = Get-WebConfigurationProperty -Filter "system.webServer/sites/site/bindings/binding" -Name "*" 2>$null
            $conflictingBindings = $allBindings | Where-Object { $_.bindingInformation -like "*:${Port}:*" }
            if ($conflictingBindings) {
                Write-Warning "  ⚠️  Found existing IIS bindings on port ${Port}:"
                foreach ($binding in $conflictingBindings) {
                    Write-Host "    $($binding.bindingInformation)" -ForegroundColor Gray
                }
            }
        } catch {
            # Binding check failed, continue
        }
    }
    
    # Wait a moment for services to start
    Start-Sleep -Seconds 3

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
    Write-Host "  Status: Deployment completed (check manually if needed)"
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
    Write-Host "  Remove Site: Remove-Website -Name '$SiteName'"
    Write-Host "  Remove App Pool: Remove-WebAppPool -Name '$AppPoolName'"
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
        Write-Host "  Deployment completed with error-resistant approach" -ForegroundColor Magenta
        Write-Host "  Site Name: $SiteName" -ForegroundColor Magenta
        Write-Host "  App Pool Name: $AppPoolName" -ForegroundColor Magenta
        Write-Host "  Port: $Port" -ForegroundColor Magenta
        Write-Host "  Physical Path: $sitePath" -ForegroundColor Magenta
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