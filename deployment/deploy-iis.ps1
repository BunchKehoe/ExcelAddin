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
function Test-SiteConfiguration {
    param(
        [string]$SiteName,
        [string]$Port,
        [string]$SitePath
    )
    
    Write-Host "Diagnosing site configuration..." -ForegroundColor Yellow
    
    # Check if web.config exists and is valid
    $webConfigPath = Join-Path $SitePath "web.config"
    if (Test-Path $webConfigPath) {
        Write-Host "  ✅ web.config file exists" -ForegroundColor Green
        try {
            [xml]$webConfig = Get-Content $webConfigPath
            Write-Host "  ✅ web.config XML is valid" -ForegroundColor Green
            
            # Check for duplicate httpProtocol sections
            $httpProtocolNodes = $webConfig.configuration.'system.webServer'.ChildNodes | Where-Object { $_.Name -eq 'httpProtocol' }
            if ($httpProtocolNodes.Count -eq 1) {
                Write-Host "  ✅ Single httpProtocol section found" -ForegroundColor Green
            } elseif ($httpProtocolNodes.Count -gt 1) {
                Write-Warning "  ⚠️  Multiple httpProtocol sections found: $($httpProtocolNodes.Count)"
            }
        } catch {
            Write-Warning "  ⚠️  web.config XML validation failed: $($_.Exception.Message)"
        }
    } else {
        Write-Warning "  ⚠️  web.config file not found"
    }
    
    # Check IIS site status
    try {
        $site = Get-WebSite -Name $SiteName -ErrorAction SilentlyContinue
        if ($site) {
            Write-Host "  ✅ IIS site exists - State: $($site.State)" -ForegroundColor Green
            Write-Host "  📁 Physical Path: $($site.PhysicalPath)" -ForegroundColor Gray
            Write-Host "  🌐 Bindings: $($site.Bindings.Collection.bindingInformation -join ', ')" -ForegroundColor Gray
        } else {
            Write-Warning "  ⚠️  IIS site not found"
        }
    } catch {
        Write-Warning "  ⚠️  Cannot check IIS site status: $($_.Exception.Message)"
    }
    
    # Check application pool status
    try {
        $appPool = Get-WebAppPool -Name $SiteName -ErrorAction SilentlyContinue
        if ($appPool) {
            Write-Host "  ✅ Application pool exists - State: $($appPool.State)" -ForegroundColor Green
        } else {
            Write-Warning "  ⚠️  Application pool not found"
        }
    } catch {
        Write-Warning "  ⚠️  Cannot check application pool status: $($_.Exception.Message)"
    }
}

function Clear-PortConflicts {
    param([int]$Port)
    
    Write-Host "Checking for port conflicts on port $Port..." -ForegroundColor Yellow
    
    try {
        # First, check if anything is actually listening on the port
        Write-Host "  Checking for processes listening on port $Port..." -ForegroundColor Gray
        $listeningProcesses = netstat -ano | Where-Object { $_ -match ":${Port}\s" }
        if ($listeningProcesses) {
            Write-Host "    ⚠️  Found processes listening on port ${Port}:" -ForegroundColor Yellow
            foreach ($proc in $listeningProcesses) {
                Write-Host "      $proc" -ForegroundColor Gray
            }
        } else {
            Write-Host "    ✅ No processes found listening on port $Port" -ForegroundColor Green
        }
        
        # Check for any IIS bindings on this port using multiple approaches
        Write-Host "  Checking IIS bindings..." -ForegroundColor Gray
        
        $conflictingSites = @()
        
        # Try to get all sites using configuration approach
        try {
            $webConfig = Get-WebConfiguration -Filter "system.webServer/sites/site" -ErrorAction SilentlyContinue 2>$null
            if ($webConfig) {
                foreach ($site in $webConfig) {
                    try {
                        $siteName = $site.name
                        if ($siteName) {
                            # Check bindings for this site
                            $bindings = Get-WebBinding -Name $siteName -ErrorAction SilentlyContinue 2>$null
                            foreach ($binding in $bindings) {
                                if ($binding.bindingInformation -like "*:${Port}:*") {
                                    $conflictingSites += @{
                                        SiteName = $siteName
                                        Binding = $binding.bindingInformation
                                        Protocol = $binding.protocol
                                    }
                                    Write-Host "      ⚠️  Found conflicting binding: $siteName ($($binding.bindingInformation))" -ForegroundColor Yellow
                                }
                            }
                        }
                    } catch {
                        # Skip sites we can't check
                        continue
                    }
                }
            }
        } catch {
            Write-Host "    ⚠️  Could not enumerate sites using configuration approach" -ForegroundColor Yellow
        }
        
        # Remove conflicting bindings
        if ($conflictingSites.Count -gt 0) {
            Write-Host "    Removing conflicting IIS bindings..." -ForegroundColor Yellow
            foreach ($conflict in $conflictingSites) {
                try {
                    Remove-WebBinding -Name $conflict.SiteName -Port $Port -Protocol $conflict.Protocol -ErrorAction SilentlyContinue
                    Write-Host "      ✅ Removed binding from $($conflict.SiteName)" -ForegroundColor Green
                } catch {
                    Write-Host "      ⚠️  Could not remove binding from $($conflict.SiteName): $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
        } else {
            Write-Host "    ✅ No conflicting IIS bindings found" -ForegroundColor Green
        }
        
        # Clean up SSL certificate bindings for the port
        Write-Host "  Cleaning up SSL certificate bindings..." -ForegroundColor Gray
        try {
            $sslBindings = netsh http show sslcert | Where-Object { $_ -match "0\.0\.0\.0:$Port" -or $_ -match "\[::\]:$Port" }
            if ($sslBindings) {
                Write-Host "    Removing SSL bindings for port $Port..." -ForegroundColor Yellow
                & netsh http delete sslcert ipport=0.0.0.0:$Port 2>$null
                & netsh http delete sslcert ipport=[::]:$Port 2>$null
                Write-Host "    ✅ SSL bindings cleaned up" -ForegroundColor Green
            } else {
                Write-Host "    ✅ No SSL bindings to clean up" -ForegroundColor Green
            }
        } catch {
            Write-Host "    ⚠️  Could not clean SSL bindings: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        
        Write-Host "  ✅ Port conflict cleanup completed" -ForegroundColor Green
        
    } catch {
        Write-Warning "  ⚠️  Port conflict cleanup failed: $($_.Exception.Message)"
    }
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
    
    # Create status page (using safer HTML generation to avoid PowerShell parsing issues)
    $statusPageContent = @'
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
            <h3>&#x2705; Proxy Server Active</h3>
            <p>This server is acting as a reverse proxy for the Prime Capital Excel Add-in services.</p>
        </div>

        <div class="info">
            <h4>&#x1F4CB; Service Architecture</h4>
            <ul>
                <li><strong>Frontend Service:</strong> FRONTEND_URL_PLACEHOLDER (Excel Add-in UI)</li>
                <li><strong>Backend Service:</strong> BACKEND_URL_PLACEHOLDER (API and data processing)</li>
                <li><strong>Proxy Port:</strong> PORT_PLACEHOLDER (This server)</li>
            </ul>
        </div>

        <div class="endpoints">
            <h3>&#x1F517; Available Endpoints</h3>
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
'@

    # Replace placeholders with actual values to avoid PowerShell variable parsing issues
    $statusPageContent = $statusPageContent -replace "FRONTEND_URL_PLACEHOLDER", $FrontendUrl
    $statusPageContent = $statusPageContent -replace "BACKEND_URL_PLACEHOLDER", $BackendUrl
    $statusPageContent = $statusPageContent -replace "PORT_PLACEHOLDER", $Port

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
    
    <!-- Combined HTTP headers for CORS and security -->
    <httpProtocol>
      <customHeaders>
        <!-- CORS headers for Excel Add-in compatibility -->
        <add name="Access-Control-Allow-Origin" value="*" />
        <add name="Access-Control-Allow-Methods" value="GET, POST, PUT, DELETE, PATCH, OPTIONS" />
        <add name="Access-Control-Allow-Headers" value="Content-Type, Authorization, X-Requested-With, Accept" />
        <add name="Access-Control-Max-Age" value="86400" />
        <!-- Security headers -->
        <add name="X-Content-Type-Options" value="nosniff" />
        <add name="X-Frame-Options" value="SAMEORIGIN" />
        <add name="X-XSS-Protection" value="1; mode=block" />
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
  </system.webServer>
</configuration>
"@
        
        # Validate the web.config XML before writing
        try {
            [xml]$xmlValidation = $webConfigContent
            Write-Host "  ✅ Web.config XML validation passed" -ForegroundColor Green
        } catch {
            Write-Error "Web.config XML validation failed: $($_.Exception.Message)"
            throw "Invalid web.config generated"
        }
        
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

    # Start the website and app pool with improved error handling
    Write-Host "Starting application pool and website..." -ForegroundColor Yellow
    
    # Start application pool first
    try {
        Start-WebAppPool -Name $AppPoolName -ErrorAction Stop
        Write-Host "  ✅ Application pool started" -ForegroundColor Green
    } catch {
        Write-Warning "  ⚠️  Failed to start application pool: $($_.Exception.Message)"
        
        # Try to get more diagnostic information
        try {
            $appPoolState = (Get-WebAppPool -Name $AppPoolName).State
            Write-Host "    Current app pool state: $appPoolState" -ForegroundColor Gray
        } catch {
            Write-Host "    Could not determine app pool state" -ForegroundColor Gray
        }
    }
    
    # Wait a moment for app pool to fully initialize
    Start-Sleep -Seconds 2
    
    # Start website with retry logic
    $websiteStarted = $false
    $retryCount = 0
    $maxRetries = 3
    
    while (-not $websiteStarted -and $retryCount -lt $maxRetries) {
        try {
            Start-Website -Name $SiteName -ErrorAction Stop
            Write-Host "  ✅ Website started successfully" -ForegroundColor Green
            $websiteStarted = $true
            
            # Run diagnostic to confirm configuration
            Test-SiteConfiguration -SiteName $SiteName -Port $Port -SitePath $sitePath
            
        } catch {
            $retryCount++
            Write-Warning "  ⚠️  Website start attempt $retryCount failed: $($_.Exception.Message)"
            
            if ($retryCount -lt $maxRetries) {
                Write-Host "    Retrying in 3 seconds..." -ForegroundColor Yellow
                Start-Sleep -Seconds 3
                
                # Clear port conflicts before retry
                Clear-PortConflicts -Port $Port
            } else {
                Write-Host "    Maximum retries reached. Running diagnostics..." -ForegroundColor Yellow
                
                # Run comprehensive diagnostics
                Test-SiteConfiguration -SiteName $SiteName -Port $Port -SitePath $sitePath
                
                # Enhanced port conflict detection
                Write-Host "  Performing enhanced port conflict analysis..." -ForegroundColor Yellow
                try {
                    $portCheck = netstat -ano | Where-Object { $_ -match ":$Port\s" }
                    if ($portCheck) {
                        Write-Warning "    ⚠️  Port $Port is in use by other processes:"
                        foreach ($proc in $portCheck) {
                            Write-Host "      $proc" -ForegroundColor Gray
                        }
                    } else {
                        Write-Host "    ✅ Port $Port appears to be available" -ForegroundColor Green
                    }
                } catch {
                    Write-Host "    Could not check port usage" -ForegroundColor Gray
                }
                
                # Check for IIS binding conflicts
                try {
                    $allBindings = Get-WebConfigurationProperty -Filter "system.webServer/sites/site/bindings/binding" -Name "*" -ErrorAction SilentlyContinue 2>$null
                    $conflictingBindings = $allBindings | Where-Object { $_.bindingInformation -like "*:${Port}:*" }
                    if ($conflictingBindings) {
                        Write-Warning "    ⚠️  Found existing IIS bindings on port ${Port}:"
                        foreach ($binding in $conflictingBindings) {
                            Write-Host "      $($binding.bindingInformation)" -ForegroundColor Gray
                        }
                    } else {
                        Write-Host "    ✅ No conflicting IIS bindings found" -ForegroundColor Green
                    }
                } catch {
                    Write-Host "    Could not check IIS bindings" -ForegroundColor Gray
                }
            }
        }
    }
    
    # Wait a moment for services to start, but don't wait too long to avoid hanging
    Write-Host "Waiting for services to initialize..." -ForegroundColor Yellow
    Start-Sleep -Seconds 3

    # Test the proxy with improved error handling
    Write-Host "Testing IIS proxy..." -ForegroundColor Yellow
    
    $testResults = @()
    
    # Test main site with shorter timeout and better error handling
    Write-Host "  Testing main site connectivity..." -ForegroundColor Gray
    try {
        $protocol = if ($Port -eq 443 -or $Port -eq 9443) { "https" } else { "http" }
        $testUrl = "${protocol}://localhost:${Port}"
        
        # Use shorter timeout and ignore SSL errors for testing
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "IIS-Deployment-Test")
        
        # Set timeout to 5 seconds to prevent hanging
        $response = $null
        $testJob = Start-Job -ScriptBlock {
            param($url)
            try {
                $request = [System.Net.WebRequest]::Create($url)
                $request.Timeout = 5000
                $request.Method = "GET"
                if ($url.StartsWith("https://")) {
                    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
                    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
                }
                $response = $request.GetResponse()
                return @{
                    StatusCode = [int]$response.StatusCode
                    Success = $true
                }
            } catch {
                return @{
                    StatusCode = 0
                    Success = $false
                    Error = $_.Exception.Message
                }
            }
        } -ArgumentList $testUrl
        
        # Wait up to 8 seconds for the test to complete
        $testCompleted = Wait-Job $testJob -Timeout 8
        if ($testCompleted) {
            $result = Receive-Job $testJob
            Remove-Job $testJob
            
            if ($result.Success -and $result.StatusCode -eq 200) {
                $testResults += "✅ Main site responding (HTTP $($result.StatusCode))"
            } elseif ($result.Success) {
                $testResults += "⚠️  Main site returned HTTP $($result.StatusCode)"
            } else {
                $testResults += "❌ Main site test failed: $($result.Error)"
            }
        } else {
            Remove-Job $testJob -Force
            $testResults += "❌ Main site test timed out (site may not be properly configured)"
        }
        
        $webClient.Dispose()
    } catch {
        $testResults += "❌ Main site test failed: $($_.Exception.Message)"
    }
    
    # Test backend and frontend services with shorter timeouts
    Write-Host "  Testing backend/frontend service connectivity..." -ForegroundColor Gray
    try {
        $frontendJob = Start-Job -ScriptBlock {
            param($url)
            try {
                $request = [System.Net.WebRequest]::Create("$url/health")
                $request.Timeout = 3000
                $response = $request.GetResponse()
                return @{ Success = $true; StatusCode = [int]$response.StatusCode }
            } catch {
                return @{ Success = $false; Error = $_.Exception.Message }
            }
        } -ArgumentList $FrontendUrl
        
        if (Wait-Job $frontendJob -Timeout 5) {
            $frontendResult = Receive-Job $frontendJob
            if ($frontendResult.Success) {
                $testResults += "✅ Frontend service responding (HTTP $($frontendResult.StatusCode))"
            } else {
                $testResults += "⚠️  Frontend service not responding - proxy forwarding may fail"
            }
        } else {
            $testResults += "⚠️  Frontend service test timed out - proxy forwarding may fail"
        }
        Remove-Job $frontendJob -Force
    } catch {
        $testResults += "⚠️  Frontend service not responding - proxy forwarding may fail"
    }
    
    try {
        $backendJob = Start-Job -ScriptBlock {
            param($url)
            try {
                $request = [System.Net.WebRequest]::Create("$url/api/health")
                $request.Timeout = 3000
                $response = $request.GetResponse()
                return @{ Success = $true; StatusCode = [int]$response.StatusCode }
            } catch {
                return @{ Success = $false; Error = $_.Exception.Message }
            }
        } -ArgumentList $BackendUrl
        
        if (Wait-Job $backendJob -Timeout 5) {
            $backendResult = Receive-Job $backendJob
            if ($backendResult.Success) {
                $testResults += "✅ Backend service responding (HTTP $($backendResult.StatusCode))"
            } else {
                $testResults += "⚠️  Backend service not responding - API proxy forwarding may fail"
            }
        } else {
            $testResults += "⚠️  Backend service test timed out - API proxy forwarding may fail"
        }
        Remove-Job $backendJob -Force
    } catch {
        $testResults += "⚠️  Backend service not responding - API proxy forwarding may fail"
    }

    # Ensure all background jobs are cleaned up to prevent hanging
    Write-Host "  Cleaning up test jobs..." -ForegroundColor Gray
    try {
        Get-Job | Where-Object { $_.Name -like "*test*" -or $_.State -eq "Running" } | Remove-Job -Force -ErrorAction SilentlyContinue
        Write-Host "  ✅ All test jobs cleaned up" -ForegroundColor Green
    } catch {
        Write-Host "  ⚠️  Warning: Some jobs may not have been cleaned up properly" -ForegroundColor Yellow
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
    if ($testResults -and $testResults.Count -gt 0) {
        foreach ($result in $testResults) {
            if ($result) {
                Write-Host "  $result"
            }
        }
    } else {
        Write-Host "  ⚠️  No test results available - check site manually" -ForegroundColor Yellow
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