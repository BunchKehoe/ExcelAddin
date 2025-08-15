# IIS Configuration Script for ExcelAddin
# Sets up IIS reverse proxy for ExcelAddin services

param(
    [string]$SiteName = "ExcelAddin",
    [string]$AppPoolName = "ExcelAddin",
    [int]$Port = 9443,
    [string]$FrontendUrl = "http://localhost:3000",
    [string]$BackendUrl = "http://localhost:5000",
    [switch]$Force
)

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Green
Write-Host "  IIS Configuration for ExcelAddin" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

# Check if running as administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script must be run as Administrator"
}

# Check if IIS is installed
$iisFeatures = Get-WindowsOptionalFeature -Online -FeatureName "IIS-*" | Where-Object { $_.State -eq "Enabled" }
if (-not $iisFeatures) {
    Write-Error "IIS is not installed. Please install IIS first."
}

# Import WebAdministration module
Import-Module WebAdministration -ErrorAction Stop

Write-Host "Configuring IIS for ExcelAddin..." -ForegroundColor Yellow
Write-Host "  Site Name: $SiteName"
Write-Host "  Port: $Port" 
Write-Host "  Frontend URL: $FrontendUrl"
Write-Host "  Backend URL: $BackendUrl"
Write-Host ""

# Remove existing site if it exists
$existingSite = Get-Website -Name $SiteName -ErrorAction SilentlyContinue
if ($existingSite) {
    if ($Force) {
        Write-Host "Removing existing site..." -ForegroundColor Yellow
        Remove-Website -Name $SiteName
    } else {
        Write-Error "Site '$SiteName' already exists. Use -Force to override."
    }
}

# Remove existing app pool if it exists
$existingPool = Get-IISAppPool -Name $AppPoolName -ErrorAction SilentlyContinue
if ($existingPool) {
    if ($Force) {
        Write-Host "Removing existing application pool..." -ForegroundColor Yellow
        Remove-WebAppPool -Name $AppPoolName
    } else {
        Write-Error "Application pool '$AppPoolName' already exists. Use -Force to override."
    }
}

# Create Application Pool
Write-Host "Creating application pool..." -ForegroundColor Yellow
New-WebAppPool -Name $AppPoolName
Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name "processModel.identityType" -Value "ApplicationPoolIdentity"
Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name "enable32BitAppOnWin64" -Value $false
Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name "managedRuntimeVersion" -Value ""  # No managed code

# Create Website
Write-Host "Creating website..." -ForegroundColor Yellow
$sitePath = "C:\inetpub\wwwroot\$SiteName"
if (-not (Test-Path $sitePath)) {
    New-Item -ItemType Directory -Path $sitePath -Force | Out-Null
}

# Create a simple default page
$defaultContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>ExcelAddin Proxy</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; margin: 40px; }
        .status { background: #f0f9ff; padding: 20px; border-radius: 8px; border: 1px solid #0078d4; }
    </style>
</head>
<body>
    <h1>ExcelAddin IIS Proxy</h1>
    <div class="status">
        <h3>✅ IIS Proxy Active</h3>
        <p>This server is acting as a reverse proxy for the ExcelAddin services.</p>
        <ul>
            <li><a href="/excellence/taskpane.html">Excel Taskpane</a></li>
            <li><a href="/api/health">API Health Check</a></li>
        </ul>
    </div>
</body>
</html>
"@

$defaultContent | Out-File -FilePath (Join-Path $sitePath "default.htm") -Encoding UTF8

New-Website -Name $SiteName -Port $Port -PhysicalPath $sitePath -ApplicationPool $AppPoolName

# Install URL Rewrite module (if not already installed)
$urlRewriteModule = Get-WebConfigurationProperty -Filter "system.webServer/modules/add[@name='RewriteModule']" -Name "name" -ErrorAction SilentlyContinue
if (-not $urlRewriteModule) {
    Write-Warning "URL Rewrite module is not installed. Please install it from: https://www.iis.net/downloads/microsoft/url-rewrite"
    Write-Warning "Continuing without URL rewrite rules..."
} else {
    Write-Host "Configuring URL rewrite rules..." -ForegroundColor Yellow
    
    # Create web.config with rewrite rules
    $webConfigContent = @"
<?xml version="1.0" encoding="UTF-8"?>
<configuration>
  <system.webServer>
    <rewrite>
      <rules>
        <!-- Frontend routes to port 3000 -->
        <rule name="Frontend Assets" stopProcessing="true">
          <match url="^excellence/(.*)" />
          <action type="Rewrite" url="$FrontendUrl/excellence/{R:1}" />
        </rule>
        
        <!-- API routes to port 5000 -->
        <rule name="Backend API" stopProcessing="true">
          <match url="^api/(.*)" />
          <action type="Rewrite" url="$BackendUrl/api/{R:1}" />
        </rule>
        
        <!-- Root manifest and functions -->
        <rule name="Root Files" stopProcessing="true">
          <match url="^(manifest.*\.xml|functions\.json)$" />
          <action type="Rewrite" url="$FrontendUrl/{R:1}" />
        </rule>
      </rules>
    </rewrite>
    
    <!-- CORS headers for Excel Add-in -->
    <httpProtocol>
      <customHeaders>
        <add name="Access-Control-Allow-Origin" value="*" />
        <add name="Access-Control-Allow-Methods" value="GET, POST, PUT, DELETE, PATCH, OPTIONS" />
        <add name="Access-Control-Allow-Headers" value="X-Requested-With, content-type, Authorization" />
      </customHeaders>
    </httpProtocol>
    
    <!-- Default document -->
    <defaultDocument>
      <files>
        <add value="default.htm" />
      </files>
    </defaultDocument>
  </system.webServer>
</configuration>
"@
    
    $webConfigContent | Out-File -FilePath (Join-Path $sitePath "web.config") -Encoding UTF8
}

# Configure HTTPS binding if needed
if ($Port -eq 443 -or $Port -eq 9443) {
    Write-Host "Setting up HTTPS binding..." -ForegroundColor Yellow
    
    # Check for existing certificate
    $cert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { 
        $_.Subject -like "*server-vs81t*" -or $_.Subject -like "*localhost*" 
    } | Select-Object -First 1
    
    if ($cert) {
        Write-Host "Using existing certificate: $($cert.Subject)" -ForegroundColor Green
        New-WebBinding -Name $SiteName -Protocol "https" -Port $Port -SslFlags 1 -Thumbprint $cert.Thumbprint
    } else {
        Write-Warning "No suitable SSL certificate found for HTTPS binding"
        Write-Warning "You'll need to configure SSL certificates manually"
    }
}

# Test the site
Write-Host "Testing IIS configuration..." -ForegroundColor Yellow
Start-Sleep -Seconds 2

try {
    $testUrl = "http://localhost:$Port"
    $response = Invoke-WebRequest -Uri $testUrl -TimeoutSec 5 -ErrorAction SilentlyContinue
    if ($response.StatusCode -eq 200) {
        Write-Host "✅ IIS site is responding" -ForegroundColor Green
    } else {
        Write-Warning "IIS site returned status code: $($response.StatusCode)"
    }
} catch {
    Write-Warning "Could not test IIS site: $($_.Exception.Message)"
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  IIS Configuration Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

Write-Host "Site Details:" -ForegroundColor Cyan
Write-Host "  Name: $SiteName"
Write-Host "  Port: $Port"
Write-Host "  Path: $sitePath"
Write-Host "  App Pool: $AppPoolName"
Write-Host ""

Write-Host "Test URLs:" -ForegroundColor Cyan
Write-Host "  Site: http://localhost:$Port"
Write-Host "  Excel Taskpane: http://localhost:$Port/excellence/taskpane.html"
Write-Host "  API Health: http://localhost:$Port/api/health"
Write-Host ""

Write-Host "Management Commands:" -ForegroundColor Cyan
Write-Host "  Start Site: Start-Website -Name '$SiteName'"
Write-Host "  Stop Site: Stop-Website -Name '$SiteName'"
Write-Host "  Check Status: Get-Website -Name '$SiteName'"
Write-Host ""

if (-not $urlRewriteModule) {
    Write-Host "IMPORTANT: Install URL Rewrite module for full functionality!" -ForegroundColor Red
    Write-Host "Download: https://www.iis.net/downloads/microsoft/url-rewrite" -ForegroundColor Yellow
}