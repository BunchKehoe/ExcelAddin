<#
.SYNOPSIS
    Sets up IIS with required modules and configures Excel Add-in site
.DESCRIPTION
    This script replaces nginx with IIS for hosting the Excel Add-in.
    Installs required IIS features, configures SSL, and sets up the application.
.EXAMPLE
    .\setup-iis.ps1 -Force
#>

param(
    [switch]$Force = $false
)

# Check if running as Administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script must be run as Administrator. Right-click PowerShell and select 'Run as Administrator'"
    exit 1
}

Write-Host "Setting up IIS for Excel Add-in..." -ForegroundColor Green

# Variables
$SiteName = "ExcelAddin"
$AppPoolName = "ExcelAddinAppPool"
$PhysicalPath = "C:\inetpub\wwwroot\ExcelAddin\dist"
$Port = 9443
$CertPath = "C:\Cert\server-vs81t.crt"
$KeyPath = "C:\Cert\server-vs81t.key"
$ServerName = "server-vs81t.intranet.local"

try {
    # Step 1: Enable IIS and required features
    Write-Host "1. Installing IIS and required features..." -ForegroundColor Cyan
    
    $features = @(
        "IIS-WebServerRole",
        "IIS-WebServer", 
        "IIS-CommonHttpFeatures",
        "IIS-HttpErrors",
        "IIS-HttpLogging",
        "IIS-HttpRedirect",
        "IIS-ApplicationDevelopment",
        "IIS-NetFxExtensibility45",
        "IIS-ISAPIExtensions",
        "IIS-ISAPIFilter",
        "IIS-DefaultDocument",
        "IIS-DirectoryBrowsing",
        "IIS-StaticContent",
        "IIS-Security",
        "IIS-RequestFiltering",
        "IIS-Performance",
        "IIS-WebServerManagementTools",
        "IIS-ManagementConsole",
        "IIS-IIS6ManagementCompatibility",
        "IIS-Metabase"
    )
    
    foreach ($feature in $features) {
        Enable-WindowsOptionalFeature -Online -FeatureName $feature -All -NoRestart
    }

    # Step 2: Install URL Rewrite and Application Request Routing
    Write-Host "2. Installing URL Rewrite and ARR modules..." -ForegroundColor Cyan
    Write-Host "   Please install these modules manually if not already installed:" -ForegroundColor Yellow
    Write-Host "   - URL Rewrite Module: https://www.iis.net/downloads/microsoft/url-rewrite" -ForegroundColor Yellow
    Write-Host "   - Application Request Routing: https://www.iis.net/downloads/microsoft/application-request-routing" -ForegroundColor Yellow
    
    # Check if modules are installed
    try {
        Import-Module WebAdministration -ErrorAction Stop
        Write-Host "   WebAdministration module loaded successfully" -ForegroundColor Green
    } catch {
        Write-Error "WebAdministration module not available. Please ensure IIS is properly installed."
        exit 1
    }

    # Step 3: Create application pool
    Write-Host "3. Creating application pool '$AppPoolName'..." -ForegroundColor Cyan
    
    if (Get-IISAppPool -Name $AppPoolName -ErrorAction SilentlyContinue) {
        if ($Force) {
            Remove-WebAppPool -Name $AppPoolName
            Write-Host "   Removed existing application pool" -ForegroundColor Yellow
        } else {
            Write-Host "   Application pool already exists (use -Force to recreate)" -ForegroundColor Yellow
        }
    }
    
    if (-not (Get-IISAppPool -Name $AppPoolName -ErrorAction SilentlyContinue)) {
        New-WebAppPool -Name $AppPoolName
        Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name processModel.identityType -Value ApplicationPoolIdentity
        Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name processModel.loadUserProfile -Value $false
        Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name recycling.periodicRestart.time -Value "00:00:00"
        Write-Host "   Created application pool '$AppPoolName'" -ForegroundColor Green
    }

    # Step 4: Create physical directory
    Write-Host "4. Creating physical directory..." -ForegroundColor Cyan
    
    if (-not (Test-Path $PhysicalPath)) {
        New-Item -ItemType Directory -Path $PhysicalPath -Force
        Write-Host "   Created directory: $PhysicalPath" -ForegroundColor Green
    } else {
        Write-Host "   Directory already exists: $PhysicalPath" -ForegroundColor Green
    }

    # Step 5: Remove existing site and create new one
    Write-Host "5. Creating IIS site '$SiteName'..." -ForegroundColor Cyan
    
    if (Get-Website -Name $SiteName -ErrorAction SilentlyContinue) {
        if ($Force) {
            Remove-Website -Name $SiteName
            Write-Host "   Removed existing website" -ForegroundColor Yellow
        } else {
            Write-Error "Website '$SiteName' already exists (use -Force to recreate)"
            exit 1
        }
    }
    
    # Create the website with HTTPS binding
    New-Website -Name $SiteName -PhysicalPath $PhysicalPath -Port $Port -Ssl -ApplicationPool $AppPoolName
    Write-Host "   Created website '$SiteName' on port $Port" -ForegroundColor Green

    # Step 6: Configure SSL certificate
    Write-Host "6. Configuring SSL certificate..." -ForegroundColor Cyan
    
    # Check if certificate files exist
    if (-not (Test-Path $CertPath)) {
        Write-Warning "Certificate file not found: $CertPath"
        Write-Host "   Please ensure SSL certificate is available at: $CertPath" -ForegroundColor Yellow
    } else {
        # Import certificate to certificate store
        try {
            # Try to import certificate to personal store
            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
            $cert.Import($CertPath)
            
            $store = New-Object System.Security.Cryptography.X509Certificates.X509Store([System.Security.Cryptography.X509Certificates.StoreName]::My, [System.Security.Cryptography.X509Certificates.StoreLocation]::LocalMachine)
            $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
            $store.Add($cert)
            $store.Close()
            
            # Bind certificate to site
            $binding = Get-WebBinding -Name $SiteName -Protocol "https"
            $binding.AddSslCertificate($cert.Thumbprint, "my")
            
            Write-Host "   SSL certificate configured successfully" -ForegroundColor Green
        } catch {
            Write-Warning "Could not automatically configure SSL certificate: $($_.Exception.Message)"
            Write-Host "   Please configure SSL certificate manually in IIS Manager" -ForegroundColor Yellow
            Write-Host "   Certificate path: $CertPath" -ForegroundColor Yellow
        }
    }

    # Step 7: Copy web.config
    Write-Host "7. Configuring web.config..." -ForegroundColor Cyan
    
    $webConfigSource = Join-Path $PSScriptRoot "..\iis\web.config"
    $webConfigDest = Join-Path $PhysicalPath "web.config"
    
    if (Test-Path $webConfigSource) {
        Copy-Item $webConfigSource $webConfigDest -Force
        Write-Host "   Copied web.config to $webConfigDest" -ForegroundColor Green
    } else {
        Write-Error "web.config not found at: $webConfigSource"
        Write-Host "   Please ensure web.config is available in the deployment/iis directory" -ForegroundColor Yellow
    }

    # Step 8: Configure ARR proxy
    Write-Host "8. Configuring Application Request Routing..." -ForegroundColor Cyan
    
    try {
        # Enable proxy functionality
        Set-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' -Filter 'system.webServer/proxy' -Name 'enabled' -Value 'true'
        Set-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' -Filter 'system.webServer/proxy' -Name 'preserveHostHeader' -Value 'false'
        Write-Host "   ARR proxy configured successfully" -ForegroundColor Green
    } catch {
        Write-Warning "Could not configure ARR: $($_.Exception.Message)"
        Write-Host "   Please ensure Application Request Routing is installed and configure manually" -ForegroundColor Yellow
    }

    # Step 9: Set permissions
    Write-Host "9. Setting directory permissions..." -ForegroundColor Cyan
    
    $acl = Get-Acl $PhysicalPath
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("IIS_IUSRS", "ReadAndExecute", "ContainerInherit,ObjectInherit", "None", "Allow")
    $acl.SetAccessRule($accessRule)
    $accessRule2 = New-Object System.Security.AccessControl.FileSystemAccessRule("IUSR", "ReadAndExecute", "ContainerInherit,ObjectInherit", "None", "Allow")
    $acl.SetAccessRule($accessRule2)
    Set-Acl -Path $PhysicalPath -AclObject $acl
    Write-Host "   Directory permissions configured" -ForegroundColor Green

    # Step 10: Configure firewall
    Write-Host "10. Configuring Windows Firewall..." -ForegroundColor Cyan
    
    try {
        # Remove existing rules for port 9443
        Get-NetFirewallRule -DisplayName "*Excel*9443*" -ErrorAction SilentlyContinue | Remove-NetFirewallRule
        Get-NetFirewallRule -DisplayName "*nginx*9443*" -ErrorAction SilentlyContinue | Remove-NetFirewallRule
        
        # Add new rule for IIS
        New-NetFirewallRule -DisplayName "Excel Add-in IIS HTTPS (9443)" -Direction Inbound -Protocol TCP -LocalPort 9443 -Action Allow
        Write-Host "   Firewall rule added for port $Port" -ForegroundColor Green
    } catch {
        Write-Warning "Could not configure firewall: $($_.Exception.Message)"
        Write-Host "   Please add firewall rule manually for port $Port" -ForegroundColor Yellow
    }

    # Step 11: Start services
    Write-Host "11. Starting IIS services..." -ForegroundColor Cyan
    
    Start-Service W3SVC -ErrorAction SilentlyContinue
    Start-WebAppPool -Name $AppPoolName
    Start-Website -Name $SiteName
    Write-Host "   IIS services started" -ForegroundColor Green

    Write-Host "`nIIS setup completed successfully!" -ForegroundColor Green -BackgroundColor DarkGreen
    Write-Host "Website URL: https://$ServerName`:$Port/excellence/" -ForegroundColor Cyan
    Write-Host "Health check: https://$ServerName`:$Port/health" -ForegroundColor Cyan
    Write-Host "`nNext steps:" -ForegroundColor Yellow
    Write-Host "1. Build frontend: npm run build:staging" -ForegroundColor White
    Write-Host "2. Copy dist/* files to: $PhysicalPath" -ForegroundColor White
    Write-Host "3. Test the website: https://$ServerName`:$Port/excellence/taskpane.html" -ForegroundColor White
    Write-Host "4. Ensure Flask backend is running on port 5000" -ForegroundColor White

} catch {
    Write-Error "Setup failed: $($_.Exception.Message)"
    Write-Host "Please check the error and run the script again" -ForegroundColor Red
    exit 1
}