<#
.SYNOPSIS
    Sets up IIS with required modules and configures Excel Add-in site
.DESCRIPTION
    This script configures the Excel Add-in on an existing IIS installation.
    Works with both new and existing IIS deployments, installing only what's needed.
.PARAMETER Force
    Force recreate existing application pools and sites
.PARAMETER ExistingIIS
    Optimize for existing IIS deployment (skip feature installation)
.EXAMPLE
    .\setup-iis.ps1 -ExistingIIS
    .\setup-iis.ps1 -Force -ExistingIIS
#>

param(
    [switch]$Force = $false,
    [switch]$ExistingIIS = $false
)

# Check if running as Administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script must be run as Administrator. Right-click PowerShell and select 'Run as Administrator'"
    exit 1
}

if ($ExistingIIS) {
    Write-Host "Setting up Excel Add-in on existing IIS server..." -ForegroundColor Green
} else {
    Write-Host "Setting up IIS for Excel Add-in..." -ForegroundColor Green
}

# Variables
$SiteName = "ExcelAddin"
$AppPoolName = "ExcelAddinAppPool"
$PhysicalPath = "C:\inetpub\wwwroot\ExcelAddin\dist"
$Port = 9443
$CertPath = "C:\Cert\server-vs81t.crt"
$KeyPath = "C:\Cert\server-vs81t.key"
$ServerName = "server-vs81t.intranet.local"

try {
    # Step 1: Check IIS status and optionally install features
    Write-Host "1. Checking IIS installation..." -ForegroundColor Cyan
    
    # Test if IIS is already installed
    $iisInstalled = $false
    try {
        Import-Module WebAdministration -ErrorAction Stop
        $iisService = Get-Service W3SVC -ErrorAction Stop
        $iisInstalled = $true
        Write-Host "   IIS is already installed and accessible" -ForegroundColor Green
    } catch {
        Write-Host "   IIS not fully installed or configured" -ForegroundColor Yellow
    }
    
    if (-not $iisInstalled -and -not $ExistingIIS) {
        Write-Host "   Installing IIS and required features..." -ForegroundColor Cyan
        
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
            try {
                Enable-WindowsOptionalFeature -Online -FeatureName $feature -All -NoRestart -ErrorAction Stop
            } catch {
                Write-Warning "   Could not install feature $feature : $($_.Exception.Message)"
            }
        }
    } elseif ($ExistingIIS) {
        Write-Host "   Skipping IIS installation (using existing IIS)" -ForegroundColor Green
    }

    # Step 2: Check for required modules (more flexible for existing IIS)
    Write-Host "2. Checking URL Rewrite and ARR modules..." -ForegroundColor Cyan
    
    # Ensure WebAdministration is available
    try {
        Import-Module WebAdministration -ErrorAction Stop
        Write-Host "   WebAdministration module loaded successfully" -ForegroundColor Green
    } catch {
        Write-Error "WebAdministration module not available. Please ensure IIS is properly installed."
        exit 1
    }
    
    # Check for URL Rewrite module
    $urlRewriteInstalled = $false
    try {
        Get-WebConfiguration -Filter "system.webServer/rewrite/rules" -PSPath "IIS:\" -ErrorAction Stop | Out-Null
        $urlRewriteInstalled = $true
        Write-Host "   URL Rewrite module is available" -ForegroundColor Green
    } catch {
        Write-Warning "   URL Rewrite module may not be installed"
        Write-Host "   Install from: https://www.iis.net/downloads/microsoft/url-rewrite" -ForegroundColor Yellow
    }
    
    # Check for ARR module
    $arrInstalled = $false
    try {
        Get-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' -Filter 'system.webServer/proxy' -Name 'enabled' -ErrorAction Stop | Out-Null
        $arrInstalled = $true
        Write-Host "   Application Request Routing is available" -ForegroundColor Green
    } catch {
        Write-Warning "   Application Request Routing may not be installed"
        Write-Host "   Install from: https://www.iis.net/downloads/microsoft/application-request-routing" -ForegroundColor Yellow
    }

    # Step 3: Create or update application pool
    Write-Host "3. Configuring application pool '$AppPoolName'..." -ForegroundColor Cyan
    
    $existingAppPool = Get-IISAppPool -Name $AppPoolName -ErrorAction SilentlyContinue
    if ($existingAppPool) {
        if ($Force) {
            Remove-WebAppPool -Name $AppPoolName
            Write-Host "   Removed existing application pool" -ForegroundColor Yellow
            $existingAppPool = $null
        } else {
            Write-Host "   Application pool already exists, updating settings..." -ForegroundColor Green
            # Update existing app pool settings
            Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name processModel.identityType -Value ApplicationPoolIdentity
            Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name processModel.loadUserProfile -Value $false
            Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name recycling.periodicRestart.time -Value "00:00:00"
            Write-Host "   Updated application pool settings" -ForegroundColor Green
        }
    }
    
    if (-not $existingAppPool) {
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

    # Step 5: Create or update IIS site
    Write-Host "5. Configuring IIS site '$SiteName'..." -ForegroundColor Cyan
    
    $existingSite = Get-Website -Name $SiteName -ErrorAction SilentlyContinue
    if ($existingSite) {
        if ($Force) {
            Remove-Website -Name $SiteName
            Write-Host "   Removed existing website" -ForegroundColor Yellow
            $existingSite = $null
        } else {
            Write-Host "   Website already exists, checking configuration..." -ForegroundColor Green
            
            # Check if the binding matches what we want
            $binding = Get-WebBinding -Name $SiteName -Protocol "https" -Port $Port -ErrorAction SilentlyContinue
            if (-not $binding) {
                # Add HTTPS binding if it doesn't exist
                try {
                    New-WebBinding -Name $SiteName -Protocol "https" -Port $Port -SslFlags 0
                    Write-Host "   Added HTTPS binding on port $Port" -ForegroundColor Green
                } catch {
                    Write-Warning "   Could not add HTTPS binding: $($_.Exception.Message)"
                }
            } else {
                Write-Host "   HTTPS binding already configured on port $Port" -ForegroundColor Green
            }
            
            # Update physical path and app pool
            Set-ItemProperty -Path "IIS:\Sites\$SiteName" -Name physicalPath -Value $PhysicalPath
            Set-ItemProperty -Path "IIS:\Sites\$SiteName" -Name applicationPool -Value $AppPoolName
            Write-Host "   Updated website configuration" -ForegroundColor Green
        }
    }
    
    if (-not $existingSite) {
        # Create the website with HTTPS binding
        New-Website -Name $SiteName -PhysicalPath $PhysicalPath -Port $Port -Ssl -ApplicationPool $AppPoolName
        Write-Host "   Created website '$SiteName' on port $Port" -ForegroundColor Green
    }

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

    # Step 8: Configure ARR proxy (if available)
    Write-Host "8. Configuring Application Request Routing..." -ForegroundColor Cyan
    
    if ($arrInstalled) {
        try {
            # Enable proxy functionality
            Set-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' -Filter 'system.webServer/proxy' -Name 'enabled' -Value 'true'
            Set-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' -Filter 'system.webServer/proxy' -Name 'preserveHostHeader' -Value 'false'
            Write-Host "   ARR proxy configured successfully" -ForegroundColor Green
        } catch {
            Write-Warning "Could not configure ARR: $($_.Exception.Message)"
            Write-Host "   This may affect API proxying functionality" -ForegroundColor Yellow
        }
    } else {
        Write-Host "   ARR not available - API proxying may not work properly" -ForegroundColor Yellow
        Write-Host "   Install ARR from: https://www.iis.net/downloads/microsoft/application-request-routing" -ForegroundColor Yellow
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
    
    try {
        # Start IIS service if not running
        $w3svc = Get-Service W3SVC -ErrorAction SilentlyContinue
        if ($w3svc -and $w3svc.Status -ne 'Running') {
            Start-Service W3SVC
            Write-Host "   Started IIS service" -ForegroundColor Green
        } elseif ($w3svc) {
            Write-Host "   IIS service already running" -ForegroundColor Green
        } else {
            Write-Warning "   IIS service not found"
        }
        
        # Start application pool
        $appPoolState = Get-WebAppPoolState -Name $AppPoolName
        if ($appPoolState.Value -ne 'Started') {
            Start-WebAppPool -Name $AppPoolName
            Write-Host "   Started application pool '$AppPoolName'" -ForegroundColor Green
        } else {
            Write-Host "   Application pool '$AppPoolName' already running" -ForegroundColor Green
        }
        
        # Start website
        $siteState = Get-WebsiteState -Name $SiteName
        if ($siteState.Value -ne 'Started') {
            Start-Website -Name $SiteName
            Write-Host "   Started website '$SiteName'" -ForegroundColor Green
        } else {
            Write-Host "   Website '$SiteName' already running" -ForegroundColor Green
        }
    } catch {
        Write-Warning "Error managing services: $($_.Exception.Message)"
        Write-Host "   Please check IIS Manager for service status" -ForegroundColor Yellow
    }

    Write-Host "`nIIS setup completed successfully!" -ForegroundColor Green -BackgroundColor DarkGreen
    Write-Host "Website URL: https://$ServerName`:$Port/excellence/" -ForegroundColor Cyan
    Write-Host "Health check: https://$ServerName`:$Port/health" -ForegroundColor Cyan
    Write-Host "`nNext steps:" -ForegroundColor Yellow
    Write-Host "1. Build frontend: npm run build:staging" -ForegroundColor White
    Write-Host "2. Copy dist/* files to: $PhysicalPath" -ForegroundColor White
    Write-Host "3. Test the website: https://$ServerName`:$Port/excellence/taskpane.html" -ForegroundColor White
    Write-Host "4. Ensure Flask backend is running on port 5000" -ForegroundColor White
    
    if (-not $urlRewriteInstalled -or -not $arrInstalled) {
        Write-Host "`nIMPORTANT: Missing IIS modules detected!" -ForegroundColor Red
        if (-not $urlRewriteInstalled) {
            Write-Host "- Install URL Rewrite: https://www.iis.net/downloads/microsoft/url-rewrite" -ForegroundColor Yellow
        }
        if (-not $arrInstalled) {
            Write-Host "- Install ARR: https://www.iis.net/downloads/microsoft/application-request-routing" -ForegroundColor Yellow
        }
        Write-Host "Without these modules, the application may not function correctly." -ForegroundColor Red
    }

} catch {
    Write-Error "Setup failed: $($_.Exception.Message)"
    Write-Host "Please check the error and run the script again" -ForegroundColor Red
    exit 1
}