<#
.SYNOPSIS
    Deploys Excel Add-in to existing IIS server
.DESCRIPTION
    Lightweight script to configure Excel Add-in on existing IIS deployment.
    Only creates the application and site without installing IIS features.
.PARAMETER Force
    Force recreate existing application pools and sites
.PARAMETER SiteName
    Name of the IIS site (default: ExcelAddin)
.PARAMETER Port
    HTTPS port for the site (default: 9443)
.EXAMPLE
    .\deploy-to-existing-iis.ps1
    .\deploy-to-existing-iis.ps1 -Force -SiteName "MyExcelApp" -Port 8443
#>

param(
    [switch]$Force = $false,
    [string]$SiteName = "ExcelAddin",
    [int]$Port = 9443
)

# Check if running as Administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script must be run as Administrator. Right-click PowerShell and select 'Run as Administrator'"
    exit 1
}

Write-Host "Deploying Excel Add-in to existing IIS server..." -ForegroundColor Green

# Variables
$AppPoolName = "${SiteName}AppPool"
$WebsiteRoot = "C:\inetpub\wwwroot\$SiteName"
$PhysicalPath = $WebsiteRoot
$CertPath = "C:\Cert\server-vs81t.crt"
$KeyPath = "C:\Cert\server-vs81t.key"
$ServerName = "server-vs81t.intranet.local"

try {
    # Step 1: Verify IIS is available
    Write-Host "1. Verifying IIS installation..." -ForegroundColor Cyan
    
    try {
        Import-Module WebAdministration -ErrorAction Stop
        $iisService = Get-Service W3SVC -ErrorAction Stop
        Write-Host "   IIS is available and ready" -ForegroundColor Green
    } catch {
        Write-Error "IIS is not properly installed or configured. Please install IIS first."
        exit 1
    }

    # Step 2: Create directory structure
    Write-Host "2. Setting up directory structure..." -ForegroundColor Cyan
    
    # Create main website root directory
    if (-not (Test-Path $PhysicalPath)) {
        New-Item -ItemType Directory -Path $PhysicalPath -Force
        Write-Host "   Created directory: $PhysicalPath" -ForegroundColor Green
    } else {
        Write-Host "   Directory already exists: $PhysicalPath" -ForegroundColor Green
    }
    
    # Create excellence subdirectory for React app
    $ExcellencePath = Join-Path $PhysicalPath "excellence"
    if (-not (Test-Path $ExcellencePath)) {
        New-Item -ItemType Directory -Path $ExcellencePath -Force
        Write-Host "   Created directory: $ExcellencePath" -ForegroundColor Green
    } else {
        Write-Host "   Directory already exists: $ExcellencePath" -ForegroundColor Green
    }

    # Step 3: Configure application pool
    Write-Host "3. Configuring application pool '$AppPoolName'..." -ForegroundColor Cyan
    
    $existingAppPool = Get-IISAppPool -Name $AppPoolName -ErrorAction SilentlyContinue
    if ($existingAppPool) {
        if ($Force) {
            Remove-WebAppPool -Name $AppPoolName
            Write-Host "   Removed existing application pool" -ForegroundColor Yellow
            $existingAppPool = $null
        } else {
            Write-Host "   Application pool exists, updating settings..." -ForegroundColor Green
        }
    }
    
    if (-not $existingAppPool) {
        New-WebAppPool -Name $AppPoolName
        Write-Host "   Created application pool '$AppPoolName'" -ForegroundColor Green
    }
    
    # Configure app pool settings
    Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name processModel.identityType -Value ApplicationPoolIdentity
    Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name processModel.loadUserProfile -Value $false
    Set-ItemProperty -Path "IIS:\AppPools\$AppPoolName" -Name recycling.periodicRestart.time -Value "00:00:00"
    Write-Host "   Configured application pool settings" -ForegroundColor Green

    # Step 4: Configure IIS site
    Write-Host "4. Configuring IIS site '$SiteName'..." -ForegroundColor Cyan
    
    $existingSite = Get-Website -Name $SiteName -ErrorAction SilentlyContinue
    if ($existingSite) {
        if ($Force) {
            Remove-Website -Name $SiteName
            Write-Host "   Removed existing website" -ForegroundColor Yellow
            $existingSite = $null
        } else {
            Write-Host "   Website exists, updating configuration..." -ForegroundColor Green
            Set-ItemProperty -Path "IIS:\Sites\$SiteName" -Name physicalPath -Value $PhysicalPath
            Set-ItemProperty -Path "IIS:\Sites\$SiteName" -Name applicationPool -Value $AppPoolName
        }
    }
    
    if (-not $existingSite) {
        # Check if port is already in use by another site
        $portInUse = Get-Website | Get-WebBinding | Where-Object {$_.bindingInformation -like "*:${Port}:*"}
        if ($portInUse -and -not $Force) {
            Write-Error "Port $Port is already in use by another site. Use -Force to override or choose a different port."
            exit 1
        }
        
        New-Website -Name $SiteName -PhysicalPath $PhysicalPath -ApplicationPool $AppPoolName
        New-WebBinding -Name $SiteName -Protocol "https" -Port $Port -SslFlags 0
        Write-Host "   Created website '$SiteName' on port $Port" -ForegroundColor Green
    }

    # Step 5: Configure web.config and copy files
    Write-Host "5. Configuring web.config and preparing file structure..." -ForegroundColor Cyan
    
    # Copy web.config to root directory
    $webConfigSource = Join-Path $PSScriptRoot "..\iis\web.config"
    $webConfigDest = Join-Path $PhysicalPath "web.config"
    
    if (Test-Path $webConfigSource) {
        Copy-Item $webConfigSource $webConfigDest -Force
        Write-Host "   Copied web.config to $webConfigDest" -ForegroundColor Green
    } else {
        Write-Error "web.config not found at: $webConfigSource"
        Write-Host "   Please ensure web.config is available in the deployment/iis directory" -ForegroundColor Yellow
    }
    
    # Check if dist directory exists and suggest copying files
    $DistPath = Join-Path (Get-Location).Path "dist"
    if (Test-Path $DistPath) {
        Write-Host "   Found dist directory at: $DistPath" -ForegroundColor Green
        Write-Host "   Copying dist files to excellence directory..." -ForegroundColor Cyan
        
        # Copy all files from dist to excellence directory
        Copy-Item -Path "$DistPath\*" -Destination $ExcellencePath -Recurse -Force
        Write-Host "   Copied all dist files to $ExcellencePath" -ForegroundColor Green
    } else {
        Write-Host "   dist directory not found. Run 'npm run build:staging' first" -ForegroundColor Yellow
        Write-Host "   Then copy dist/* files to: $ExcellencePath" -ForegroundColor Yellow
    }

    # Step 6: Configure SSL (if certificate exists)
    Write-Host "6. Configuring SSL certificate..." -ForegroundColor Cyan
    
    if (Test-Path $CertPath) {
        try {
            # Import certificate to certificate store
            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
            $cert.Import($CertPath)
            
            $store = New-Object System.Security.Cryptography.X509Certificates.X509Store([System.Security.Cryptography.X509Certificates.StoreName]::My, [System.Security.Cryptography.X509Certificates.StoreLocation]::LocalMachine)
            $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
            $store.Add($cert)
            $store.Close()
            
            # Bind certificate to site
            $binding = Get-WebBinding -Name $SiteName -Protocol "https"
            if ($binding) {
                $binding.AddSslCertificate($cert.Thumbprint, "my")
                Write-Host "   SSL certificate configured successfully" -ForegroundColor Green
            }
        } catch {
            Write-Warning "Could not automatically configure SSL certificate: $($_.Exception.Message)"
            Write-Host "   Please configure SSL certificate manually in IIS Manager" -ForegroundColor Yellow
        }
    } else {
        Write-Host "   Certificate not found at $CertPath - configure SSL manually" -ForegroundColor Yellow
    }

    # Step 7: Set directory permissions
    Write-Host "7. Setting directory permissions..." -ForegroundColor Cyan
    
    try {
        $acl = Get-Acl $PhysicalPath
        $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("IIS_IUSRS", "ReadAndExecute", "ContainerInherit,ObjectInherit", "None", "Allow")
        $acl.SetAccessRule($accessRule)
        $accessRule2 = New-Object System.Security.AccessControl.FileSystemAccessRule("IUSR", "ReadAndExecute", "ContainerInherit,ObjectInherit", "None", "Allow")
        $acl.SetAccessRule($accessRule2)
        Set-Acl -Path $PhysicalPath -AclObject $acl
        Write-Host "   Directory permissions configured" -ForegroundColor Green
    } catch {
        Write-Warning "Could not set directory permissions: $($_.Exception.Message)"
    }

    # Step 8: Configure firewall
    Write-Host "8. Configuring Windows Firewall..." -ForegroundColor Cyan
    
    try {
        # Remove old rules and add new one
        Get-NetFirewallRule -DisplayName "*Excel*${Port}*" -ErrorAction SilentlyContinue | Remove-NetFirewallRule
        New-NetFirewallRule -DisplayName "Excel Add-in IIS HTTPS (${Port})" -Direction Inbound -Protocol TCP -LocalPort $Port -Action Allow
        Write-Host "   Firewall rule added for port $Port" -ForegroundColor Green
    } catch {
        Write-Warning "Could not configure firewall: $($_.Exception.Message)"
        Write-Host "   Please add firewall rule manually for port $Port" -ForegroundColor Yellow
    }

    # Step 9: Enable ARR proxy (if available)
    Write-Host "9. Checking proxy configuration..." -ForegroundColor Cyan
    
    try {
        Set-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' -Filter 'system.webServer/proxy' -Name 'enabled' -Value 'true' -ErrorAction Stop
        Write-Host "   ARR proxy enabled" -ForegroundColor Green
    } catch {
        Write-Host "   ARR not available - install if API proxying is needed" -ForegroundColor Yellow
    }

    # Step 10: Start services
    Write-Host "10. Starting services..." -ForegroundColor Cyan
    
    try {
        Start-WebAppPool -Name $AppPoolName -ErrorAction SilentlyContinue
        Start-Website -Name $SiteName -ErrorAction SilentlyContinue
        Write-Host "   Services started successfully" -ForegroundColor Green
    } catch {
        Write-Warning "Could not start services: $($_.Exception.Message)"
    }

    Write-Host "`nDeployment completed successfully!" -ForegroundColor Green -BackgroundColor DarkGreen
    Write-Host "Website URL: https://$ServerName`:$Port/excellence/" -ForegroundColor Cyan
    Write-Host "Health check: https://$ServerName`:$Port/health" -ForegroundColor Cyan
    Write-Host "`nNext steps:" -ForegroundColor Yellow
    Write-Host "1. Build frontend: npm run build:staging" -ForegroundColor White
    Write-Host "2. Files are automatically copied to: $ExcellencePath" -ForegroundColor White
    Write-Host "3. Test: https://$ServerName`:$Port/excellence/taskpane.html" -ForegroundColor White
    Write-Host "4. Start Flask backend on port 5000" -ForegroundColor White

} catch {
    Write-Error "Deployment failed: $($_.Exception.Message)"
    Write-Host "Please check the error and run the script again" -ForegroundColor Red
    exit 1
}