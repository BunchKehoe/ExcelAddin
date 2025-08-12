# ExcelAddin Complete Deployment Script
# Performs initial deployment of all services

param(
    [switch]$Force,
    [switch]$SkipIIS
)

# Import common functions
. "$PSScriptRoot\scripts\common.ps1"

$SiteName = "ExcelAddin"
$Port = 9443
$CertificateThumbprint = ""  # Will be determined automatically

Write-Header "ExcelAddin Complete Deployment"
Write-Host "This script will deploy all ExcelAddin services:"
Write-Host "- Backend (Python Flask via NSSM)"
Write-Host "- Frontend (React via NSSM)"
Write-Host "- IIS Reverse Proxy (HTTPS on port $Port)"
Write-Host ""

# Check prerequisites
if (-not (Test-Prerequisites)) {
    Write-Error "Prerequisites check failed. Please resolve issues before continuing."
    exit 1
}

try {
    # Clean up any existing PM2 configurations
    Write-Header "Step 0: Clean up PM2 (if present)"
    
    $cleanupScript = "$PSScriptRoot\scripts\cleanup-pm2.ps1"
    if (Test-Path $cleanupScript) {
        try {
            & $cleanupScript -Force
            Write-Success "PM2 cleanup completed"
        } catch {
            Write-Warning "PM2 cleanup had issues, but continuing: $($_.Exception.Message)"
        }
    } else {
        Write-Host "PM2 cleanup script not found, skipping..."
    }

try {
    # Deploy Backend Service
    Write-Header "Step 1: Deploy Backend Service"
    
    $backendArgs = @()
    if ($Force) { $backendArgs += "-Force" }
    
    & "$PSScriptRoot\deploy-backend.ps1" @backendArgs
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Backend deployment failed"
        exit 1
    }
    Write-Success "Backend deployment completed"
    
    # Deploy Frontend Service  
    Write-Header "Step 2: Deploy Frontend Service"
    
    $frontendArgs = @()
    if ($Force) { $frontendArgs += "-Force" }
    
    & "$PSScriptRoot\deploy-frontend.ps1" @frontendArgs
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Frontend deployment failed"
        exit 1
    }
    Write-Success "Frontend deployment completed"
    
    # Configure IIS Reverse Proxy
    if (-not $SkipIIS) {
        Write-Header "Step 3: Configure IIS Reverse Proxy"
        
        # Import WebAdministration module
        Import-Module WebAdministration -ErrorAction SilentlyContinue
        if (-not (Get-Module WebAdministration)) {
            Write-Error "WebAdministration module not available. IIS may not be properly installed."
            exit 1
        }
        
        # Check if site already exists
        $existingSite = Get-IISSite -Name $SiteName -ErrorAction SilentlyContinue
        if ($existingSite) {
            if ($Force) {
                Write-Host "Removing existing IIS site..."
                Remove-IISSite -Name $SiteName -Confirm:$false
                Start-Sleep -Seconds 2
            } else {
                Write-Error "IIS site $SiteName already exists. Use -Force to overwrite."
                exit 1
            }
        }
        
        # Create application directory
        $appPath = "C:\inetpub\wwwroot\$SiteName"
        if (-not (Test-Path $appPath)) {
            New-Item -ItemType Directory -Path $appPath -Force | Out-Null
            Write-Host "Created application directory: $appPath"
        }
        
        # Copy web.config
        $webConfigSource = "$PSScriptRoot\config\web.config"
        $webConfigDest = "$appPath\web.config"
        Copy-Item $webConfigSource $webConfigDest -Force
        Write-Host "Copied web.config to application directory"
        
        # Find SSL certificate for the domain
        Write-Host "Looking for SSL certificate for server-vs81t.intranet.local..."
        $certificates = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { 
            $_.Subject -like "*server-vs81t.intranet.local*" -or 
            $_.DnsNameList -like "*server-vs81t.intranet.local*" 
        }
        
        if ($certificates) {
            $cert = $certificates | Select-Object -First 1
            $CertificateThumbprint = $cert.Thumbprint
            Write-Success "Found SSL certificate: $($cert.Subject) (Thumbprint: $CertificateThumbprint)"
        } else {
            Write-Warning "No SSL certificate found for server-vs81t.intranet.local"
            Write-Host "You will need to manually bind an SSL certificate to the site"
        }
        
        # Create IIS site
        Write-Host "Creating IIS site: $SiteName"
        New-IISSite -Name $SiteName -PhysicalPath $appPath -BindingInformation "*:$Port" -Protocol https
        
        # Bind SSL certificate if found
        if ($CertificateThumbprint) {
            Write-Host "Binding SSL certificate to site..."
            try {
                $binding = Get-IISSiteBinding -Name $SiteName -Protocol https
                $binding.AddSslCertificate($CertificateThumbprint, "my")
                Write-Success "SSL certificate bound successfully"
            } catch {
                Write-Warning "Failed to bind SSL certificate automatically: $($_.Exception.Message)"
                Write-Host "Please bind the certificate manually in IIS Manager"
            }
        }
        
        # Configure application pool
        $appPoolName = "$SiteName-AppPool"
        if (Get-IISAppPool -Name $appPoolName -ErrorAction SilentlyContinue) {
            Remove-IISAppPool -Name $appPoolName -Confirm:$false
        }
        
        New-IISAppPool -Name $appPoolName
        Set-IISAppPool -Name $appPoolName -ManagedRuntimeVersion ""  # No managed code needed for reverse proxy
        Set-ItemProperty -Path "IIS:\AppPools\$appPoolName" -Name processModel.identityType -Value ApplicationPoolIdentity
        Set-ItemProperty -Path "IIS:\Sites\$SiteName" -Name applicationPool -Value $appPoolName
        
        Write-Success "IIS application pool configured"
        
        # Start the site
        Start-IISSite -Name $SiteName
        Write-Success "IIS site started"
        
        # Configure Windows Firewall
        Write-Host "Configuring Windows Firewall..."
        $firewallRule = Get-NetFirewallRule -DisplayName "ExcelAddin HTTPS" -ErrorAction SilentlyContinue
        if (-not $firewallRule) {
            New-NetFirewallRule -DisplayName "ExcelAddin HTTPS" -Direction Inbound -Protocol TCP -LocalPort $Port -Action Allow
            Write-Success "Firewall rule added for port $Port"
        } else {
            Write-Host "Firewall rule already exists"
        }
        
        Write-Success "IIS configuration completed"
    } else {
        Write-Warning "IIS configuration skipped"
    }
    
    # Final verification
    Write-Header "Step 4: Deployment Verification"
    
    # Wait a moment for all services to stabilize
    Start-Sleep -Seconds 10
    
    # Check backend service
    $backendService = Get-Service -Name "ExcelAddin-Backend" -ErrorAction SilentlyContinue
    if ($backendService -and $backendService.Status -eq "Running") {
        Write-Success "Backend service: Running"
        try {
            $healthCheck = Invoke-RestMethod -Uri "http://127.0.0.1:5000/api/health" -TimeoutSec 10
            Write-Success "Backend health check: $($healthCheck.status)"
        } catch {
            Write-Warning "Backend health check failed: $($_.Exception.Message)"
        }
    } else {
        Write-Error "Backend service is not running"
    }
    
    # Check frontend service
    try {
        $pm2Status = pm2 jlist | ConvertFrom-Json | Where-Object { $_.name -eq "exceladdin-frontend" }
        if ($pm2Status -and $pm2Status.pm2_env.status -eq "online") {
            Write-Success "Frontend service: Online"
            try {
                $frontendCheck = Invoke-WebRequest -Uri "http://127.0.0.1:3000" -TimeoutSec 10
                Write-Success "Frontend health check: HTTP $($frontendCheck.StatusCode)"
            } catch {
                Write-Warning "Frontend health check failed: $($_.Exception.Message)"
            }
        } else {
            Write-Error "Frontend service is not online"
        }
    } catch {
        Write-Error "Could not check frontend service status"
    }
    
    # Check IIS site
    if (-not $SkipIIS) {
        $site = Get-IISSite -Name $SiteName -ErrorAction SilentlyContinue
        if ($site -and $site.State -eq "Started") {
            Write-Success "IIS site: Started"
        } else {
            Write-Warning "IIS site is not started"
        }
    }
    
    Write-Header "Deployment Complete"
    Write-Success "ExcelAddin has been deployed successfully!"
    Write-Host ""
    Write-Host "Service Information:"
    Write-Host "- Backend: ExcelAddin-Backend (NSSM Service)"
    Write-Host "  Health: http://127.0.0.1:5000/api/health"
    Write-Host "- Frontend: exceladdin-frontend (PM2 Service)"
    Write-Host "  Local: http://127.0.0.1:3000"
    Write-Host "- Public URL: https://server-vs81t.intranet.local:$Port"
    Write-Host ""
    Write-Host "Next Steps:"
    Write-Host "1. Verify SSL certificate is properly configured"
    Write-Host "2. Test the public URL from an Excel client"
    Write-Host "3. Run test-deployment.ps1 for comprehensive testing"
    Write-Host ""

} catch {
    Write-Error "Deployment failed: $($_.Exception.Message)"
    Write-Host $_.ScriptStackTrace
    exit 1
}