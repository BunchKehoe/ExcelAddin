# ExcelAddin Complete Deployment Script
# Performs initial deployment of all services

param(
    [switch]$Force,
    [switch]$SkipIIS
)

# Import common functions
. (Join-Path $PSScriptRoot "scripts" | Join-Path -ChildPath "common.ps1")

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
if (-not (Test-Prerequisites -SkipPM2 -SkipNSSM:$false)) {
    Write-Error "Prerequisites check failed. Please resolve issues before continuing."
    exit 1
}

# Clean up any existing PM2 configurations
Write-Header "Step 0: Clean up PM2 (if present)"

$cleanupScript = Join-Path $PSScriptRoot "scripts" | Join-Path -ChildPath "cleanup-pm2.ps1"
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
        $appPath = Join-Path "C:\inetpub\wwwroot" $SiteName
        if (-not (Test-Path $appPath)) {
            New-Item -ItemType Directory -Path $appPath -Force | Out-Null
            Write-Host "Created application directory: $appPath"
        }
        
        # Copy web.config
        $webConfigSource = Join-Path $PSScriptRoot "config" | Join-Path -ChildPath "web.config"
        $webConfigDest = Join-Path $appPath "web.config"
        Copy-Item $webConfigSource $webConfigDest -Force
        Write-Host "Copied web.config to application directory"
        
        # Find SSL certificate for the domain
        Write-Host "Looking for SSL certificate for server-vs81t.intranet.local..."
        $CertificatePath = "C:\Cert"
        $CertificateThumbprint = ""
        
        # First, try to import certificate from C:\Cert\ if it exists
        if (Test-Path $CertificatePath) {
            Write-Host "Checking certificate directory: $CertificatePath"
            $certFiles = Get-ChildItem -Path $CertificatePath -Filter "*.pfx" -ErrorAction SilentlyContinue
            if (-not $certFiles) {
                $certFiles = Get-ChildItem -Path $CertificatePath -Filter "*.p12" -ErrorAction SilentlyContinue
            }
            
            if ($certFiles) {
                $certFile = $certFiles | Select-Object -First 1
                Write-Host "Found certificate file: $($certFile.Name)"
                try {
                    # Import certificate to local machine store
                    $importResult = Import-PfxCertificate -FilePath $certFile.FullName -CertStoreLocation "Cert:\LocalMachine\My" -Password (ConvertTo-SecureString -String "" -AsPlainText -Force) -ErrorAction SilentlyContinue
                    if ($importResult) {
                        $CertificateThumbprint = $importResult.Thumbprint
                        Write-Success "Imported SSL certificate: $($importResult.Subject) (Thumbprint: $CertificateThumbprint)"
                    } else {
                        Write-Warning "Failed to import certificate from $($certFile.FullName)"
                    }
                } catch {
                    Write-Warning "Could not import certificate from $($certFile.FullName): $($_.Exception.Message)"
                }
            }
        }
        
        # If no certificate imported from file, try to find existing certificate in store
        if (-not $CertificateThumbprint) {
            Write-Host "Searching certificate store for server-vs81t.intranet.local..."
            $certificates = Get-ChildItem -Path Cert:\LocalMachine\My | Where-Object { 
                $_.Subject -like "*server-vs81t.intranet.local*" -or 
                $_.DnsNameList -like "*server-vs81t.intranet.local*" 
            }
            
            if ($certificates) {
                $cert = $certificates | Select-Object -First 1
                $CertificateThumbprint = $cert.Thumbprint
                Write-Success "Found SSL certificate in store: $($cert.Subject) (Thumbprint: $CertificateThumbprint)"
            } else {
                Write-Warning "No SSL certificate found for server-vs81t.intranet.local in certificate store"
            }
        }
        
        if (-not $CertificateThumbprint) {
            Write-Warning "No SSL certificate available. Please place a .pfx/.p12 certificate file in C:\Cert\ or manually bind a certificate."
        }
        
        # Remove any conflicting sites on the same port
        Write-Host "Checking for conflicting sites on port $Port..."
        $conflictingSites = Get-IISSite | Where-Object {
            $_.Bindings | Where-Object { $_.BindingInformation -like "*:$Port*" }
        }
        foreach ($conflictingSite in $conflictingSites) {
            if ($conflictingSite.Name -ne $SiteName) {
                Write-Warning "Found conflicting site '$($conflictingSite.Name)' on port $Port. Please resolve manually."
            }
        }
        
        # Create IIS site as standalone site (not sub-site)
        Write-Host "Creating standalone IIS site: $SiteName"
        try {
            # Create site with proper binding information for standalone site
            $site = New-IISSite -Name $SiteName -PhysicalPath $appPath -BindingInformation "*:$Port:" -Protocol https
            Write-Success "IIS site '$SiteName' created successfully as standalone site"
            
            # Bind SSL certificate if found
            if ($CertificateThumbprint) {
                Write-Host "Binding SSL certificate to site..."
                try {
                    # Use netsh for more reliable certificate binding
                    $bindCommand = "netsh http add sslcert ipport=0.0.0.0:$Port certhash=$CertificateThumbprint appid={12345678-1234-1234-1234-123456789abc}"
                    Write-Host "Executing: $bindCommand"
                    
                    # First, try to remove any existing binding
                    $deleteCommand = "netsh http delete sslcert ipport=0.0.0.0:$Port"
                    Invoke-Expression $deleteCommand 2>$null
                    
                    # Add the new binding
                    $result = Invoke-Expression $bindCommand 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        Write-Success "SSL certificate bound successfully using netsh"
                    } else {
                        Write-Warning "netsh binding failed: $result"
                        # Fall back to IIS method
                        $binding = Get-IISSiteBinding -Name $SiteName -Protocol https -ErrorAction SilentlyContinue
                        if ($binding) {
                            $binding.AddSslCertificate($CertificateThumbprint, "my")
                            Write-Success "SSL certificate bound successfully using IIS method"
                        } else {
                            Write-Warning "Could not find HTTPS binding for site"
                        }
                    }
                } catch {
                    Write-Warning "Failed to bind SSL certificate automatically: $($_.Exception.Message)"
                    Write-Host "Please bind the certificate manually in IIS Manager:"
                    Write-Host "  1. Open IIS Manager"
                    Write-Host "  2. Select site '$SiteName'"
                    Write-Host "  3. Click 'Bindings...' in Actions panel"
                    Write-Host "  4. Select the HTTPS binding and click 'Edit...'"
                    Write-Host "  5. Select the SSL certificate with thumbprint: $CertificateThumbprint"
                }
            } else {
                Write-Warning "No SSL certificate available for automatic binding"
                Write-Host "Please manually configure SSL certificate in IIS Manager after placing certificate in C:\Cert\"
            }
        } catch {
            Write-Error "Failed to create IIS site: $($_.Exception.Message)"
            throw
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
    $frontendService = Get-Service -Name "ExcelAddin-Frontend" -ErrorAction SilentlyContinue
    if ($frontendService -and $frontendService.Status -eq "Running") {
        Write-Success "Frontend service: Running"
        try {
            $frontendCheck = Invoke-WebRequest -Uri "http://127.0.0.1:3000" -TimeoutSec 10
            Write-Success "Frontend health check: HTTP $($frontendCheck.StatusCode)"
        } catch {
            Write-Warning "Frontend health check failed: $($_.Exception.Message)"
        }
    } else {
        Write-Error "Frontend service is not running"
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
    Write-Host "- Frontend: ExcelAddin-Frontend (NSSM Service)"
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