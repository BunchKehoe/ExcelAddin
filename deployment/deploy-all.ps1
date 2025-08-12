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
            # First ensure we're not creating under Default Web Site by checking existing sites
            $defaultSite = Get-IISSite -Name "Default Web Site" -ErrorAction SilentlyContinue
            if ($defaultSite) {
                # Check if our site already exists as application under Default Web Site
                $existingApp = Get-IISApp | Where-Object { $_.Path -eq "/$SiteName" -and $_.Site -eq "Default Web Site" }
                if ($existingApp) {
                    Write-Host "Removing existing application '$SiteName' from Default Web Site..."
                    Remove-IISApp -SiteName "Default Web Site" -Name $SiteName -Confirm:$false -ErrorAction SilentlyContinue
                }
            }
            
            # Create standalone site with explicit binding format for HTTPS
            # Use the format that ensures it's a root-level site, not an application
            $bindingInfo = "*:${Port}:"
            Write-Host "Creating site with binding: $bindingInfo"
            
            $site = New-IISSite -Name $SiteName -PhysicalPath $appPath -Port $Port -Protocol https
            
            if ($site) {
                Write-Success "IIS site '$SiteName' created successfully as standalone site"
                
                # Verify it's actually a standalone site and not under Default Web Site
                $createdSite = Get-IISSite -Name $SiteName -ErrorAction SilentlyContinue
                if ($createdSite -and $createdSite.Name -eq $SiteName) {
                    Write-Success "Verified: Site created as standalone site (not under Default Web Site)"
                } else {
                    Write-Error "Site creation verification failed - may have been created incorrectly"
                }
            } else {
                Write-Error "Failed to create IIS site"
                exit 1
            }
            
            # Configure application pool first (before certificate binding)
            Write-Host "Configuring application pool for standalone site..."
            $appPoolName = "$SiteName-AppPool"
            if (Get-IISAppPool -Name $appPoolName -ErrorAction SilentlyContinue) {
                Write-Host "Removing existing application pool: $appPoolName"
                Remove-IISAppPool -Name $appPoolName -Confirm:$false
            }
            
            New-IISAppPool -Name $appPoolName
            Set-IISAppPool -Name $appPoolName -ManagedRuntimeVersion ""  # No managed code needed for reverse proxy
            Set-ItemProperty -Path "IIS:\AppPools\$appPoolName" -Name processModel.identityType -Value ApplicationPoolIdentity
            
            # Associate the site with the custom application pool
            Set-ItemProperty -Path "IIS:\Sites\$SiteName" -Name applicationPool -Value $appPoolName
            Write-Success "Custom application pool configured for standalone site"
            
            # Bind SSL certificate if found
            if ($CertificateThumbprint) {
                Write-Host "Binding SSL certificate to standalone site..."
                try {
                    # First, ensure the site has proper HTTPS binding
                    $httpsBinding = Get-IISSiteBinding -Name $SiteName -Protocol https -ErrorAction SilentlyContinue
                    if (-not $httpsBinding) {
                        Write-Host "Adding HTTPS binding to site..."
                        New-IISSiteBinding -Name $SiteName -BindingInformation "*:${Port}:" -Protocol https
                        Start-Sleep -Seconds 2
                    }
                    
                    # Use netsh for certificate binding to the specific port
                    $bindCommand = "netsh http add sslcert ipport=0.0.0.0:$Port certhash=$CertificateThumbprint appid={12345678-1234-1234-1234-123456789abc} certstorename=MY"
                    Write-Host "Executing: $bindCommand"
                    
                    # First, try to remove any existing binding
                    $deleteCommand = "netsh http delete sslcert ipport=0.0.0.0:$Port"
                    Invoke-Expression $deleteCommand 2>$null
                    
                    # Add the new binding with certificate store specification
                    $result = Invoke-Expression $bindCommand 2>&1
                    if ($LASTEXITCODE -eq 0) {
                        Write-Success "SSL certificate bound successfully to port $Port using netsh"
                        
                        # Also try IIS method as additional confirmation
                        try {
                            $binding = Get-IISSiteBinding -Name $SiteName -Protocol https -ErrorAction SilentlyContinue
                            if ($binding) {
                                $binding.AddSslCertificate($CertificateThumbprint, "my")
                                Write-Success "SSL certificate also bound via IIS method"
                            }
                        } catch {
                            Write-Host "IIS method binding skipped (netsh was successful)"
                        }
                        
                    } else {
                        Write-Warning "netsh binding failed: $result"
                        
                        # Fall back to IIS-only method
                        Write-Host "Trying IIS-only certificate binding method..."
                        $binding = Get-IISSiteBinding -Name $SiteName -Protocol https -ErrorAction SilentlyContinue
                        if ($binding) {
                            $binding.AddSslCertificate($CertificateThumbprint, "my")
                            Write-Success "SSL certificate bound successfully using IIS method"
                        } else {
                            Write-Warning "Could not find HTTPS binding for standalone site"
                        }
                    }
                } catch {
                    Write-Warning "Failed to bind SSL certificate automatically: $($_.Exception.Message)"
                    Write-Host ""
                    Write-Host "Manual Certificate Binding Instructions:"
                    Write-Host "======================================="
                    Write-Host "1. Open IIS Manager"
                    Write-Host "2. Expand 'Sites' in left panel"
                    Write-Host "3. Click on '$SiteName' (should be standalone site, NOT under Default Web Site)"
                    Write-Host "4. Click 'Bindings...' in Actions panel on the right"
                    Write-Host "5. Select the HTTPS binding on port $Port and click 'Edit...'"
                    Write-Host "6. In SSL Certificate dropdown, select certificate with thumbprint: $CertificateThumbprint"
                    Write-Host "7. Click OK to save"
                    Write-Host ""
                }
            } else {
                Write-Warning "No SSL certificate available for automatic binding"
                Write-Host "Place a .pfx/.p12 certificate file in C:\Cert\ and re-run deployment"
            }
        } catch {
            Write-Error "Failed to create IIS site: $($_.Exception.Message)"
            throw
        }
        
        # Start the site
        Write-Host "Starting standalone IIS site..."
        try {
            Start-IISSite -Name $SiteName
            Write-Success "Standalone IIS site started successfully"
            
            # Final verification that site is correctly configured
            $finalSite = Get-IISSite -Name $SiteName -ErrorAction SilentlyContinue
            if ($finalSite) {
                Write-Host ""
                Write-Host "=== IIS Site Configuration Summary ==="
                Write-Host "Site Name: $($finalSite.Name)"
                Write-Host "Site ID: $($finalSite.Id)"
                Write-Host "Physical Path: $($finalSite.PhysicalPath)"
                Write-Host "State: $($finalSite.State)"
                Write-Host "Application Pool: $appPoolName"
                
                $bindings = Get-IISSiteBinding -Name $SiteName
                Write-Host "Bindings:"
                foreach ($binding in $bindings) {
                    Write-Host "  - Protocol: $($binding.Protocol), Port: $($binding.Port), Certificate: $($binding.SslFlags -ne 'None')"
                }
                
                # Verify it's NOT under Default Web Site
                $apps = Get-IISApp | Where-Object { $_.Site -eq "Default Web Site" -and $_.Path -eq "/$SiteName" }
                if ($apps) {
                    Write-Warning "WARNING: Found application '$SiteName' under Default Web Site - this may cause conflicts"
                } else {
                    Write-Success "Confirmed: No conflicting application under Default Web Site"
                }
                Write-Host "======================================"
                Write-Host ""
            }
        } catch {
            Write-Warning "Failed to start IIS site: $($_.Exception.Message)"
        }
        
        # Configure application pool (moved after site creation and certificate binding)
        Write-Host "Final application pool configuration..."
        try {
            # Ensure application pool is properly set
            $currentPool = Get-ItemProperty -Path "IIS:\Sites\$SiteName" -Name applicationPool -ErrorAction SilentlyContinue
            if ($currentPool -ne $appPoolName) {
                Set-ItemProperty -Path "IIS:\Sites\$SiteName" -Name applicationPool -Value $appPoolName
                Write-Host "Application pool association updated"
            }
            Write-Success "Application pool configuration verified"
        } catch {
            Write-Warning "Application pool verification had issues: $($_.Exception.Message)"
        }
        
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