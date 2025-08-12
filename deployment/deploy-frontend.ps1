# ExcelAddin Frontend Deployment Script
# Deploys React frontend as NSSM service

param(
    [switch]$Force,      # Kept for compatibility but not required for overwriting services
    [switch]$SkipBuild,
    [switch]$SkipInstall,
    [switch]$ConfigureIIS  # New parameter to optionally configure IIS
)

# Import common functions
. (Join-Path $PSScriptRoot "scripts" | Join-Path -ChildPath "common.ps1")

$ServiceName = "ExcelAddin-Frontend"
$ServiceDisplayName = "ExcelAddin Frontend Service"
$ServiceDescription = "Excel Add-in Frontend Web Server"
$FrontendServerScript = Join-Path $PSScriptRoot "config" | Join-Path -ChildPath "frontend-server.js"

Write-Header "ExcelAddin Frontend Deployment"

# Check prerequisites
if (-not (Test-Prerequisites -SkipPM2 -SkipNSSM:$false)) {
    Write-Error "Prerequisites check failed. Please resolve issues before continuing."
    exit 1
}

# Get paths
$ProjectRoot = Get-ProjectRoot
$FrontendPath = Get-FrontendPath

Write-Host "Project Root: $ProjectRoot"
Write-Host "Frontend Path: $FrontendPath"

# Navigate to project root
Push-Location $FrontendPath

try {
    # Install Node.js dependencies
    Write-Header "Installing Node.js Dependencies"
    
    Write-Host "Running npm install..."
    npm install
    if ($LASTEXITCODE -ne 0) {
        Write-Error "npm install failed"
        exit 1
    }
    Write-Success "Dependencies installed successfully"
    
    # Build the frontend
    if (-not $SkipBuild) {
        Write-Header "Building Frontend Application"
        
        Write-Host "Building for staging environment..."
        npm run build:staging
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Frontend build failed"
            exit 1
        }
        
        # Verify dist directory exists
        if (-not (Test-Path "dist")) {
            Write-Error "Build output directory 'dist' not found"
            exit 1
        }
        
        Write-Success "Frontend built successfully"
    } else {
        Write-Warning "Skipping build - using existing dist directory"
        if (-not (Test-Path "dist")) {
            Write-Error "dist directory not found. Cannot skip build."
            exit 1
        }
    }
    
    # Test the built frontend
    Write-Header "Testing Frontend Build"
    
    # Test using our new frontend server
    Write-Host "Testing frontend build with new server..."
    $nodeCmd = Get-Command node -ErrorAction Stop
    $testProcess = Start-Process -FilePath $nodeCmd.Source -ArgumentList $FrontendServerScript -PassThru -NoNewWindow -WorkingDirectory $FrontendPath
    Start-Sleep -Seconds 5
    
    if ($testProcess -and -not $testProcess.HasExited) {
        try {
            $response = Invoke-WebRequest -Uri "http://127.0.0.1:3000" -TimeoutSec 10
            Write-Success "Frontend test passed - HTTP Status: $($response.StatusCode)"
        } catch {
            Write-Warning "Frontend test failed, but continuing: $($_.Exception.Message)"
        }
        
        # Stop test process
        Stop-Process -Id $testProcess.Id -Force
        Start-Sleep -Seconds 2
    }
    
    # Configure NSSM service
    if (-not $SkipInstall) {
        Write-Header "Configuring NSSM Service"
        
        # Remove existing service if it exists (always overwrite by default)
        $existingService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if ($existingService) {
            Write-Host "Existing service detected. Stopping and removing service: $ServiceName"
            Stop-ServiceSafely -ServiceName $ServiceName
            nssm remove $ServiceName confirm
            Start-Sleep -Seconds 2
            Write-Success "Existing service removed successfully"
        }
        
        # Install NSSM service
        Write-Host "Installing NSSM service..."
        
        # Get Node.js executable path
        try {
            $nodeCmd = Get-Command node -ErrorAction Stop
            $nodePath = $nodeCmd.Source
            Write-Host "Using Node.js executable: $nodePath"
        } catch {
            Write-Error "Node.js executable not found in PATH. Please ensure Node.js is installed and accessible."
            exit 1
        }
        
        # Verify Node.js executable works
        try {
            $nodeVersion = & $nodePath --version 2>&1
            Write-Host "Node.js version: $nodeVersion"
        } catch {
            Write-Error "Node.js executable is not working properly: $($_.Exception.Message)"
            exit 1
        }
        
        # Verify frontend server script exists
        if (-not (Test-Path $FrontendServerScript)) {
            Write-Error "Frontend server script not found: $FrontendServerScript"
            exit 1
        }
        Write-Host "Using frontend server script: $FrontendServerScript"
        
        nssm install $ServiceName $nodePath $FrontendServerScript
        
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Failed to install NSSM service"
            exit 1
        }
        
        # Configure service parameters
        nssm set $ServiceName DisplayName $ServiceDisplayName
        nssm set $ServiceName Description $ServiceDescription
        nssm set $ServiceName AppDirectory $FrontendPath
        nssm set $ServiceName Start SERVICE_AUTO_START
        
        # Configure restart behavior for better reliability
        nssm set $ServiceName AppExit Default Restart
        nssm set $ServiceName AppRestartDelay 5000
        nssm set $ServiceName AppStopMethodSkip 0
        nssm set $ServiceName AppStopMethodConsole 10000
        nssm set $ServiceName AppStopMethodWindow 10000
        nssm set $ServiceName AppStopMethodThreads 10000
        nssm set $ServiceName AppThrottle 5000
        
        # Configure logging
        $logDir = "C:\Logs\ExcelAddin"
        if (-not (Test-Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }
        
        nssm set $ServiceName AppStdout "$logDir\frontend-stdout.log"
        nssm set $ServiceName AppStderr "$logDir\frontend-stderr.log"
        nssm set $ServiceName AppRotateFiles 1
        nssm set $ServiceName AppRotateOnline 1
        nssm set $ServiceName AppRotateSeconds 86400
        nssm set $ServiceName AppRotateBytes 10485760
        
        # Set environment variables
        nssm set $ServiceName AppEnvironmentExtra NODE_ENV=production PORT=3000 HOST=127.0.0.1
        
        Write-Success "NSSM service configured successfully"
        
        # Check for and resolve port conflicts before starting
        Write-Host "Checking for port conflicts..."
        try {
            $portConflict = Get-NetTCPConnection -LocalPort 3000 -ErrorAction SilentlyContinue
            if ($portConflict) {
                Write-Warning "Port 3000 is already in use by process ID: $($portConflict.OwningProcess)"
                $conflictProcess = Get-Process -Id $portConflict.OwningProcess -ErrorAction SilentlyContinue
                if ($conflictProcess) {
                    Write-Warning "Conflicting process: $($conflictProcess.ProcessName) (PID: $($conflictProcess.Id))"
                    
                    # Don't kill system processes (PID 0, 4) or critical Windows processes
                    if ($conflictProcess.Id -gt 4 -and $conflictProcess.ProcessName -notin @("System", "Idle", "svchost", "winlogon", "csrss")) {
                        Write-Host "Attempting to stop conflicting process..."
                        try {
                            Stop-Process -Id $conflictProcess.Id -Force -ErrorAction Stop
                            Write-Success "Successfully stopped process $($conflictProcess.ProcessName) (PID: $($conflictProcess.Id))"
                            Start-Sleep -Seconds 2
                            
                            # Verify port is now free
                            $portConflictAfter = Get-NetTCPConnection -LocalPort 3000 -ErrorAction SilentlyContinue
                            if (-not $portConflictAfter) {
                                Write-Success "Port 3000 is now available"
                            } else {
                                Write-Warning "Port 3000 is still in use after stopping the process"
                            }
                        } catch {
                            Write-Warning "Failed to stop conflicting process: $($_.Exception.Message)"
                            Write-Host "The service may fail to start due to this port conflict."
                        }
                    } else {
                        Write-Warning "Cannot stop system/critical process. The service may fail to start due to this port conflict."
                    }
                } else {
                    Write-Warning "Could not identify the conflicting process"
                }
            } else {
                Write-Success "Port 3000 is available"
            }
        } catch {
            Write-Host "Could not check for port conflicts (this is normal on some Windows versions)"
        }
    }
    # Start the service
    Write-Header "Starting Frontend Service"
    
    try {
        Start-Service -Name $ServiceName
        Start-Sleep -Seconds 5
        
        $service = Get-Service -Name $ServiceName
        if ($service.Status -eq "Running") {
            Write-Success "Frontend service started successfully"
            
            # Wait a bit longer for the service to fully initialize
            Write-Host "Waiting for service to initialize..."
            Start-Sleep -Seconds 10
            
            # Check if port is actually listening
            Write-Host "Verifying port 3000 is listening..."
            try {
                $portTest = Test-NetConnection -ComputerName "127.0.0.1" -Port 3000 -InformationLevel Quiet -ErrorAction SilentlyContinue
                if ($portTest) {
                    Write-Success "Port 3000 is listening"
                    
                    # Now test HTTP response
                    Write-Host "Testing HTTP response..."
                    try {
                        $response = Invoke-WebRequest -Uri "http://127.0.0.1:3000" -TimeoutSec 15
                        Write-Success "Frontend service is responding - HTTP Status: $($response.StatusCode)"
                    } catch {
                        Write-Warning "Port is listening but HTTP request failed: $($_.Exception.Message)"
                        Write-Host "This may indicate a server configuration issue."
                        
                        # Additional diagnostics for 404 errors
                        if ($_.Exception.Message -like "*404*") {
                            Write-Host ""
                            Write-Host "404 Error Diagnostics:"
                            
                            # Check dist directory
                            $distPath = Join-Path $ProjectRoot "dist"
                            Write-Host "Checking static file directory: $distPath"
                            if (Test-Path $distPath) {
                                $distFiles = Get-ChildItem $distPath -Recurse | Where-Object { -not $_.PSIsContainer } | Measure-Object
                                Write-Host "  Directory exists with $($distFiles.Count) files"
                                
                                # Check for key files
                                $indexPath = Join-Path $distPath "index.html"
                                if (Test-Path $indexPath) {
                                    Write-Host "  ✓ index.html found"
                                } else {
                                    Write-Warning "  ✗ index.html missing - this will cause 404 errors"
                                }
                                
                                # Show some files in dist
                                $sampleFiles = Get-ChildItem $distPath -File | Select-Object -First 5
                                if ($sampleFiles) {
                                    Write-Host "  Sample files in dist:"
                                    $sampleFiles | ForEach-Object { Write-Host "    - $($_.Name)" }
                                }
                            } else {
                                Write-Warning "  ✗ Static directory missing: $distPath"
                                Write-Host "    Run 'npm run build:staging' to create the dist directory"
                            }
                        }
                        
                        # Show recent service logs if available
                        $logFile = "C:\Logs\ExcelAddin\frontend-stderr.log"
                        if (Test-Path $logFile) {
                            Write-Host ""
                            Write-Host "Recent error logs:"
                            Get-Content $logFile -Tail 10 | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
                        }
                        
                        # Show stdout logs too for additional context
                        $stdoutLogFile = "C:\Logs\ExcelAddin\frontend-stdout.log"
                        if (Test-Path $stdoutLogFile) {
                            Write-Host ""
                            Write-Host "Recent service output:"
                            Get-Content $stdoutLogFile -Tail 15 | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
                        }
                    }
                } else {
                    Write-Warning "Service is running but port 3000 is not listening"
                    Write-Host "This indicates a server startup failure."
                    
                    # Show service logs for debugging
                    $logFile = "C:\Logs\ExcelAddin\frontend-stdout.log"
                    if (Test-Path $logFile) {
                        Write-Host ""
                        Write-Host "Recent service logs:"
                        Get-Content $logFile -Tail 10 | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
                    }
                    
                    $errorLogFile = "C:\Logs\ExcelAddin\frontend-stderr.log"
                    if (Test-Path $errorLogFile) {
                        Write-Host ""
                        Write-Host "Recent error logs:"
                        Get-Content $errorLogFile -Tail 10 | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
                    }
                }
            } catch {
                Write-Warning "Port connectivity test failed: $($_.Exception.Message)"
            }
        } else {
            Write-Error "Frontend service failed to start - Status: $($service.Status)"
            
            # Show service logs for debugging
            $logFile = "C:\Logs\ExcelAddin\frontend-stderr.log"
            if (Test-Path $logFile) {
                Write-Host ""
                Write-Host "Recent error logs:"
                Get-Content $logFile -Tail 10 | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
            }
            exit 1
        }
    } catch {
        Write-Error "Failed to start frontend service: $($_.Exception.Message)"
        exit 1
    }
    
    # Configure IIS if requested
    if ($ConfigureIIS) {
        Write-Header "Configuring IIS Reverse Proxy"
        
        $SiteName = "ExcelAddin"
        $Port = 9443
        
        try {
            # Import WebAdministration module
            Import-Module WebAdministration -ErrorAction SilentlyContinue
            if (-not (Get-Module WebAdministration)) {
                Write-Error "WebAdministration module not available. IIS may not be properly installed."
                exit 1
            }
            
            # Check if site already exists
            $existingSite = Get-IISSite -Name $SiteName -ErrorAction SilentlyContinue
            if ($existingSite) {
                Write-Host "Removing existing IIS site..."
                Remove-IISSite -Name $SiteName -Confirm:$false
                Start-Sleep -Seconds 2
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
            if (Test-Path $webConfigSource) {
                Copy-Item $webConfigSource $webConfigDest -Force
                Write-Host "Copied web.config to application directory"
            } else {
                Write-Warning "web.config not found at $webConfigSource"
            }
            
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
                        $bindCommand = "netsh http add sslcert ipport=0.0.0.0:$Port certhash=$CertificateThumbprint appid='{12345678-1234-1234-1234-123456789abc}' certstorename=MY"
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
            
            Write-Host ""
            Write-Host "IIS Configuration Complete:"
            Write-Host "  Site Name: $SiteName"
            Write-Host "  Port: $Port"
            Write-Host "  Application Path: $appPath"
            if ($CertificateThumbprint) {
                Write-Host "  SSL Certificate: Configured (Thumbprint: $CertificateThumbprint)"
            } else {
                Write-Host "  SSL Certificate: Not configured - place .pfx/.p12 file in C:\Cert\ and re-run deployment"
            }
            Write-Host ""
            
        } catch {
            Write-Warning "IIS configuration failed: $($_.Exception.Message)"
            Write-Host "You can configure IIS manually or run deploy-all.ps1 for complete setup"
        }
    }
    
    Write-Header "Frontend Deployment Complete"
    Write-Success "ExcelAddin Frontend has been deployed successfully"
    Write-Host ""
    Write-Host "Service Information:"
    Write-Host "  Name: $ServiceName"
    Write-Host "  Display Name: $ServiceDisplayName"
    Write-Host "  Status: Running"
    Write-Host "  URL: http://127.0.0.1:3000"
    Write-Host ""
    if ($ConfigureIIS) {
        Write-Host "IIS Configuration:"
        Write-Host "  Site: ExcelAddin"
        Write-Host "  Public URL: https://server-vs81t.intranet.local:9443"
        Write-Host ""
    } else {
        Write-Host "Note: Run with -ConfigureIIS to set up IIS reverse proxy"
        Write-Host "Or use deploy-all.ps1 for complete deployment"
        Write-Host ""
    }
    Write-Host "Service Management Commands:"
    Write-Host "  Status: Get-Service -Name '$ServiceName'"
    Write-Host "  Start: Start-Service -Name '$ServiceName'"
    Write-Host "  Stop: Stop-Service -Name '$ServiceName'"
    Write-Host "  Restart: Restart-Service -Name '$ServiceName'"
    Write-Host "  Logs: Check C:\Logs\ExcelAddin\frontend-*.log"
    Write-Host ""

} catch {
    Write-Error "Deployment failed: $($_.Exception.Message)"
    Write-Host $_.ScriptStackTrace
    exit 1
} finally {
    Pop-Location
}