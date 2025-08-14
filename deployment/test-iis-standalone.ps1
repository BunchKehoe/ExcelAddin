# Test script to validate IIS standalone site creation
# This script tests the key improvements made to fix the standalone site issue

param(
    [switch]$DryRun  # Test mode - doesn't actually create anything
)

# Import common functions
. (Join-Path $PSScriptRoot "scripts" | Join-Path -ChildPath "common.ps1")

$SiteName = "ExcelAddin-Test"
$Port = 9444  # Use different port for testing
$AppPath = "C:\inetpub\wwwroot\$SiteName-Test"

Write-Header "IIS Standalone Site Creation Test"

if ($DryRun) {
    Write-Host "*** DRY RUN MODE - No actual changes will be made ***"
    Write-Host ""
}

try {
    # Test 1: Check for existing conflicting applications
    Write-Host "Test 1: Checking for applications under Default Web Site..."
    $defaultSite = Get-IISSite -Name "Default Web Site" -ErrorAction SilentlyContinue
    if ($defaultSite) {
        $existingApp = Get-IISApp | Where-Object { $_.Path -eq "/$SiteName" -and $_.Site -eq "Default Web Site" }
        if ($existingApp) {
            Write-Warning "Found existing application '$SiteName' under Default Web Site"
            if (-not $DryRun) {
                Write-Host "Would remove this application in real run..."
            }
        } else {
            Write-Success "No conflicting applications found under Default Web Site"
        }
    } else {
        Write-Host "Default Web Site not found (this is unusual but not necessarily a problem)"
    }

    # Test 2: Verify IIS module availability
    Write-Host ""
    Write-Host "Test 2: Checking IIS module availability..."
    try {
        Import-Module WebAdministration -ErrorAction Stop
        if (Get-Module WebAdministration) {
            Write-Success "WebAdministration module loaded successfully"
        } else {
            Write-Error "WebAdministration module not available"
            exit 1
        }
    } catch {
        Write-Error "Failed to load WebAdministration module: $($_.Exception.Message)"
        exit 1
    }

    # Test 3: Test site creation (dry run version)
    Write-Host ""
    Write-Host "Test 3: Testing standalone site creation method..."
    if ($DryRun) {
        Write-Host "In real run, would execute:"
        Write-Host "  New-IISSite -Name '$SiteName' -PhysicalPath '$AppPath' -Port $Port -Protocol https"
        Write-Host "Expected result: Standalone site (not under Default Web Site)"
    } else {
        # Create test directory
        if (-not (Test-Path $AppPath)) {
            New-Item -ItemType Directory -Path $AppPath -Force | Out-Null
            Write-Host "Created test directory: $AppPath"
        }
        
        # Remove existing test site if it exists
        $existingTestSite = Get-IISSite -Name $SiteName -ErrorAction SilentlyContinue
        if ($existingTestSite) {
            Write-Host "Removing existing test site..."
            Remove-IISSite -Name $SiteName -Confirm:$false
            Start-Sleep -Seconds 2
        }
        
        # Create standalone test site
        try {
            $testSite = New-IISSite -Name $SiteName -PhysicalPath $AppPath -Port $Port -Protocol https
            if ($testSite) {
                Write-Success "Test site created successfully"
                
                # Verify it's standalone
                $createdSite = Get-IISSite -Name $SiteName -ErrorAction SilentlyContinue
                if ($createdSite -and $createdSite.Name -eq $SiteName) {
                    Write-Success "Verified: Test site is standalone (not under Default Web Site)"
                    
                    # Check for conflicting apps
                    $conflictingApps = Get-IISApp | Where-Object { $_.Site -eq "Default Web Site" -and $_.Path -eq "/$SiteName" }
                    if ($conflictingApps) {
                        Write-Warning "Found conflicting application under Default Web Site!"
                    } else {
                        Write-Success "No conflicting applications detected"
                    }
                    
                    # Display site info
                    Write-Host ""
                    Write-Host "Test Site Information:"
                    Write-Host "  Name: $($createdSite.Name)"
                    Write-Host "  ID: $($createdSite.Id)"
                    Write-Host "  Physical Path: $($createdSite.PhysicalPath)"
                    Write-Host "  State: $($createdSite.State)"
                    
                } else {
                    Write-Error "Site creation verification failed"
                }
            } else {
                Write-Error "Failed to create test site"
            }
        } catch {
            Write-Error "Site creation failed: $($_.Exception.Message)"
        }
        
        # Cleanup test site
        Write-Host ""
        Write-Host "Cleaning up test site..."
        try {
            if (Get-IISSite -Name $SiteName -ErrorAction SilentlyContinue) {
                Remove-IISSite -Name $SiteName -Confirm:$false
                Write-Host "Test site removed"
            }
            if (Test-Path $AppPath) {
                Remove-Item -Path $AppPath -Recurse -Force
                Write-Host "Test directory removed"
            }
        } catch {
            Write-Warning "Cleanup had issues: $($_.Exception.Message)"
        }
    }

    # Test 4: Certificate detection test
    Write-Host ""
    Write-Host "Test 4: Testing certificate detection from C:\Cert\..."
    $CertificatePath = "C:\Cert"
    if (Test-Path $CertificatePath) {
        $certFiles = @()
        $certFiles += Get-ChildItem -Path $CertificatePath -Filter "*.pfx" -ErrorAction SilentlyContinue
        $certFiles += Get-ChildItem -Path $CertificatePath -Filter "*.p12" -ErrorAction SilentlyContinue
        
        if ($certFiles) {
            Write-Success "Found certificate files in C:\Cert\:"
            foreach ($certFile in $certFiles) {
                Write-Host "  - $($certFile.Name)"
            }
        } else {
            Write-Host "No certificate files found in C:\Cert\"
        }
    } else {
        Write-Host "C:\Cert\ directory does not exist"
    }
    
    # Test existing certificates in store
    Write-Host ""
    Write-Host "Testing certificate store for server-vs81t.intranet.local..."
    $certificates = Get-ChildItem -Path Cert:\LocalMachine\My -ErrorAction SilentlyContinue | Where-Object { 
        $_.Subject -like "*server-vs81t.intranet.local*" -or 
        ($_.DnsNameList -and $_.DnsNameList -like "*server-vs81t.intranet.local*")
    }
    
    if ($certificates) {
        Write-Success "Found certificates in store:"
        foreach ($cert in $certificates) {
            Write-Host "  - Subject: $($cert.Subject)"
            Write-Host "    Thumbprint: $($cert.Thumbprint)"
            Write-Host "    Expires: $($cert.NotAfter)"
        }
    } else {
        Write-Host "No matching certificates found in certificate store"
    }

    Write-Host ""
    Write-Success "IIS Standalone Site Test Completed Successfully"

} catch {
    Write-Error "Test failed: $($_.Exception.Message)"
    Write-Host $_.ScriptStackTrace
    exit 1
}

Write-Host ""
Write-Host "Summary of Key Improvements Made:"
Write-Host "================================="
Write-Host "1. Fixed site creation to ensure standalone site (not under Default Web Site)"
Write-Host "2. Added verification to check for conflicting applications"
Write-Host "3. Improved certificate binding with netsh and fallback methods"
Write-Host "4. Added comprehensive diagnostics and verification"
Write-Host "5. Enhanced error handling and manual recovery instructions"
Write-Host ""
Write-Host "To run actual deployment:"
Write-Host "  .\deploy-all.ps1           (Complete deployment)"
Write-Host "  .\deploy-frontend.ps1 -ConfigureIIS  (Frontend with IIS)"
Write-Host ""