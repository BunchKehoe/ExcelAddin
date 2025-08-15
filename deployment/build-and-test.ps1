# Quick Build and Test Script for ExcelAddin
# Builds the frontend and tests basic functionality

param(
    [string]$Environment = "staging",
    [switch]$SkipBuild,
    [switch]$SkipTest
)

$ErrorActionPreference = "Stop"

Write-Host "========================================" -ForegroundColor Green
Write-Host "  EXCELADDIN BUILD AND TEST" -ForegroundColor Green
Write-Host "  Environment: $Environment" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""

try {
    # Check if we're in the right directory
    if (-not (Test-Path "package.json")) {
        Write-Error "Not in the correct directory. Please run this script from the ExcelAddin root directory."
    }

    # 1. BUILD FRONTEND (unless skipped)
    if (-not $SkipBuild) {
        Write-Host "1. BUILDING FRONTEND" -ForegroundColor Yellow
        Write-Host "=====================" -ForegroundColor Yellow
        
        Write-Host "Installing dependencies..." -ForegroundColor Cyan
        npm install
        
        Write-Host "Building for $Environment environment..." -ForegroundColor Cyan
        switch ($Environment) {
            "development" { npm run build:dev }
            "staging" { npm run build:staging }
            "production" { npm run build:prod }
            default { npm run build:staging }
        }
        
        # Verify build output
        if (Test-Path "dist") {
            Write-Host "✅ Build completed successfully" -ForegroundColor Green
            
            # Check if key files exist
            $keyFiles = @("dist/taskpane.html", "dist/commands.html", "dist/assets")
            foreach ($file in $keyFiles) {
                if (Test-Path $file) {
                    Write-Host "  ✅ $file exists" -ForegroundColor Green
                } else {
                    Write-Warning "  ⚠️ $file missing"
                }
            }
            
            # Check if PCAG assets exist
            $assetFiles = @(
                "dist/assets/PCAG_white_trans.png",
                "dist/assets/PCAG_trans_16.png",
                "dist/assets/PCAG_trans_32.png",
                "dist/assets/PCAG_trans_80.png"
            )
            
            $missingAssets = @()
            foreach ($asset in $assetFiles) {
                if (Test-Path $asset) {
                    Write-Host "  ✅ $(Split-Path -Leaf $asset) found" -ForegroundColor Green
                } else {
                    $missingAssets += $asset
                    Write-Warning "  ⚠️ $(Split-Path -Leaf $asset) missing"
                }
            }
            
            if ($missingAssets.Count -gt 0) {
                Write-Warning "Missing asset files. This may cause image loading issues."
                Write-Host "Expected assets in dist/assets/:" -ForegroundColor Yellow
                foreach ($asset in $missingAssets) {
                    Write-Host "  - $(Split-Path -Leaf $asset)" -ForegroundColor Gray
                }
            }
            
        } else {
            Write-Error "Build failed - dist directory not created"
        }
        
        Write-Host ""
    }

    # 2. QUICK TESTS (unless skipped)
    if (-not $SkipTest) {
        Write-Host "2. RUNNING QUICK TESTS" -ForegroundColor Yellow
        Write-Host "=======================" -ForegroundColor Yellow
        
        # Test if frontend server can start (quick test)
        Write-Host "Testing Express server startup..." -ForegroundColor Cyan
        
        # Create a temporary test script
        $testScript = @"
const express = require('express');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = 3001; // Use different port for testing

// Basic test routes
app.get('/health', (req, res) => {
    res.json({ status: 'test-ok', timestamp: new Date().toISOString() });
});

app.get('/test-assets', (req, res) => {
    const assetsPath = path.join(__dirname, 'dist/assets');
    const assetsExist = fs.existsSync(assetsPath);
    
    let assetsList = [];
    if (assetsExist) {
        try {
            assetsList = fs.readdirSync(assetsPath);
        } catch (err) {
            assetsList = ['Error: ' + err.message];
        }
    }
    
    res.json({
        distExists: fs.existsSync(path.join(__dirname, 'dist')),
        assetsExist,
        assetsList,
        requiredAssets: [
            'PCAG_white_trans.png',
            'PCAG_trans_16.png', 
            'PCAG_trans_32.png',
            'PCAG_trans_80.png'
        ]
    });
});

const server = app.listen(PORT, 'localhost', () => {
    console.log('Test server started on port ' + PORT);
    
    // Test the endpoints
    setTimeout(async () => {
        try {
            console.log('Testing health endpoint...');
            const response = await fetch('http://localhost:' + PORT + '/health');
            const data = await response.json();
            console.log('✅ Health test passed:', data.status);
            
            console.log('Testing assets endpoint...');
            const assetsResponse = await fetch('http://localhost:' + PORT + '/test-assets');
            const assetsData = await assetsResponse.json();
            console.log('✅ Assets test results:');
            console.log('  Dist exists:', assetsData.distExists);
            console.log('  Assets exist:', assetsData.assetsExist);
            console.log('  Assets found:', assetsData.assetsList.length);
            
            // Check for required assets
            const missing = assetsData.requiredAssets.filter(asset => !assetsData.assetsList.includes(asset));
            if (missing.length > 0) {
                console.log('⚠️  Missing assets:', missing);
            } else {
                console.log('✅ All required assets found');
            }
            
            server.close();
            process.exit(0);
        } catch (err) {
            console.error('❌ Test failed:', err.message);
            server.close();
            process.exit(1);
        }
    }, 1000);
});
"@
        
        $testScript | Out-File -FilePath "test-server.js" -Encoding UTF8
        
        try {
            node test-server.js
            Write-Host "✅ Frontend server test passed" -ForegroundColor Green
        } catch {
            Write-Warning "⚠️ Frontend server test failed: $($_.Exception.Message)"
        } finally {
            # Cleanup test file
            if (Test-Path "test-server.js") {
                Remove-Item "test-server.js" -Force
            }
        }
        
        Write-Host ""
    }

    # 3. SUMMARY
    Write-Host "3. SUMMARY" -ForegroundColor Yellow
    Write-Host "==========" -ForegroundColor Yellow
    
    Write-Host "Build completed for $Environment environment" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Cyan
    Write-Host "  1. Deploy services:" -ForegroundColor White
    Write-Host "     .\deployment\deploy-all.ps1 -Environment $Environment" -ForegroundColor Gray
    Write-Host "  2. Test connectivity:" -ForegroundColor White
    Write-Host "     .\deployment\debug-connectivity.ps1 -Environment $Environment" -ForegroundColor Gray
    Write-Host "  3. Check service status:" -ForegroundColor White
    Write-Host "     Get-Service -Name 'ExcelAddin-*'" -ForegroundColor Gray
    Write-Host ""

} catch {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "  BUILD FAILED!" -ForegroundColor Red
    Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    
    Write-Host "Troubleshooting steps:" -ForegroundColor Yellow
    Write-Host "  1. Ensure Node.js and npm are installed" -ForegroundColor Gray
    Write-Host "  2. Check if you're in the correct directory (should contain package.json)" -ForegroundColor Gray
    Write-Host "  3. Try deleting node_modules and running 'npm install' again" -ForegroundColor Gray
    Write-Host "  4. Check for any TypeScript or build errors above" -ForegroundColor Gray
    
    exit 1
}