# Fix Frontend 404 Issues - Diagnose and repair missing static files
# This script specifically addresses the "Self-test successful: HTTP 404" issue

param(
    [switch]$AutoFix
)

. (Join-Path $PSScriptRoot "scripts\common.ps1")

Write-Header "Frontend 404 Issue Diagnostics and Fix"

$ServiceName = "ExcelAddin-Frontend"
$ProjectRoot = Get-ProjectRoot
$FrontendPath = Get-FrontendPath
$ServerScript = Join-Path $PSScriptRoot "config\frontend-server.js"

# Check service status first
$service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($service) {
    Write-Host "Service Status: $($service.Status)" -ForegroundColor $(if ($service.Status -eq 'Running') { 'Green' } else { 'Red' })
} else {
    Write-Error "Service '$ServiceName' not found! Run deploy-frontend.ps1 first."
    exit 1
}

Write-Host "Project Root: $ProjectRoot"
Write-Host "Frontend Path: $FrontendPath"
Write-Host ""

# Critical Path Analysis
Write-Host "=== CRITICAL PATH ANALYSIS ===" -ForegroundColor Cyan

$pathsToCheck = @{
    "Project Root" = $ProjectRoot
    "Frontend Path" = $FrontendPath  
    "Dist Directory" = (Join-Path $FrontendPath "dist")
    "Index.html" = (Join-Path $FrontendPath "dist\index.html")
    "Package.json" = (Join-Path $FrontendPath "package.json")
}

$missingPaths = @()
$criticalIssues = @()

foreach ($pathName in $pathsToCheck.Keys) {
    $pathValue = $pathsToCheck[$pathName]
    Write-Host "$pathName :" -ForegroundColor Yellow
    Write-Host "  Path: $pathValue"
    
    if (Test-Path $pathValue) {
        $item = Get-Item $pathValue
        if ($item.PSIsContainer) {
            Write-Host "  Status: EXISTS (Directory)" -ForegroundColor Green
            
            # Special check for dist directory
            if ($pathName -eq "Dist Directory") {
                try {
                    $contents = Get-ChildItem $pathValue -Force
                    Write-Host "  File Count: $($contents.Count)"
                    
                    if ($contents.Count -eq 0) {
                        Write-Host "  ERROR: Dist directory is EMPTY!" -ForegroundColor Red
                        $criticalIssues += "Dist directory exists but is empty"
                    } else {
                        Write-Host "  Contents:" -ForegroundColor Green
                        $contents | Select-Object -First 5 | ForEach-Object { 
                            Write-Host "    $($_.Name) ($($_.Length) bytes)" 
                        }
                        if ($contents.Count -gt 5) {
                            Write-Host "    ... and $($contents.Count - 5) more files"
                        }
                    }
                } catch {
                    Write-Host "  ERROR: Cannot list directory contents: $($_.Exception.Message)" -ForegroundColor Red
                    $criticalIssues += "Cannot access dist directory contents"
                }
            }
        } else {
            Write-Host "  Status: EXISTS (File)" -ForegroundColor Green
            Write-Host "  Size: $($item.Length) bytes"
            
            if ($item.Length -eq 0) {
                Write-Host "  WARNING: File is empty!" -ForegroundColor Yellow
                $criticalIssues += "$pathName file is empty"
            }
        }
    } else {
        Write-Host "  Status: MISSING" -ForegroundColor Red
        $missingPaths += $pathName
        
        if ($pathName -in @("Dist Directory", "Index.html")) {
            $criticalIssues += "$pathName is missing"
        }
    }
    Write-Host ""
}

# Test the server's path resolution directly
Write-Host "=== SERVER PATH RESOLUTION TEST ===" -ForegroundColor Cyan

Write-Host "Testing server script path resolution..."
Set-Location $FrontendPath

$env:NODE_ENV = "production"
$env:PORT = "3000"
$env:HOST = "127.0.0.1"

# Create a test script to check path resolution
$testScript = @"
const path = require('path');
const fs = require('fs');

// Same logic as frontend-server.js
let PROJECT_ROOT;
if (process.cwd().endsWith('deployment')) {
    PROJECT_ROOT = path.resolve(process.cwd(), '..');
} else if (fs.existsSync(path.join(process.cwd(), 'dist'))) {
    PROJECT_ROOT = process.cwd();
} else {
    PROJECT_ROOT = path.resolve(__dirname, '../..');
}

const STATIC_DIR = path.join(PROJECT_ROOT, 'dist');

console.log('=== PATH RESOLUTION ===');
console.log('Current working directory:', process.cwd());
console.log('Script directory (__dirname):', __dirname);
console.log('Resolved PROJECT_ROOT:', PROJECT_ROOT);
console.log('Resolved STATIC_DIR:', STATIC_DIR);
console.log('');

console.log('=== FILE SYSTEM CHECK ===');
console.log('STATIC_DIR exists:', fs.existsSync(STATIC_DIR));
if (fs.existsSync(STATIC_DIR)) {
    try {
        const contents = fs.readdirSync(STATIC_DIR);
        console.log('STATIC_DIR contents count:', contents.length);
        contents.slice(0, 5).forEach(file => console.log('  -', file));
        if (contents.length > 5) console.log('  ... and', contents.length - 5, 'more files');
    } catch (err) {
        console.log('Error reading STATIC_DIR:', err.message);
    }
} else {
    console.log('STATIC_DIR path does not exist!');
    
    // Check parent directories
    let checkPath = PROJECT_ROOT;
    console.log('Checking PROJECT_ROOT contents:');
    try {
        const rootContents = fs.readdirSync(checkPath);
        rootContents.forEach(item => {
            const itemPath = path.join(checkPath, item);
            const isDir = fs.statSync(itemPath).isDirectory();
            console.log('  ' + (isDir ? '[DIR]' : '[FILE]') + ' ' + item);
        });
    } catch (err) {
        console.log('Error checking PROJECT_ROOT:', err.message);
    }
}

const indexPath = path.join(STATIC_DIR, 'index.html');
console.log('');
console.log('index.html path:', indexPath);
console.log('index.html exists:', fs.existsSync(indexPath));
"@

$tempTestFile = Join-Path $env:TEMP "frontend-path-test.js"
$testScript | Out-File -FilePath $tempTestFile -Encoding UTF8

try {
    Write-Host "Running path resolution test..."
    $testOutput = & node $tempTestFile 2>&1
    $testOutput | ForEach-Object { Write-Host "  $_" -ForegroundColor Cyan }
} catch {
    Write-Host "Path resolution test failed: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    Remove-Item $tempTestFile -ErrorAction SilentlyContinue
}

# Analyze issues
Write-Host ""
Write-Host "=== ISSUE ANALYSIS ===" -ForegroundColor Cyan

if ($criticalIssues.Count -gt 0) {
    Write-Host "CRITICAL ISSUES FOUND:" -ForegroundColor Red
    $criticalIssues | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
} else {
    Write-Host "No critical file system issues detected" -ForegroundColor Green
}

if ($missingPaths.Count -gt 0) {
    Write-Host ""
    Write-Host "MISSING PATHS:" -ForegroundColor Yellow
    $missingPaths | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
}

# Auto-fix logic
Write-Host ""
Write-Host "=== RECOMMENDED FIXES ===" -ForegroundColor Cyan

$fixes = @()

if ("Dist Directory" -in $missingPaths -or "Index.html" -in $missingPaths) {
    $fixes += @{
        Issue = "Missing dist directory or index.html"
        Command = "npm run build:staging"
        Description = "Build the frontend application to create dist directory and files"
    }
}

if ($criticalIssues -contains "Dist directory exists but is empty") {
    $fixes += @{
        Issue = "Empty dist directory"
        Command = "npm run build:staging"
        Description = "Rebuild the frontend to populate dist directory"
    }
}

# Check if we need to restart the service
$needsServiceRestart = $false
if ($fixes.Count -gt 0) {
    $needsServiceRestart = $true
    $fixes += @{
        Issue = "Service may be caching old configuration"
        Command = "Restart-Service -Name '$ServiceName'"
        Description = "Restart the NSSM service to pick up new files"
    }
}

if ($fixes.Count -eq 0) {
    Write-Host "No obvious fixes needed - this may be a different issue" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Additional troubleshooting steps:"
    Write-Host "1. Run: .\test-frontend-server.ps1"
    Write-Host "2. Check NSSM working directory: nssm get $ServiceName AppDirectory"
    Write-Host "3. Check service logs: Get-Content C:\Logs\ExcelAddin\frontend-*.log"
} else {
    Write-Host "Recommended fixes:" -ForegroundColor Green
    for ($i = 0; $i -lt $fixes.Count; $i++) {
        $fix = $fixes[$i]
        Write-Host "  $($i + 1). $($fix.Issue)"
        Write-Host "     Command: $($fix.Command)" -ForegroundColor Yellow
        Write-Host "     $($fix.Description)"
        Write-Host ""
    }
    
    if ($AutoFix) {
        Write-Host "AUTO-FIX MODE: Applying fixes..." -ForegroundColor Green
        
        # Navigate to frontend directory
        Set-Location $FrontendPath
        
        # Fix 1: Rebuild if needed
        if ("npm run build:staging" -in ($fixes | ForEach-Object { $_.Command })) {
            Write-Host "Building frontend application..." -ForegroundColor Yellow
            npm run build:staging
            if ($LASTEXITCODE -eq 0) {
                Write-Success "Frontend built successfully"
            } else {
                Write-Error "Frontend build failed"
                exit 1
            }
        }
        
        # Fix 2: Restart service if needed
        if ($needsServiceRestart) {
            Write-Host "Restarting service..." -ForegroundColor Yellow
            try {
                Restart-Service -Name $ServiceName -ErrorAction Stop
                Start-Sleep -Seconds 5
                Write-Success "Service restarted successfully"
            } catch {
                Write-Warning "Service restart failed: $($_.Exception.Message)"
                Write-Host "Try manually: Restart-Service -Name '$ServiceName'"
            }
        }
        
        # Test the fix
        Write-Host "Testing fix..." -ForegroundColor Yellow
        Start-Sleep -Seconds 3
        
        try {
            $testResponse = Invoke-WebRequest -Uri "http://127.0.0.1:3000" -TimeoutSec 10 -ErrorAction Stop
            if ($testResponse.StatusCode -eq 200) {
                Write-Success "FIX SUCCESSFUL! Frontend is now responding correctly"
                Write-Host "Status: $($testResponse.StatusCode)"
                Write-Host "Content Length: $($testResponse.Content.Length) bytes"
            } else {
                Write-Warning "Fix applied but still getting HTTP $($testResponse.StatusCode)"
            }
        } catch {
            Write-Warning "Fix applied but HTTP test still fails: $($_.Exception.Message)"
            Write-Host ""
            Write-Host "Additional steps needed:"
            Write-Host "1. Check service status: Get-Service -Name '$ServiceName'"
            Write-Host "2. Check service logs: Get-Content C:\Logs\ExcelAddin\frontend-*.log"
            Write-Host "3. Run manual test: .\test-frontend-server.ps1"
        }
    }
}

Write-Host ""
Write-Header "404 Diagnostics Complete"

if (-not $AutoFix -and $fixes.Count -gt 0) {
    Write-Host ""
    Write-Host "To automatically apply fixes, run:" -ForegroundColor Green
    Write-Host "  .\fix-frontend-404.ps1 -AutoFix" -ForegroundColor Yellow
}