# Test script to manually run and test the frontend server
# This helps diagnose NSSM deployment issues

param(
    [int]$Port = 3000,
    [string]$Host = "127.0.0.1",
    [int]$TestTimeoutSeconds = 10
)

. (Join-Path $PSScriptRoot "scripts\common.ps1")

Write-Header "Frontend Server Manual Test"

# Get paths
$ProjectRoot = Get-ProjectRoot
$FrontendPath = Get-FrontendPath
$ServerScript = Join-Path $PSScriptRoot "config\frontend-server.js"

Write-Host "Project Root: $ProjectRoot"
Write-Host "Frontend Path: $FrontendPath"
Write-Host "Server Script: $ServerScript"
Write-Host "Target URL: http://${Host}:${Port}"

# Check prerequisites
Write-Host ""
Write-Host "Checking prerequisites..."

if (-not (Test-Path $ServerScript)) {
    Write-Error "Server script not found: $ServerScript"
    exit 1
}

$distPath = Join-Path $FrontendPath "dist"
if (-not (Test-Path $distPath)) {
    Write-Error "Dist directory not found: $distPath"
    Write-Host "Please run 'npm run build:staging' first"
    exit 1
}

$indexPath = Join-Path $distPath "index.html"
if (-not (Test-Path $indexPath)) {
    Write-Error "index.html not found: $indexPath"
    exit 1
}

Write-Success "Prerequisites OK"

# Check for port conflicts
Write-Host ""
Write-Host "Checking for port conflicts on $Port..."
try {
    $portConflict = Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue
    if ($portConflict) {
        $conflictProcess = Get-Process -Id $portConflict.OwningProcess -ErrorAction SilentlyContinue
        if ($conflictProcess) {
            Write-Warning "Port $Port is in use by: $($conflictProcess.ProcessName) (PID: $($conflictProcess.Id))"
            Write-Host "You may want to kill this process first:"
            Write-Host "  Stop-Process -Id $($conflictProcess.Id) -Force"
        }
    } else {
        Write-Success "Port $Port is available"
    }
} catch {
    Write-Success "Port $Port appears to be available"
}

# Set working directory
Write-Host ""
Write-Host "Setting working directory to: $FrontendPath"
Set-Location $FrontendPath

# Set environment variables
Write-Host "Setting environment variables..."
$env:NODE_ENV = "production"
$env:PORT = $Port.ToString()
$env:HOST = $Host

Write-Host "  NODE_ENV = $env:NODE_ENV"
Write-Host "  PORT = $env:PORT"
Write-Host "  HOST = $env:HOST"

# Test Node.js availability
Write-Host ""
Write-Host "Testing Node.js..."
try {
    $nodeVersion = node --version
    Write-Success "Node.js version: $nodeVersion"
} catch {
    Write-Error "Node.js not available: $($_.Exception.Message)"
    exit 1
}

Write-Host ""
Write-Host "Starting server manually..."
Write-Host "Press Ctrl+C to stop the server"
Write-Host ""

# Start the server in a separate process so we can test it
$serverJob = Start-Job -ScriptBlock {
    param($ServerScript, $FrontendPath, $Port, $Host)
    
    Set-Location $FrontendPath
    $env:NODE_ENV = "production"
    $env:PORT = $Port
    $env:HOST = $Host
    
    & node $ServerScript
} -ArgumentList $ServerScript, $FrontendPath, $Port, $Host

# Wait for server to start
Write-Host "Waiting for server to start..."
Start-Sleep -Seconds 3

# Test the server
Write-Host "Testing server response..."
$testSuccess = $false
$testAttempts = 0
$maxAttempts = 5

while (-not $testSuccess -and $testAttempts -lt $maxAttempts) {
    $testAttempts++
    try {
        Write-Host "Test attempt $testAttempts/$maxAttempts..."
        
        # Test port connection
        $portTest = Test-NetConnection -ComputerName $Host -Port $Port -InformationLevel Quiet -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        if ($portTest) {
            Write-Success "Port $Port is listening"
            
            # Test HTTP response
            $response = Invoke-WebRequest -Uri "http://${Host}:${Port}" -TimeoutSec $TestTimeoutSeconds -ErrorAction Stop
            Write-Success "HTTP test passed - Status: $($response.StatusCode)"
            Write-Host "Response content length: $($response.Content.Length) bytes"
            
            if ($response.Content -like "*<title>*") {
                Write-Success "HTML content detected - server is working correctly"
                $testSuccess = $true
            } else {
                Write-Warning "Response doesn't look like HTML content"
            }
        } else {
            Write-Warning "Port $Port is not responding"
        }
    } catch {
        Write-Warning "Test failed: $($_.Exception.Message)"
    }
    
    if (-not $testSuccess -and $testAttempts -lt $maxAttempts) {
        Write-Host "Waiting 2 seconds before retry..."
        Start-Sleep -Seconds 2
    }
}

# Show job output
Write-Host ""
Write-Host "Server output:"
$jobOutput = Receive-Job -Job $serverJob
if ($jobOutput) {
    $jobOutput | ForEach-Object { Write-Host "  $_" }
} else {
    Write-Warning "No output from server job"
}

# Clean up
Write-Host ""
Write-Host "Stopping server..."
Stop-Job -Job $serverJob -ErrorAction SilentlyContinue
Remove-Job -Job $serverJob -Force -ErrorAction SilentlyContinue

if ($testSuccess) {
    Write-Success "Manual test PASSED - Server is working correctly"
    Write-Host ""
    Write-Host "The issue is likely with NSSM configuration, not the server itself."
} else {
    Write-Error "Manual test FAILED - Server is not working"
    Write-Host ""
    Write-Host "This indicates a fundamental issue with the server setup."
}

Write-Host ""
Write-Host "Next steps:"
if ($testSuccess) {
    Write-Host "1. Check NSSM service configuration"
    Write-Host "2. Check NSSM service logs: C:\Logs\ExcelAddin\frontend-*.log"
    Write-Host "3. Verify NSSM service working directory and environment"
} else {
    Write-Host "1. Check if dist folder contains built files"
    Write-Host "2. Verify Node.js and npm installation"
    Write-Host "3. Try rebuilding: npm run build:staging"
}