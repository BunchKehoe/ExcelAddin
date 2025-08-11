# DEPRECATED: This script has been replaced with setup-backend-iis.ps1
# PowerShell script redirecting to new IIS-based backend setup
# Usage: .\setup-backend-iis.ps1 -BackendPath "C:\inetpub\wwwroot\ExcelAddin\backend"

param(
    [Parameter(Mandatory=$false)]
    [string]$BackendPath = "C:\inetpub\wwwroot\ExcelAddin\backend",
    
    [Parameter(Mandatory=$false)]
    [string]$PythonPath = "",
    
    [Parameter(Mandatory=$false)]
    [switch]$Uninstall,
    
    [Parameter(Mandatory=$false)]
    [switch]$Force,
    
    [Parameter(Mandatory=$false)]
    [switch]$Debug
)

Write-Host "DEPRECATED: This script has been replaced" -ForegroundColor Yellow
Write-Host "Please use the new IIS-based backend setup instead:" -ForegroundColor Yellow
Write-Host ".\setup-backend-iis.ps1 -BackendPath `"$BackendPath`"" -ForegroundColor Cyan
Write-Host ""
Write-Host "The new approach hosts the backend directly in IIS," -ForegroundColor Green
Write-Host "eliminating the need for NSSM (Non-Sucking Service Manager)." -ForegroundColor Green
Write-Host ""
Write-Host "Press Enter to run the new script, or Ctrl+C to cancel..." -ForegroundColor Yellow
Read-Host

# Redirect to the new script
$scriptDir = Split-Path $PSCommandPath -Parent
$newScriptPath = Join-Path $scriptDir "setup-backend-iis.ps1"
if (Test-Path $newScriptPath) {
    & $newScriptPath @PSBoundParameters
} else {
    Write-Error "New script not found at: $newScriptPath"
    Write-Host "Please run: .\deployment\scripts\setup-backend-iis.ps1" -ForegroundColor Yellow
}
exit
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Administrator)) {
    Write-Error "This script must be run as Administrator"
    exit 1
}

Write-Host "Excel Add-in Backend Windows Service Setup" -ForegroundColor Cyan
Write-Host "=" * 50

# Validate backend directory
if (-not (Test-Path $BackendPath)) {
    Write-Error "Backend directory not found: $BackendPath"
    exit 1
}

if (-not (Test-Path "$BackendPath\service_wrapper.py")) {
    Write-Error "service_wrapper.py not found in $BackendPath"
    exit 1
}

# Find Python executable
function Find-PythonExecutable {
    $pythonPaths = @()
    
    # Check if provided path works
    if ($PythonPath -and (Test-Path $PythonPath)) {
        $pythonPaths += $PythonPath
    }
    
    # Check common Python locations
    $commonPaths = @(
        "python",
        "python.exe",
        "C:\Python39\python.exe",
        "C:\Python310\python.exe",
        "C:\Python311\python.exe",
        "C:\Python312\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python39\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python310\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe"
    )
    
    foreach ($path in $commonPaths) {
        try {
            $fullPath = (Get-Command $path -ErrorAction SilentlyContinue).Source
            if ($fullPath) {
                $pythonPaths += $fullPath
            }
        } catch {
            # Try as direct path
            if (Test-Path $path) {
                $pythonPaths += $path
            }
        }
    }
    
    # Test each Python path
    foreach ($path in $pythonPaths) {
        try {
            $version = & "$path" --version 2>&1
            if ($LASTEXITCODE -eq 0 -and $version -match "Python \d+\.\d+") {
                Write-Host "[OK] Found Python: $path ($version)" -ForegroundColor Green
                return $path
            }
        } catch {
            continue
        }
    }
    
    return $null
}

$pythonExe = Find-PythonExecutable
if (-not $pythonExe) {
    Write-Error "Python executable not found. Please install Python or specify -PythonPath parameter."
    Write-Host "Common Python installation locations:" -ForegroundColor Yellow
    Write-Host "• C:\Python39\python.exe" -ForegroundColor Yellow
    Write-Host "• %LOCALAPPDATA%\Programs\Python\Python39\python.exe" -ForegroundColor Yellow
    exit 1
}

# Check if NSSM is available
try {
    $nssmVersion = & nssm version 2>$null
    if ($LASTEXITCODE -ne 0) {
        throw "NSSM not found"
    }
    Write-Host "[OK] NSSM found: $nssmVersion" -ForegroundColor Green
} catch {
    Write-Error "NSSM (Non-Sucking Service Manager) is required but not found."
    Write-Host "Please download and install NSSM from: https://nssm.cc/" -ForegroundColor Yellow
    Write-Host "1. Download nssm from https://nssm.cc/download" -ForegroundColor Yellow
    Write-Host "2. Extract to C:\Tools\nssm (or add to PATH)" -ForegroundColor Yellow
    Write-Host "3. Re-run this script" -ForegroundColor Yellow
    exit 1
}

# Uninstall existing service if requested
if ($Uninstall) {
    Write-Host "Uninstalling backend service..." -ForegroundColor Yellow
    
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($service) {
        if ($service.Status -eq 'Running') {
            Write-Host "Stopping backend service..."
            Stop-Service -Name $ServiceName -Force
            Start-Sleep -Seconds 5
        }
        
        & nssm remove $ServiceName confirm
        if ($LASTEXITCODE -eq 0) {
            Write-Host "[OK] Backend service removed successfully" -ForegroundColor Green
        } else {
            Write-Error "Failed to remove backend service"
        }
    } else {
        Write-Host "Backend service not found" -ForegroundColor Yellow
    }
    exit 0
}

# Stop existing service if it exists
$existingService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($existingService) {
    if ($Force) {
        Write-Host "Removing existing backend service..." -ForegroundColor Yellow
        if ($existingService.Status -eq 'Running') {
            Stop-Service -Name $ServiceName -Force
            Start-Sleep -Seconds 5
        }
        & nssm remove $ServiceName confirm
        Start-Sleep -Seconds 3
    } else {
        Write-Error "Service '$ServiceName' already exists. Use -Force to replace it."
        exit 1
    }
}

# Create log directory
$logDir = "C:\Logs\ExcelAddin"
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    Write-Host "[OK] Created log directory: $logDir" -ForegroundColor Green
}

# Test Python dependencies
Write-Host "Testing Python environment..." -ForegroundColor Yellow
Push-Location $BackendPath
try {
    # Test if service_wrapper can be imported
    $testScript = @"
import sys
sys.path.insert(0, r'$BackendPath')
try:
    import service_wrapper
    print("SUCCESS: service_wrapper imported successfully")
except Exception as e:
    print(f"ERROR: Failed to import service_wrapper: {e}")
    sys.exit(1)

try:
    from app import create_app
    app = create_app()
    print("SUCCESS: Flask app created successfully")
except Exception as e:
    print(f"ERROR: Failed to create Flask app: {e}")
    sys.exit(1)
"@
    
    $testResult = & "$pythonExe" -c $testScript 2>&1
    Write-Host "Python test result: $testResult"
    
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Python environment test failed. Please check dependencies."
        Write-Host "Try running: poetry install && poetry shell" -ForegroundColor Yellow
        exit 1
    }
    
    Write-Host "[OK] Python environment test passed" -ForegroundColor Green
    
} catch {
    Write-Error "Error testing Python environment: $_"
    exit 1
} finally {
    Pop-Location
}

# Install backend service with NSSM
Write-Host "Installing backend service..." -ForegroundColor Yellow

# Use full absolute paths for reliability
$serviceWrapperPath = Join-Path $BackendPath "service_wrapper.py"

# Install service with full paths
& nssm install $ServiceName "$pythonExe" "$serviceWrapperPath"
if ($LASTEXITCODE -ne 0) {
    Write-Error "Failed to install backend service"
    exit 1
}

# Configure service settings
Write-Host "Configuring service settings..." -ForegroundColor Yellow

# Basic service configuration
& nssm set $ServiceName DisplayName "Excel Add-in Backend Service"
& nssm set $ServiceName Description "Python Flask backend service for Excel Add-in"
& nssm set $ServiceName Start SERVICE_AUTO_START

# Set working directory to backend path
& nssm set $ServiceName AppDirectory "$BackendPath"

# Configure logging with rotation
& nssm set $ServiceName AppStdout "$logDir\backend-service-stdout.log"
& nssm set $ServiceName AppStderr "$logDir\backend-service-stderr.log"
& nssm set $ServiceName AppRotateFiles 1
& nssm set $ServiceName AppRotateOnline 1
& nssm set $ServiceName AppRotateBytes 10485760  # 10MB

# Set environment variables for production
$envVars = @(
    "FLASK_ENV=production",
    "DEBUG=false", 
    "HOST=127.0.0.1",
    "PORT=5000",
    "PYTHONPATH=$BackendPath",
    "PYTHONIOENCODING=utf-8",
    "PYTHONUNBUFFERED=1"
)
& nssm set $ServiceName AppEnvironmentExtra ($envVars -join "`0")

# Windows-specific optimizations
& nssm set $ServiceName AppPriority NORMAL_PRIORITY_CLASS
& nssm set $ServiceName AppNoConsole 1

# Configure service recovery options
& nssm set $ServiceName AppThrottle 5000  # Wait 5 seconds between restart attempts
& nssm set $ServiceName AppExit Default Restart
& nssm set $ServiceName AppRestartDelay 10000  # Wait 10 seconds before restart

# Configure service shutdown (give Flask time to cleanup)
& nssm set $ServiceName AppStopMethodSkip 0
& nssm set $ServiceName AppStopMethodConsole 15000
& nssm set $ServiceName AppStopMethodWindow 15000
& nssm set $ServiceName AppStopMethodThreads 15000

Write-Host "[OK] Backend service configured successfully" -ForegroundColor Green

# Create debugging batch file
Write-Host "Creating debugging tools..." -ForegroundColor Yellow

$debugBat = @"
@echo off
echo Excel Add-in Backend Service Debug
echo =================================
echo.
echo Python executable: $pythonExe
echo Backend directory: $BackendPath
echo Service wrapper: $serviceWrapperPath
echo.
echo Testing Python environment...
cd /d "$BackendPath"
"$pythonExe" -c "import sys; print('Python version:', sys.version); print('Python path:', sys.path)"
echo.
echo Testing service wrapper import...
"$pythonExe" -c "import service_wrapper; print('Service wrapper imported successfully')"
echo.
echo Testing Flask app creation...
"$pythonExe" -c "from app import create_app; app = create_app(); print('Flask app created successfully')"
echo.
echo Running service wrapper directly (Ctrl+C to stop)...
"$pythonExe" "$serviceWrapperPath"
"@

$debugBat | Set-Content -Path "$BackendPath\debug-service.bat" -Encoding ASCII

$serviceMgmtBat = @"
@echo off
echo Excel Add-in Backend Service Management
echo =====================================
echo.
echo Current service status:
sc query $ServiceName
echo.
echo Service configuration:
nssm dump $ServiceName
echo.
echo Recent service logs (stdout):
echo --- STDOUT LOG ---
type "$logDir\backend-service-stdout.log" 2>nul | findstr /n /c:"" | tail -20
echo.
echo --- STDERR LOG ---
type "$logDir\backend-service-stderr.log" 2>nul | findstr /n /c:"" | tail -20
echo.
echo Management commands:
echo   Start:   net start $ServiceName
echo   Stop:    net stop $ServiceName
echo   Restart: net stop $ServiceName ^&^& net start $ServiceName
echo   Edit:    nssm edit $ServiceName
echo   Remove:  nssm remove $ServiceName
pause
"@

$serviceMgmtBat | Set-Content -Path "$BackendPath\manage-service.bat" -Encoding ASCII

Write-Host "[OK] Debug tools created in $BackendPath" -ForegroundColor Green

# Start the service
Write-Host "Starting backend service..." -ForegroundColor Yellow
Start-Service -Name $ServiceName

# Wait and check service status
Start-Sleep -Seconds 10
$service = Get-Service -Name $ServiceName
Write-Host "Service status: $($service.Status)" -ForegroundColor $(if ($service.Status -eq 'Running') { 'Green' } else { 'Red' })

if ($service.Status -eq 'Running') {
    Write-Host "[OK] Backend service is running successfully" -ForegroundColor Green
    
    # Test if Flask is responding
    Write-Host "Testing Flask backend response..." -ForegroundColor Yellow
    Start-Sleep -Seconds 5
    try {
        $response = Invoke-WebRequest -Uri "http://127.0.0.1:5000/health" -TimeoutSec 10 -ErrorAction SilentlyContinue
        if ($response.StatusCode -eq 200) {
            Write-Host "[OK] Flask backend is responding on port 5000" -ForegroundColor Green
        } else {
            Write-Warning "Flask backend returned status: $($response.StatusCode)"
        }
    } catch {
        Write-Warning "Could not connect to Flask backend on port 5000"
        Write-Host "This may be normal if the service is still starting up." -ForegroundColor Yellow
    }
} else {
    Write-Warning "Backend service failed to start. Status: $($service.Status)"
    Write-Host "Checking recent logs..." -ForegroundColor Yellow
    
    # Show recent logs
    if (Test-Path "$logDir\backend-service-stderr.log") {
        Write-Host "`n--- STDERR LOG (last 20 lines) ---" -ForegroundColor Yellow
        Get-Content "$logDir\backend-service-stderr.log" -Tail 20 | ForEach-Object { Write-Host $_ }
    }
    
    if (Test-Path "$logDir\backend-service-stdout.log") {
        Write-Host "`n--- STDOUT LOG (last 20 lines) ---" -ForegroundColor Yellow  
        Get-Content "$logDir\backend-service-stdout.log" -Tail 20 | ForEach-Object { Write-Host $_ }
    }
}

Write-Host "`n" + "=" * 50 -ForegroundColor Cyan
Write-Host "SERVICE SETUP COMPLETE" -ForegroundColor Cyan
Write-Host "=" * 50 -ForegroundColor Cyan

Write-Host "`nService Configuration:" -ForegroundColor Green
Write-Host "• Name: $ServiceName"
Write-Host "• Python: $pythonExe"
Write-Host "• Backend: $BackendPath" 
Write-Host "• Service Wrapper: $serviceWrapperPath"

Write-Host "`nService Management:" -ForegroundColor Green
Write-Host "• Start:   Start-Service $ServiceName (or net start $ServiceName)"
Write-Host "• Stop:    Stop-Service $ServiceName (or net stop $ServiceName)"
Write-Host "• Restart: Restart-Service $ServiceName"
Write-Host "• Status:  Get-Service $ServiceName"
Write-Host "• Edit:    nssm edit $ServiceName"

Write-Host "`nDebugging Tools:" -ForegroundColor Green
Write-Host "• Debug script: $BackendPath\debug-service.bat"
Write-Host "• Management: $BackendPath\manage-service.bat"
Write-Host "• Manual test: cd $BackendPath && $pythonExe service_wrapper.py"

Write-Host "`nLog Files:" -ForegroundColor Green
Write-Host "• Service STDOUT: $logDir\backend-service-stdout.log"
Write-Host "• Service STDERR: $logDir\backend-service-stderr.log"
Write-Host "• Application Log: $logDir\backend-service.log"

if ($service.Status -ne 'Running') {
    Write-Host "`nTROUBLESHOoting:" -ForegroundColor Yellow
    Write-Host "1. Run debug-service.bat to test the service manually"
    Write-Host "2. Check the log files listed above for error messages"  
    Write-Host "3. Ensure all Python dependencies are installed: poetry install && poetry shell"
    Write-Host "4. Try running manually: cd $BackendPath && $pythonExe service_wrapper.py"
    Write-Host "5. Use 'nssm edit $ServiceName' to modify service settings"
    Write-Host "6. Check Windows Event Viewer for additional service errors"
} else {
    Write-Host "`nService is running successfully!" -ForegroundColor Green
    Write-Host "Flask backend should be available at http://127.0.0.1:5000" -ForegroundColor Green
}