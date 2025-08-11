# Test Flask backend standalone (outside of NSSM service)
# Usage: .\test-backend-standalone.ps1 [-Port 5000] [-Host "127.0.0.1"]
param(
    [int]$Port = 5000,
    [string]$Host = "127.0.0.1"
)

Write-Host "🧪 Testing Flask Backend Standalone" -ForegroundColor Cyan
Write-Host "This script runs the Flask backend directly for debugging" -ForegroundColor Yellow
Write-Host ""

# Check if Python is available
try {
    $pythonVersion = python --version 2>&1
    Write-Host "✅ Python found: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "❌ Python not found in PATH" -ForegroundColor Red
    Write-Host "💡 Make sure Python is installed and added to PATH" -ForegroundColor Yellow
    exit 1
}

# Navigate to backend directory
$backendPath = Join-Path $PSScriptRoot "..\..\backend"
if (-not (Test-Path $backendPath)) {
    Write-Host "❌ Backend directory not found: $backendPath" -ForegroundColor Red
    exit 1
}

Write-Host "📁 Backend directory: $backendPath" -ForegroundColor White
Set-Location $backendPath

# Check if pyproject.toml exists
if (-not (Test-Path "pyproject.toml")) {
    Write-Host "❌ pyproject.toml not found in backend directory" -ForegroundColor Red
    exit 1
}

# Check if main Flask app file exists (common names)
$appFiles = @("app.py", "main.py", "server.py", "run.py")
$appFile = $null
foreach ($file in $appFiles) {
    if (Test-Path $file) {
        $appFile = $file
        break
    }
}

if ($appFile -eq $null) {
    Write-Host "❌ Flask app file not found. Looking for: $($appFiles -join ', ')" -ForegroundColor Red
    Write-Host "💡 Please specify the correct Flask app file in this script" -ForegroundColor Yellow
    exit 1
}

Write-Host "✅ Flask app file found: $appFile" -ForegroundColor Green

# Install/upgrade dependencies
Write-Host ""
Write-Host "📦 Installing Python dependencies..." -ForegroundColor Yellow
try {
    # Check if Poetry is available
    $poetryCheck = & poetry --version 2>&1
    if ($LASTEXITCODE -eq 0) {
        poetry install --no-dev --quiet
        Write-Host "✅ Dependencies installed successfully with Poetry" -ForegroundColor Green
    } else {
        Write-Host "❌ Poetry not found. Please install Poetry first." -ForegroundColor Red
        Write-Host "   Run: (Invoke-WebRequest -Uri https://install.python-poetry.org -UseBasicParsing).Content | python -" -ForegroundColor Yellow
        exit 1
    }
} catch {
    Write-Host "⚠️  Warning: Some dependencies might have failed to install" -ForegroundColor Yellow
    Write-Host "    The app might still work if core dependencies are available" -ForegroundColor Yellow
}

# Set Flask environment variables for development/testing
$env:FLASK_ENV = "development"
$env:FLASK_DEBUG = "1"
$env:FLASK_APP = $appFile

Write-Host ""
Write-Host "🚀 Starting Flask backend..." -ForegroundColor Yellow
Write-Host "   Host: $Host" -ForegroundColor White
Write-Host "   Port: $Port" -ForegroundColor White
Write-Host "   App:  $appFile" -ForegroundColor White
Write-Host ""
Write-Host "💡 The backend will run in DEBUG mode for easier troubleshooting" -ForegroundColor Yellow
Write-Host "💡 Press Ctrl+C to stop the server" -ForegroundColor Yellow
Write-Host ""
Write-Host "🌐 API Health Check URLs:" -ForegroundColor Green
Write-Host "   http://$Host`:$Port/api/health" -ForegroundColor White
Write-Host "   http://$Host`:$Port/health" -ForegroundColor White
Write-Host ""

# Run Flask development server
try {
    python -m flask run --host=$Host --port=$Port --debug
} catch {
    Write-Host ""
    Write-Host "❌ Failed to start Flask server" -ForegroundColor Red
    Write-Host "💡 Try running manually:" -ForegroundColor Yellow
    Write-Host "   cd $backendPath" -ForegroundColor White
    Write-Host "   python $appFile" -ForegroundColor White
    exit 1
}