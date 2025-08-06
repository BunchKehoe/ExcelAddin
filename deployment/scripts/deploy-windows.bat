@echo off
REM Excel Add-in Windows Server Deployment Script
REM This script deploys the Excel Add-in to a Windows Server with nginx
REM Run as Administrator

echo ================================================================
echo Excel Add-in Windows Server Deployment
echo ================================================================
echo.

REM Check if running as administrator
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: This script must be run as Administrator
    echo Please right-click and select "Run as administrator"
    pause
    exit /b 1
)

REM Configuration variables (update these for your environment)
set DEPLOY_DIR=C:\inetpub\wwwroot\ExcelAddin
set LOG_DIR=C:\Logs\ExcelAddin
set NGINX_DIR=C:\nginx
set SERVICE_NAME=ExcelAddinBackend
set DOMAIN_NAME=your-staging-domain.com

echo Deployment Configuration:
echo - Deploy Directory: %DEPLOY_DIR%
echo - Log Directory: %LOG_DIR%
echo - nginx Directory: %NGINX_DIR%
echo - Service Name: %SERVICE_NAME%
echo - Domain Name: %DOMAIN_NAME%
echo.

REM Create directories
echo Creating directories...
if not exist "%DEPLOY_DIR%" mkdir "%DEPLOY_DIR%"
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"
if not exist "%NGINX_DIR%\logs" mkdir "%NGINX_DIR%\logs"
if not exist "C:\Logs\nginx" mkdir "C:\Logs\nginx"

REM Stop existing services
echo Stopping existing services...
sc query %SERVICE_NAME% >nul 2>&1
if %errorlevel% equ 0 (
    echo Stopping %SERVICE_NAME% service...
    net stop %SERVICE_NAME%
)

REM Stop nginx if running
tasklist /FI "IMAGENAME eq nginx.exe" 2>NUL | find /I /N "nginx.exe">NUL
if "%ERRORLEVEL%"=="0" (
    echo Stopping nginx...
    taskkill /F /IM nginx.exe >nul 2>&1
)

REM Copy application files
echo Copying application files...
xcopy /E /Y dist "%DEPLOY_DIR%\dist\"
xcopy /E /Y backend "%DEPLOY_DIR%\backend\"
xcopy /Y manifest-staging.xml "%DEPLOY_DIR%\"

REM Copy nginx configuration
echo Copying nginx configuration...
if not exist "%NGINX_DIR%\conf\conf.d" mkdir "%NGINX_DIR%\conf\conf.d"
copy deployment\nginx\excel-addin.conf "%NGINX_DIR%\conf\conf.d\"

REM Update nginx configuration with actual domain name
echo Updating nginx configuration...
powershell -Command "(Get-Content '%NGINX_DIR%\conf\conf.d\excel-addin.conf') -replace 'your-staging-domain.com', '%DOMAIN_NAME%' | Set-Content '%NGINX_DIR%\conf\conf.d\excel-addin.conf'"

REM Set up backend environment
echo Setting up backend environment...
cd /d "%DEPLOY_DIR%\backend"
copy .env.production .env

REM Update backend environment with actual domain
powershell -Command "(Get-Content '.env') -replace 'your-staging-domain.com', 'https://%DOMAIN_NAME%' | Set-Content '.env'"

REM Install Python dependencies (assuming Python is installed)
echo Installing Python dependencies...
pip install -r requirements.txt --user

REM Install NSSM for service management (download manually if not available)
echo Checking for NSSM...
nssm version >nul 2>&1
if %errorlevel% neq 0 (
    echo WARNING: NSSM not found. Please install NSSM from https://nssm.cc/
    echo Then rerun the service installation portion of this script.
    goto :skip_service
)

REM Install Windows Service using NSSM
echo Installing Windows Service...
nssm install %SERVICE_NAME% python "%DEPLOY_DIR%\backend\service_wrapper.py"
nssm set %SERVICE_NAME% AppDirectory "%DEPLOY_DIR%\backend"
nssm set %SERVICE_NAME% DisplayName "Excel Add-in Backend Service"
nssm set %SERVICE_NAME% Description "Python Flask backend service for Excel Add-in"
nssm set %SERVICE_NAME% Start SERVICE_AUTO_START
nssm set %SERVICE_NAME% AppStdout "%LOG_DIR%\service-stdout.log"
nssm set %SERVICE_NAME% AppStderr "%LOG_DIR%\service-stderr.log"

REM Set service environment variables
nssm set %SERVICE_NAME% AppEnvironmentExtra FLASK_ENV=production DEBUG=false HOST=127.0.0.1 PORT=5000 PYTHONPATH=%DEPLOY_DIR%\backend

:skip_service

REM Set file permissions
echo Setting file permissions...
icacls "%DEPLOY_DIR%" /grant "IIS_IUSRS:(OI)(CI)R" /T
icacls "%LOG_DIR%" /grant "IIS_IUSRS:(OI)(CI)F" /T

REM Create basic error pages
echo Creating error pages...
echo ^<html^>^<head^>^<title^>Page Not Found^</title^>^</head^>^<body^>^<h1^>404 - Page Not Found^</h1^>^</body^>^</html^> > "%DEPLOY_DIR%\dist\404.html"
echo ^<html^>^<head^>^<title^>Server Error^</title^>^</head^>^<body^>^<h1^>50x - Server Error^</h1^>^</body^>^</html^> > "%DEPLOY_DIR%\dist\50x.html"

REM Start services
echo Starting services...
net start %SERVICE_NAME%

REM Start nginx
echo Starting nginx...
cd /d "%NGINX_DIR%"
start nginx

REM Wait a moment for services to start
timeout /t 5 /nobreak >nul

REM Check service status
echo.
echo Checking service status...
sc query %SERVICE_NAME%

REM Check if nginx is running
tasklist /FI "IMAGENAME eq nginx.exe" 2>NUL | find /I /N "nginx.exe">NUL
if "%ERRORLEVEL%"=="0" (
    echo nginx is running
) else (
    echo WARNING: nginx is not running
)

echo.
echo ================================================================
echo Deployment Complete!
echo ================================================================
echo.
echo Next Steps:
echo 1. Update DNS to point %DOMAIN_NAME% to this server
echo 2. Install SSL certificate for %DOMAIN_NAME%
echo 3. Update nginx SSL certificate paths in %NGINX_DIR%\conf\conf.d\excel-addin.conf
echo 4. Test the application at https://%DOMAIN_NAME%
echo 5. Deploy the manifest file to Excel users
echo.
echo Logs are available at:
echo - Application: %LOG_DIR%\
echo - nginx: C:\Logs\nginx\
echo.
echo Service management:
echo - Start: net start %SERVICE_NAME%
echo - Stop: net stop %SERVICE_NAME%
echo - Status: sc query %SERVICE_NAME%
echo.

pause