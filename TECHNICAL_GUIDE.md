# PrimeExcelence Excel Add-in - Technical Overview & Installation Guide

## Architecture Overview

PrimeExcelence is a comprehensive Excel JavaScript add-in built with modern web technologies that provides a sophisticated sidebar interface for financial data management, applications, dashboards, and Excel functions.

### System Architecture

```
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│   Excel Client  │◄──►│  Frontend (Web)  │◄──►│ Backend (Flask) │
│                 │    │                  │    │                 │
│ - Task Pane     │    │ - React 19       │    │ - Python API    │
│ - Office.js API │    │ - TypeScript     │    │ - Data Sources  │
│ - Custom Funcs  │    │ - Material UI    │    │ - Business Logic│
└─────────────────┘    └──────────────────┘    └─────────────────┘
```

### Technology Stack

**Frontend**
- **React 19.1.1**: Modern UI framework with hooks and functional components
- **TypeScript 5.9.2**: Type-safe JavaScript with full IntelliSense support
- **Material UI 7.3.1**: Google's Material Design components for React
- **Vite**: Fast build tool and development server
- **Office.js**: Microsoft's JavaScript API for Office integration

**Backend**
- **Python 3.x**: Runtime environment
- **Flask**: Lightweight web framework for REST API
- **Poetry**: Modern dependency management and packaging tool
- **Configuration**: Environment-based configuration for different deployments

**Infrastructure**
- **Development**: HTTPS localhost with self-signed certificates
- **Production**: IIS reverse proxy with enterprise SSL certificates
- **Service Management**: NSSM (backend) + PM2/node-windows (frontend)

## Installation Guide

### Prerequisites

#### Development Environment
- **Node.js** (v16 or higher)
- **Python** (3.8 or higher) 
- **Poetry** (for Python dependency management)
- **Microsoft Excel** (Office 365, Excel 2016+)
- **PowerShell** (for Windows deployment scripts)

#### Production Environment (Windows Server)
- **Windows Server 2016+**
- **IIS** with FastCGI support
- **Node.js 18+**
- **Python 3.8+**
- **NSSM** (Non-Sucking Service Manager)
- **Enterprise SSL certificates**

### Local Development Setup

#### 1. Clone and Install Dependencies
```bash
# Clone repository
git clone <repository-url>
cd ExcelAddin

# Install frontend dependencies
npm install

# Install backend dependencies
cd backend
poetry install
poetry shell
```

#### 2. Certificate Setup for HTTPS
Excel Add-ins require HTTPS connections. For local development:

```bash
# Install Office Add-in development certificates
npm run cert:install

# Verify certificates are working
npm run cert:verify
```

**If you encounter certificate issues:**
- See the detailed [Certificate Troubleshooting](#certificate-troubleshooting) section below
- The certificates are automatically managed by the `office-addin-dev-certs` package
- Certificates are stored in `~/.office-addin-dev-certs/`

#### 3. Start Development Services
```bash
# Terminal 1 - Start backend server
cd backend
poetry shell
python run.py
# Backend runs on http://localhost:5000

# Terminal 2 - Start frontend development server  
npm run dev
# Frontend runs on https://localhost:3000
```

#### 4. Load Add-in in Excel
1. Open Microsoft Excel
2. Go to **Developer** tab → **Add-ins** → **My Add-ins**
3. Click **Upload My Add-in**
4. Select `manifest.xml` from the project root
5. The add-in sidebar should appear in Excel

### Production Deployment (Windows Server)

#### Quick Deployment
```powershell
# Run as Administrator
cd deployment
.\deploy-all.ps1 -Environment staging
```

#### Manual Step-by-Step Deployment

**1. Prepare Environment**
```powershell
# Build application for target environment
npm run build:staging    # or build:prod

# Ensure services are stopped
net stop "ExcelAddin-Backend" 2>$null
Stop-Service "ExcelAddin Frontend" -ErrorAction SilentlyContinue
```

**2. Deploy Backend Service**
```powershell
.\deploy-backend.ps1 -Environment staging
```

**3. Deploy Frontend Service**
```powershell
.\deploy-frontend.ps1 -Environment staging
```

**4. Configure IIS Reverse Proxy**
```powershell
.\deploy-iis.ps1
```

**5. Verify Deployment**
```powershell
# Test all services and connectivity
.\troubleshooting.ps1 -TestAll
```

## Environment Configuration

### Multi-Environment Support

The add-in automatically detects the environment and configures API endpoints:

| Environment | URL | API Endpoint | Manifest |
|-------------|-----|--------------|----------|
| **Development** | https://localhost:3000 | http://localhost:5000/api | `manifest.xml` |
| **Staging** | https://server-vs81t.intranet.local:9443/excellence/ | https://server-vs81t.intranet.local:9443/excellence/api | `manifest-staging.xml` |
| **Production** | https://server-vs84.intranet.local:9443/excellence/ | https://server-vs84.intranet.local:9443/excellence/api | `manifest-prod.xml` |

### Environment Variables

**Backend (.env files)**
```bash
# Development (.env.development)
DEBUG=true
CORS_ORIGINS=https://localhost:3000
DATABASE_URL=sqlite:///app.db

# Staging (.env.staging)  
DEBUG=false
CORS_ORIGINS=https://server-vs81t.intranet.local:9443
DATABASE_URL=postgresql://user:pass@db-server/staging

# Production (.env.production)
DEBUG=false
CORS_ORIGINS=https://server-vs84.intranet.local:9443  
DATABASE_URL=postgresql://user:pass@db-server/prod
```

## Application Features

### Core Functionality

**1. Database Page**
- KVG Data management with fund selection
- Data type filtering capabilities
- Direct Excel integration for data insertion

**2. Applications Page** 
- Launch buttons for Kassandra, Infinity, and Pandora applications
- External application integration

**3. Dashboards Page**
- Interactive Windpark A dashboard
- Multi-colored line charts and data visualization

**4. Excel Functions Page**
- Collapsible descriptions of available Excel functions
- Documentation for custom functions

### Custom Excel Functions

Two powerful custom functions available in Excel:

```excel
# Calculate aggregate IRR
=PC.AGGIRR(expectedFutureValue, originalBeginningValue)
=PC.AGGIRR(150, 100)  # Returns 1.5

# Join cell ranges with custom delimiters
=PC.JOINCELLS(range, delimiter)  
=PC.JOINCELLS(A1:A5, " | ")  # Joins A1-A5 with " | "
```

## Development Workflow

### Building for Different Environments

```bash
# Development build (with source maps)
npm run build:dev

# Staging build (optimized)
npm run build:staging  

# Production build (fully optimized)
npm run build:prod

# Development server with hot reload
npm run dev
```

### Testing and Validation

```bash
# Run frontend linting
npm run lint

# Run backend tests
cd backend
poetry run pytest

# Test certificate installation
npm run cert:verify
```

## Troubleshooting

### Certificate Troubleshooting

**Problem**: "The content is blocked because it isn't signed by a valid security certificate"

**Solution**:
```bash
# Quick fix - reinstall certificates
npm run cert:install

# If that fails, clean install
npm run cert:uninstall
npm run cert:install

# Verify installation
npm run cert:verify
```

**Manual Certificate Check (Windows)**:
1. Press `Win + R`, type `certlm.msc`, press Enter
2. Navigate to: **Trusted Root Certification Authorities** → **Certificates** 
3. Look for "office-addin-dev-ca" certificate
4. Check certificate expiration date

**Certificate Files Location**:
- Windows: `%USERPROFILE%\.office-addin-dev-certs\`
- Should contain: `localhost.crt`, `localhost.key`

### Common Development Issues

**1. Backend Server Won't Start**
```bash
# Check Python environment
cd backend
poetry shell
poetry install

# Check for port conflicts
netstat -an | findstr :5000

# Check backend configuration
python -c "from src.infrastructure.config.app_config import AppConfig; print(f'Environment: {AppConfig.ENVIRONMENT}')"
```

**2. Frontend Build Failures**
```bash
# Clear node modules and reinstall
rm -rf node_modules package-lock.json
npm install

# Check TypeScript compilation
npm run lint
```

**3. Excel Add-in Not Loading**
- Ensure HTTPS is working (check browser at https://localhost:3000)
- Verify manifest.xml is valid
- Clear Excel cache: Close Excel completely and restart
- Check Excel's Trust Center settings for add-ins

**4. API Connection Errors**
- Verify backend is running on correct port
- Check CORS configuration in backend
- Ensure firewall allows connections
- Test API endpoints directly: `curl http://localhost:5000/api/health`

### Production Troubleshooting

**Service Management**
```powershell
# Check service status
Get-Service "ExcelAddin-Backend"
Get-Service "ExcelAddin Frontend"

# View service logs
Get-EventLog -LogName Application -Source "ExcelAddin*" -Newest 10

# Restart services
Restart-Service "ExcelAddin-Backend"
Restart-Service "ExcelAddin Frontend"
```

**IIS Issues**
```powershell
# Check IIS application pools
Get-WebAppPoolState -Name "*ExcelAddin*"

# Check IIS logs
Get-ChildItem "C:\inetpub\logs\LogFiles\W3SVC1" | Sort-Object LastWriteTime -Descending | Select-Object -First 1 | Get-Content -Tail 20
```

**Connectivity Testing**
```powershell
# Use the comprehensive troubleshooting script
cd deployment
.\troubleshooting.ps1 -TestAll -Verbose

# Test specific components
.\troubleshooting.ps1 -TestServices
.\troubleshooting.ps1 -TestConnectivity  
.\troubleshooting.ps1 -FixCommonIssues
```

### Performance Optimization

**Bundle Size Optimization** (Already implemented):
- Vite build system: 346KB (down from 836KB with webpack)
- Tree shaking and code splitting enabled
- Image optimization and compression

**Runtime Performance**:
- Lazy loading for dashboard components
- Efficient Excel API usage patterns
- Minimal DOM manipulation
- Optimized chart rendering

## File Structure Reference

```
ExcelAddin/
├── src/                          # Frontend source code  
│   ├── components/              # React components
│   ├── pages/                   # Main application pages
│   ├── commands/                # Excel custom functions  
│   └── taskpane/               # Task pane entry point
├── backend/                     # Python backend
│   ├── src/                    # Python source (DDD structure)
│   ├── app.py                  # Flask application factory
│   ├── wsgi_app.py            # IIS WSGI entry point
│   └── run.py                  # Development server
├── deployment/                  # Deployment scripts
│   ├── deploy-backend.ps1      # Backend service deployment
│   ├── deploy-frontend.ps1     # Frontend service deployment  
│   ├── deploy-iis.ps1         # IIS proxy configuration
│   └── troubleshooting.ps1     # Diagnostic and debug tools
├── public/                      # Static assets
│   └── assets/                 # Images and icons
├── dist/                       # Built frontend files (generated)
├── manifest*.xml               # Excel add-in manifests
├── package.json                # npm dependencies and scripts
├── vite.config.ts             # Vite configuration
└── tsconfig.json               # TypeScript configuration
```

## Security Considerations

### HTTPS Requirements
- All Office add-ins must use HTTPS in production
- Excel validates SSL certificates and domain trust
- Self-signed certificates only work for localhost development

### CORS Configuration  
- Backend must allow requests from Excel domains
- Production CORS should be restrictive to actual deployment domains
- Development CORS can be more permissive

### Content Security Policy
- Ensure CSP headers don't block Office.js
- Allow inline styles/scripts required by Office framework
- Configure proper script-src and connect-src directives

This technical guide provides comprehensive coverage of the PrimeExcelence Excel add-in architecture, installation, and troubleshooting procedures. For deployment-specific procedures, refer to the [Deployment Guide](deployment/README.md).