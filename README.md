# PrimeExcelence Excel Add-in

A comprehensive Excel JavaScript add-in built with modern web technologies (TypeScript, React, Material UI) that provides a sophisticated sidebar interface for financial data management, applications, dashboards, and custom Excel functions.

## ✨ Features

- **📊 Task Pane Interface** - Modern React-based UI with Material-UI components
- **🧮 Custom Functions** - Two powerful custom functions available in Excel:
  - `PC.AGGIRR(expectedFutureValue, originalBeginningValue)` - Calculate aggregate IRR
  - `PC.JOINCELLS(range, delimiter)` - Join cell ranges with custom delimiters
- **🔧 Multi-Environment Support** - Configured for local development, staging, and production
- **🌐 Dynamic Configuration** - Automatically detects environment and configures API endpoints
- **🚀 Modern Build System** - Vite-based with TypeScript, hot reloading, and optimized production builds

## Quick Start

### For Local Development
```bash
# 1. Install dependencies
npm install
cd backend && poetry install && poetry shell

# 2. Install HTTPS certificates for Office Add-ins
npm run cert:install

# 3. Start development servers
npm run dev                    # Frontend on https://localhost:3000
cd backend && python run.py   # Backend on http://localhost:5000

# 4. Load in Excel
# Developer tab → Add-ins → Upload manifest.xml
```

### For Production Deployment
```powershell
# Deploy to Windows Server (run as Administrator)
cd deployment
.\deploy-backend.ps1 -Environment staging
.\deploy-frontend.ps1 -Environment staging  
.\deploy-iis.ps1

# Verify deployment
.\troubleshooting.ps1 -TestAll
```

### Testing Custom Functions
Once loaded in Excel, try these examples:
```excel
=PC.AGGIRR(150, 100)           # Returns 1.5
=PC.JOINCELLS(A1:A5, " | ")    # Joins A1-A5 with " | "
```

## 🌍 Environment Support

The add-in automatically detects the environment based on the hostname and configures API endpoints dynamically:

| Environment | URL | API Endpoint | Manifest |
|-------------|-----|--------------|----------|
| **Local Development** | https://localhost:3000 | http://localhost:5000/api | `manifest.xml` |
| **Staging** | https://server-vs81t.intranet.local:9443/excellence/ | https://server-vs81t.intranet.local:9443/excellence/api | `manifest-staging.xml` |
| **Production** | https://server-vs84.intranet.local:9443/excellence/ | https://server-vs84.intranet.local:9443/excellence/api | `manifest-prod.xml` |

## 📚 Documentation

This project includes comprehensive documentation organized into four main guides:

📖 **[Technical Guide](TECHNICAL_GUIDE.md)** - Detailed technical overview, installation procedures, architecture documentation, and troubleshooting

🚀 **[Deployment Guide](deployment/README.md)** - Production deployment procedures for Windows Server environments with service management and configuration

🖼️ **[Asset Serving Guide](ASSET_SERVING.md)** - Environment-specific asset serving and conditional image references

🛠️ **Troubleshooting** - Use `.\deployment\troubleshooting.ps1` for comprehensive diagnostics and automated fixes

## Key Application Features

- **Database Page**: KVG Data management with fund selection, data type filtering, and Excel integration
- **Applications Page**: Launch buttons for Kassandra, Infinity, and Pandora applications  
- **Dashboards Page**: Interactive Windpark A dashboard with multi-colored line charts
- **Excel Functions Page**: Collapsible descriptions of available Excel functions

## Architecture Overview

### Frontend
- **React 19.1.1**: Modern UI framework
- **TypeScript 5.9.2**: Type-safe JavaScript
- **Material UI 7.3.1**: Google's Material Design components
- **Vite**: Fast build tool (~346KB optimized bundle, 77% faster builds)
- **Office.js**: Microsoft's JavaScript API for Office integration

### Backend
- **Python 3.x + Flask**: REST API service
- **Poetry**: Dependency management and virtual environment
- **Configuration**: Environment-based configuration
- **Service Architecture**: NSSM Windows service for production

### Production Infrastructure  
- **IIS Reverse Proxy**: HTTPS termination and routing on port 9443
- **Frontend Service**: Express.js via PM2/node-windows Windows service
- **Backend Service**: Python Flask via NSSM Windows service
- **SSL/TLS**: Enterprise certificate support

## Deployment Scenarios

- **Local Development**: `https://localhost:3000` with self-signed certificates
- **Production**: Windows Server with IIS reverse proxy, enterprise SSL certificates, service-based architecture, and subpath support (e.g., `/excellence`)

## Image Serving Architecture

**Current Structure Analysis**:
The application uses a distributed image serving approach:

```
Images Location:
├── public/assets/          # Frontend build-time assets (served by frontend service)  
│   ├── PCAG_*.png         # Company logos and branding
│   └── icon-*.png         # Excel add-in icons
├── assets/                # Legacy duplicate assets (should be consolidated)
│   └── PCAG_*.png         
└── backend/static/        # Backend static files (if needed)
```

**Recommended Structure for Better Image Serving**:
```
Consolidated Assets:
├── public/assets/images/   # All images in organized structure
│   ├── icons/             # Excel add-in icons
│   ├── logos/             # Company branding  
│   └── ui/                # UI elements and graphics
├── backend/               # Backend serves API only, no static files
└── dist/assets/images/    # Built and optimized images
```

**Image Serving Issues & Solutions**:
1. **Duplicate Assets**: Images exist in both `/public/assets/` and `/assets/` - consolidate to `/public/assets/images/`
2. **Backend Static Serving**: Backend should not serve frontend assets - remove any static file routes from Flask app
3. **IIS Routing**: Ensure IIS proxy correctly routes image requests to frontend service, not backend
4. **Build Optimization**: Vite should handle image optimization and proper URL generation

## Quick Fix Commands

**Certificate Issues**:
```bash
npm run cert:install    # Install certificates
npm run cert:verify     # Check certificate status
```

**Service Issues**:
```powershell
# Comprehensive diagnostics and fixes
.\deployment\troubleshooting.ps1 -TestAll -FixCommonIssues

# Individual service management
Get-Service "ExcelAddin*" | Restart-Service
```

**Build Issues**:
```bash
# Clean build
npm run clean && npm install
npm run build:staging  # or build:prod
```

For detailed information, troubleshooting, and deployment procedures, please refer to the **Technical Guide** and **Deployment Guide** linked above.