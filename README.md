# PrimeExcelence Excel Add-in

A comprehensive Excel JavaScript add-in built with modern web technologies (TypeScript, React, Material UI) that provides a sophisticated sidebar interface for financial data management, applications, dashboards, and custom Excel functions.

## ✨ Features

- **📊 Task Pane Interface** - Modern React-based UI with Material-UI components
- **🧮 Custom Functions** - Two powerful custom functions available in Excel:
  - `PC.AGGIRR(expectedFutureValue, originalBeginningValue)` - Calculate aggregate IRR
  - `PC.JOINCELLS(range, delimiter)` - Join cell ranges with custom delimiters
- **🔧 Multi-Environment Support** - Configured for local development, staging, and production
- **🌐 Dynamic Configuration** - Automatically detects environment and configures API endpoints
- **🚀 Modern Build System** - Webpack-based with TypeScript, hot reloading, and optimized production builds

## Quick Start

**Prerequisites**: Ensure you have Poetry installed for backend dependency management:
```bash
# Install Poetry (if not already installed)
curl -sSL https://install.python-poetry.org | python3 -
# or on Windows:
(Invoke-WebRequest -Uri https://install.python-poetry.org -UseBasicParsing).Content | python -
```

### For Local Development
```bash
# 1. Install dependencies
npm install
cd backend && poetry install && poetry shell

# 2. Install HTTPS certificates for Office Add-ins
npm run cert:install

# 3. Start development server
npm run dev                    # Runs on https://localhost:3000

# 4. Start backend server (if using API features)
cd backend && poetry shell && python -m flask run --port=5000

# 5. Load in Excel
# Developer tab → Add-ins → Upload manifest.xml
```

### For Production/Staging Deployment

**NEW SERVICE-BASED ARCHITECTURE** - The deployment system has been completely redesigned:

```bash
# Deploy to staging/production server (Windows 10)
cd deployment
.\deploy-all.ps1       # Initial deployment (as Administrator)
.\update-all.ps1       # Updates after first deployment
.\test-deployment.ps1  # Comprehensive testing

# Individual service deployment
.\deploy-backend.ps1   # Deploy Python backend via NSSM
.\deploy-frontend.ps1  # Deploy React frontend via PM2
```

**Architecture**: IIS reverse proxy → Backend service (NSSM) + Frontend service (PM2)
**Public URL**: https://server-vs81t.intranet.local:9443

### Testing Custom Functions
Once loaded in Excel, try these examples:
```excel
=PC.AGGIRR(150, 100)           # Returns 1.5
=PC.JOINCELLS(A1:A5, " | ")    # Joins A1-A5 with " | "
```

## 🌍 Environment Support

The add-in automatically detects the environment based on the hostname and configures API endpoints dynamically:

| Environment | URL | API Endpoint | Manifest | Build Command |
|-------------|-----|--------------|----------|---------------|
| **Local Development** | https://localhost:3000 | http://localhost:5000/api | `manifest.xml` | `npm run build:dev` |
| **Staging** | https://server-vs81t.intranet.local:9443/excellence/ | https://server-vs81t.intranet.local:9443/excellence/api | `manifest-staging.xml` | `npm run build:staging` |
| **Production** | https://server-vs84.intranet.local:9443/excellence/ | https://server-vs84.intranet.local:9443/excellence/api | `manifest-prod.xml` | `npm run build:prod` |

### Dynamic Environment Detection
The application automatically detects its environment based on the browser hostname:
- **localhost/127.0.0.1** → Development (uses localhost:5000 backend for local development)
- **server-vs81t.intranet.local** → Staging 
- **server-vs84.intranet.local** → Production
- **Unknown hostnames** → Defaults to development with warnings

### For Staging/Production Deployment
```powershell
# Build for specific environment
npm run build:staging    # or npm run build:prod

# Deploy to IIS (run as Administrator)
.\deployment\scripts\build-and-deploy-iis.ps1
```

## Documentation

This project includes comprehensive documentation organized into three main guides:

📖 **[Application Guide](APPLICATION_GUIDE.md)** - Overview of the application, architecture, features, and how it works

🚀 **[Deployment Guide](DEPLOYMENT_GUIDE.md)** - Detailed deployment instructions for both local development and Windows Server production with step-by-step procedures

🔧 **[Troubleshooting Guide](TROUBLESHOOTING_GUIDE.md)** - Comprehensive troubleshooting tools, common issues, solutions, and diagnostic procedures

📜 **[Certificate Guide](CERTIFICATE_GUIDE.md)** - Complete guide for managing Excel Add-in SSL certificates in local development

## Key Features

- **Database Page**: KVG Data management with fund selection, data type filtering, and Excel integration
- **Applications Page**: Launch buttons for Kassandra, Infinity, and Pandora applications  
- **Dashboards Page**: Interactive Windpark A dashboard with multi-colored line charts
- **Excel Functions Page**: Collapsible descriptions of available Excel functions

## Architecture

### Frontend
- **React 19.1.1**: Modern UI framework
- **TypeScript 5.9.2**: Type-safe JavaScript
- **Material UI 7.3.1**: Google's Material Design components
- **Webpack 5.101.0**: Optimized bundle (~338 KB, 65% smaller than before)
- **Office.js**: Microsoft's JavaScript API for Office integration

### Backend
- **Python 3.x + Flask**: REST API service
- **Poetry**: Dependency management and virtual environment
- **IIS Integration**: Runs directly in IIS using FastCGI
- **Configuration**: Environment-based configuration

### Production Infrastructure  
- **IIS**: Web server hosting both frontend and backend
- **FastCGI**: Python integration for IIS
- **SSL/TLS**: Enterprise certificate support

## Deployment Scenarios

- **Local Development**: `https://localhost:3000` with self-signed certificates
- **Windows Server Production**: IIS hosting both frontend and backend with enterprise SSL certificates, subpath support (e.g., `/excellence`), and integrated Python FastCGI support

## Support

For detailed information, troubleshooting, and deployment procedures, please refer to the documentation guides above.

## License

This project is licensed under the ISC License.