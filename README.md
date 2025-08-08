# PrimeExcelence Excel Add-in

A comprehensive Excel JavaScript add-in built with modern web technologies (TypeScript, React, Material UI) that provides a sophisticated sidebar interface for financial data management, applications, dashboards, and custom Excel functions.

## ✨ Features

- **📊 Task Pane Interface** - Modern React-based UI with Material-UI components
- **🧮 Custom Functions** - Two powerful custom functions available in Excel:
  - `PC.AGGIRR(expectedFutureValue, originalBeginningValue)` - Calculate aggregate IRR
  - `PC.JOINCELLS(range, delimiter)` - Join cell ranges with custom delimiters
- **🔧 Multi-Environment Support** - Configured for local development, staging, and production
- **🚀 Modern Build System** - Webpack-based with TypeScript, hot reloading, and optimized production builds

## Quick Start

### For Local Development
```bash
# 1. Install dependencies
npm install
cd backend && pip install -r requirements.txt

# 2. Start development server
npm run dev                    # Runs on https://localhost:3000

# 3. Load in Excel
# Developer tab → Add-ins → Upload manifest.xml
```

### Testing Custom Functions
Once loaded in Excel, try these examples:
```excel
=PC.AGGIRR(150, 100)           # Returns 1.5
=PC.JOINCELLS(A1:A5, " | ")    # Joins A1-A5 with " | "
```

## 🌍 Environment Support

| Environment | URL | Manifest | Build Command |
|-------------|-----|----------|---------------|
| **Local Development** | https://localhost:3000 | `manifest.xml` | `npm run build:dev` |
| **Staging** | https://server-vs81t.intranet.local:9443/excellence/ | `manifest-staging.xml` | `npm run build:staging` |
| **Production** | https://server-vs84.intranet.local:9443/excellence/ | `manifest-prod.xml` | `npm run build:prod` |

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
- **Windows Service**: Runs as Windows service via NSSM
- **Configuration**: Environment-based configuration

### Production Infrastructure  
- **nginx**: Reverse proxy with SSL termination
- **NSSM**: Service management for Windows
- **SSL/TLS**: Enterprise certificate support

## Deployment Scenarios

- **Local Development**: `https://localhost:3000` with self-signed certificates
- **Windows Server Production**: nginx reverse proxy with enterprise SSL certificates, subpath support (e.g., `/excellence`), and Windows service management

## Support

For detailed information, troubleshooting, and deployment procedures, please refer to the documentation guides above.

## License

This project is licensed under the ISC License.