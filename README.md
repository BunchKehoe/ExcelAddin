# PrimeExcelence Excel Add-in

A comprehensive Excel JavaScript add-in built with modern web technologies (TypeScript, React, Material UI) that provides a sophisticated sidebar interface for financial data management, applications, dashboards, and Excel functions.

## Quick Start

### For Local Development
```bash
# 1. Install dependencies
npm install
cd backend && pip install -r requirements.txt

# 2. Start services
cd backend && python run.py  # Terminal 1
npm start                     # Terminal 2

# 3. Load in Excel
# Developer tab → Add-ins → Upload manifest.xml
```

### For Windows Server Production
```powershell
# 1. Build and deploy files
npm run build:staging
# Copy dist/ and backend/ to server

# 2. Run setup scripts (as Administrator)
.\deployment\scripts\setup-backend-service.ps1
.\deployment\scripts\setup-nginx-service.ps1

# 3. Start services
Start-Service ExcelAddinBackend
Start-Service nginx
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