# PrimeExcelence Excel Add-in - Application Guide

## Overview

PrimeExcelence is a comprehensive Excel JavaScript add-in built with modern web technologies (TypeScript, React, Material UI) that provides a sophisticated sidebar interface for financial data management, applications, dashboards, and Excel functions.

The add-in consists of two main components:
1. **Frontend (React App)**: Modern web interface that runs in Excel's task pane
2. **Backend (Python Flask API)**: REST API service providing data and business logic

## Architecture

```
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│   Excel Client  │◄──►│  Frontend (Web)  │◄──►│ Backend (Flask) │
│                 │    │                  │    │                 │
│ - Task Pane     │    │ - React 19       │    │ - Python API    │
│ - Office.js API │    │ - TypeScript     │    │ - Data Sources  │
│ - Custom Funcs  │    │ - Material UI    │    │ - Business Logic│
└─────────────────┘    └──────────────────┘    └─────────────────┘
```

### Frontend Features

- **Database Page**: KVG Data management with fund selection, data type filtering, and Excel integration
- **Applications Page**: Launch buttons for Kassandra, Infinity, and Pandora applications  
- **Dashboards Page**: Interactive Windpark A dashboard with multi-colored line charts
- **Excel Functions Page**: Collapsible descriptions of available Excel functions

### Backend Features

- **RESTful API**: Provides data endpoints for frontend consumption
- **Database Integration**: Connects to various data sources for financial information
- **Authentication**: Handles user authentication and session management
- **Business Logic**: Implements financial calculations and data processing
- **Poetry-based Dependency Management**: Uses Poetry for reliable dependency and environment management

## Technology Stack

### Frontend
- **React 19.1.1**: Modern UI framework with hooks and functional components
- **TypeScript 5.9.2**: Type-safe JavaScript with full IntelliSense support
- **Material UI 7.3.1**: Google's Material Design components for React
- **Webpack 5.101.0**: Module bundler with optimized production builds
- **Office.js**: Microsoft's JavaScript API for Office integration

### Backend
- **Python 3.x**: Runtime environment
- **Flask**: Lightweight web framework for REST API
- **Poetry**: Modern dependency management and packaging tool
- **Configuration**: Environment-based configuration for different deployments

### Build & Development Tools
- **npm/Node.js**: Package management and build tooling
- **Custom Functions Metadata**: Excel custom functions generation
- **Source Maps**: Development debugging support

## Deployment Scenarios

### Local Development
- **Frontend**: Development server on `https://localhost:3000`
- **Backend**: Python development server on `http://localhost:5000`
- **SSL**: Self-signed certificates for HTTPS requirement
- **Excel Integration**: Sideloaded manifest for testing

### Production/Staging - Windows Server
- **Frontend**: Static files served by IIS
- **Backend**: Python Flask app running directly in IIS via FastCGI
- **Integration**: IIS handles both frontend and backend with unified configuration
- **SSL**: Enterprise certificates from company CA
- **Service Management**: IIS application pool management for both components

## File Structure

```
ExcelAddin/
├── src/                          # Frontend source code
│   ├── components/              # React components
│   ├── pages/                   # Main application pages
│   ├── commands/                # Excel custom functions
│   └── taskpane/               # Task pane entry point
├── backend/                     # Python backend
│   ├── src/                    # Python source code
│   ├── wsgi_app.py            # IIS WSGI entry point
│   ├── web.config             # IIS configuration for backend
│   └── run.py                  # Development server
├── deployment/                  # Deployment configuration
│   ├── iis/                    # IIS configuration files
│   ├── scripts/                # PowerShell deployment scripts
│   └── ssl/                    # SSL certificate templates
├── dist/                       # Built frontend files (generated)
├── manifest*.xml               # Excel add-in manifests
├── package.json                # npm dependencies and scripts
├── webpack.config.js           # Development build config
├── webpack.prod.config.js      # Production build config
└── tsconfig.json               # TypeScript configuration
```

## Getting Started

### Prerequisites
- **Node.js** (v16 or higher)
- **Python** (3.8 or higher)
- **Poetry** (for Python dependency management)
- **Microsoft Excel** (Office 365, Excel 2016+)
- **PowerShell** (for Windows deployment scripts)

### Quick Start - Local Development

1. **Clone and setup**:
   ```bash
   git clone <repository-url>
   cd ExcelAddin
   npm install
   cd backend
   poetry install
   ```

2. **Start services**:
   ```bash
   # Terminal 1 - Backend
   cd backend
   poetry shell
   python run.py
   
   # Terminal 2 - Frontend  
   npm start
   ```

3. **Load in Excel**:
   - Open Excel
   - Go to Developer tab → Add-ins → My Add-ins
   - Upload `manifest.xml`
   - Click "Show Taskpane" button

### Quick Start - Windows Server Production

1. **Prepare environment**:
   - Install IIS with FastCGI support, Python
   - Configure SSL certificates
   - Build application (`npm run build:staging`)

2. **Deploy using scripts**:
   ```powershell
   # Run as Administrator
   .\deployment\scripts\setup-backend-iis.ps1
   .\deployment\scripts\setup-iis.ps1
   ```

3. **Configure Excel**:
   - Deploy `manifest-staging.xml` to users
   - Configure trust settings for your domain

## Key Concepts

### Excel Integration
- **Office.js API**: Used for reading/writing Excel data, creating custom functions
- **Task Pane**: Sidebar interface that hosts the React application
- **Custom Functions**: User-defined functions available in Excel formulas
- **Ribbon Commands**: Custom buttons in Excel's ribbon interface

### Security Requirements
- **HTTPS Required**: All Office add-ins must use HTTPS in production
- **Domain Trust**: Excel validates SSL certificates and domain trust
- **CORS Configuration**: Backend must allow requests from Excel domains

### Performance Optimizations
- **Lazy Loading**: Pages load on-demand to reduce initial bundle size
- **Bundle Splitting**: Webpack separates vendor and application code
- **Optimized Charts**: Custom SVG charts instead of heavy charting libraries
- **Service Workers**: Caching for improved load times (when applicable)

This application guide provides the foundation for understanding the PrimeExcelence Excel add-in. For detailed deployment instructions, see the [Deployment Guide](DEPLOYMENT_GUIDE.md), and for troubleshooting assistance, see the [Troubleshooting Guide](TROUBLESHOOTING_GUIDE.md).