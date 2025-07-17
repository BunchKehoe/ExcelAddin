# PrimeExcelence Excel JS Addin

A comprehensive Excel JavaScript addin built with TypeScript, React, and Material UI, providing a modern sidebar interface for financial data management, applications, dashboards, and Excel functions.

## Table of Contents

- [Prerequisites](#prerequisites)
- [Initial Setup](#initial-setup)
- [Development Environment](#development-environment)
- [Local Testing with Excel](#local-testing-with-excel)
- [Production Deployment](#production-deployment)
- [SSL Certificates](#ssl-certificates)
- [Manifest Configuration](#manifest-configuration)
- [Troubleshooting](#troubleshooting)
- [Features](#features)

## Prerequisites

Before starting, ensure you have the following installed:

- **Node.js** (v16 or higher) - [Download from nodejs.org](https://nodejs.org/)
- **npm** or **yarn** package manager
- **Microsoft Excel** (Office 365, Excel 2016, or later)
- **Web browser** (Chrome, Firefox, or Edge for development)
- **Code editor** (VS Code recommended)

## Initial Setup

1. **Clone the repository**
   ```bash
   git clone <repository-url>
   cd ExcelAddin
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Verify the installation**
   ```bash
   npm run build
   ```

## Development Environment

### Starting the Development Server

1. **Start the development server**
   ```bash
   npm start
   # or
   npm run dev
   ```

2. **Access the application**
   - Open your browser and navigate to `https://localhost:3000`
   - The development server runs on port 3000 by default

### Building for Production

```bash
npm run build
```

This creates a `dist` folder with the production-ready files.

## Local Testing with Excel

### Method 1: Sideloading via Excel UI (Recommended for Development)

1. **Open Excel** (Excel 365, Excel 2016, or later)

2. **Enable Developer Mode** (if not already enabled)
   - Go to File → Options → Customize Ribbon
   - Check "Developer" in the right panel
   - Click OK

3. **Sideload the manifest**
   - Go to the Developer tab in Excel
   - Click "Add-ins" → "My Add-ins"
   - Click "Upload My Add-in"
   - Browse and select the `manifest.xml` file from the project root
   - Click "Upload"

4. **Launch the Addin**
   - The PrimeExcelence button should appear in the Home tab
   - Click "Show Taskpane" to open the sidebar

### Method 2: Sideloading via Office Developer Tools

1. **Install Office-Addin-dev-certs** (for HTTPS support)
   ```bash
   npm install -g office-addin-dev-certs
   office-addin-dev-certs install
   ```

2. **Use Yeoman Office Generator** (optional, for advanced development)
   ```bash
   npm install -g yo generator-office
   ```

### Method 3: Network Shared Folder (for Teams)

1. **Share the manifest file** on a network location
2. **Add the shared folder** as a trusted catalog in Excel:
   - File → Options → Trust Center → Trust Center Settings
   - Trusted Add-in Catalogs
   - Add the network path containing the manifest.xml
   - Check "Show in Menu"

## Production Deployment

### Hosting Requirements

For production deployment, you need:

1. **HTTPS-enabled web server** (required by Office)
2. **Valid SSL certificate** (not self-signed)
3. **Accessible domain** (not localhost)

### Deployment Steps

1. **Build the production version**
   ```bash
   npm run build
   ```

2. **Update the manifest.xml for production**
   - Replace all instances of `https://localhost:3000` with your production URL
   - Update the `<Id>` to a unique GUID for your organization
   - Update `<ProviderName>` to your organization name

3. **Deploy to your web server**
   - Upload the `dist` folder contents to your web server
   - Upload the `manifest.xml` file to an accessible location
   - Ensure all files are served over HTTPS

4. **Distribution Options**
   - **AppSource**: Submit to Microsoft AppSource for public distribution
   - **SharePoint Catalog**: Deploy to SharePoint App Catalog for organization-wide distribution
   - **Network Share**: Use a network shared folder for internal distribution

## SSL Certificates

### Development Environment

For local development, you have several options:

#### Option 1: Office-Addin-dev-certs (Recommended)

```bash
# Install the certificate utility
npm install -g office-addin-dev-certs

# Install the development certificate
office-addin-dev-certs install

# Verify installation
office-addin-dev-certs verify
```

This creates a self-signed certificate that Excel will trust for development.

#### Option 2: Manual Certificate Creation

1. **Create a self-signed certificate**
   ```bash
   # Using OpenSSL
   openssl req -x509 -newkey rsa:4096 -keyout key.pem -out cert.pem -days 365 -nodes
   ```

2. **Configure webpack for HTTPS**
   Update `webpack.config.js`:
   ```javascript
   devServer: {
     https: {
       key: fs.readFileSync('key.pem'),
       cert: fs.readFileSync('cert.pem')
     },
     // ... other settings
   }
   ```

3. **Trust the certificate** in your browser and system

#### Option 3: Using mkcert

```bash
# Install mkcert
npm install -g mkcert

# Create and install CA
mkcert -install

# Create certificate for localhost
mkcert localhost
```

### Production Environment

For production, you need a valid SSL certificate from a trusted Certificate Authority:

1. **Commercial CA**: GoDaddy, DigiCert, Comodo, etc.
2. **Free CA**: Let's Encrypt (recommended for most cases)
3. **Internal CA**: For enterprise environments

#### Let's Encrypt Setup (Example)

```bash
# Install Certbot
sudo apt-get install certbot

# Generate certificate
sudo certbot certonly --standalone -d yourdomain.com

# Certificate files will be in:
# /etc/letsencrypt/live/yourdomain.com/
```

## Manifest Configuration

### Development Manifest (manifest.xml)

```xml
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xsi:type="TaskPaneApp">
  <Id>12345678-1234-1234-1234-123456789012</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Your Organization</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="PrimeExcelence"/>
  <Description DefaultValue="Excel JS Addin for PrimeExcelence"/>
  <IconUrl DefaultValue="https://localhost:3000/assets/icon-32.png"/>
  <SupportUrl DefaultValue="https://yourdomain.com/support"/>
  <AppDomains>
    <AppDomain>https://localhost:3000</AppDomain>
  </AppDomains>
  <!-- ... rest of manifest ... -->
</OfficeApp>
```

### Production Manifest

Create a separate `manifest-prod.xml` for production:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xsi:type="TaskPaneApp">
  <Id>your-unique-guid-here</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Your Organization</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="PrimeExcelence"/>
  <Description DefaultValue="Excel JS Addin for PrimeExcelence"/>
  <IconUrl DefaultValue="https://yourdomain.com/assets/icon-32.png"/>
  <SupportUrl DefaultValue="https://yourdomain.com/support"/>
  <AppDomains>
    <AppDomain>https://yourdomain.com</AppDomain>
  </AppDomains>
  <!-- ... rest of manifest ... -->
</OfficeApp>
```

### Key Manifest Properties

- **Id**: Unique GUID for your addin (generate at https://guidgenerator.com/)
- **Version**: Semantic version (increment for updates)
- **ProviderName**: Your organization name
- **IconUrl**: Must be accessible via HTTPS
- **SourceLocation**: Main application entry point
- **AppDomains**: All domains your addin will use

## Troubleshooting

### Common Issues

1. **"Add-in won't load"**
   - Ensure the development server is running (`npm start`)
   - Check that the manifest URLs are accessible
   - Verify SSL certificates are trusted

2. **"Manifest validation errors"**
   - Use the Office Add-in Validator: https://github.com/OfficeDev/office-addin-validator
   - Check XML syntax and required elements

3. **"HTTPS required" error**
   - Install proper SSL certificates
   - Ensure all resources are served over HTTPS

4. **"CORS errors"**
   - Configure proper CORS headers on your server
   - Ensure AppDomains in manifest match your server domain

5. **"Office.js not loaded"**
   - Ensure Office.js is properly included in your HTML
   - Check that Office.onReady() is called

### Debugging Tips

1. **Use Office Developer Tools**
   - Press F12 in Excel to open developer tools
   - Check console for errors

2. **Enable logging**
   - Add console.log statements in your code
   - Use Office.js runtime logging

3. **Test in different environments**
   - Test in Excel Online and Excel desktop
   - Test with different Excel versions

### Validation Commands

```bash
# Validate manifest
npx office-addin-validator manifest.xml

# Check HTTPS configuration
curl -I https://localhost:3000

# Test Office.js loading
curl https://localhost:3000/taskpane.html
```

## Features

### Core Functionality

- **Database Page**: KVG Data with fund selection, data type filtering, and Excel integration
- **Applications Page**: Launch buttons for Kassandra, Infinity, and Pandora applications
- **Dashboards Page**: Interactive Windpark A dashboard with multi-colored line charts
- **Excel Functions Page**: Collapsible descriptions of available Excel functions

### Technical Stack

- **Frontend**: React 19, TypeScript, Material UI
- **Build System**: Webpack 5 with development and production configurations
- **Excel Integration**: Office.js API for reading/writing Excel data
- **Charts**: Recharts for dashboard visualizations
- **HTTP Client**: Axios for API communication

### Browser Support

- Microsoft Excel (Office 365, 2016+)
- Excel Online
- Modern browsers (Chrome, Firefox, Edge, Safari)

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly in Excel
5. Submit a pull request

## License

This project is licensed under the ISC License.

## Support

For technical support or questions:
- Create an issue in the GitHub repository
- Contact the development team
- Check the troubleshooting section above

---

**Note**: This addin requires Excel 2016 or later, or Office 365. For older versions of Excel, consider using VSTO or COM add-ins instead.