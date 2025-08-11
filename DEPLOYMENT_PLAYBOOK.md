# Excel Add-in Deployment Playbook

This playbook covers deployment procedures for all environments: Local Development, Staging, and Production.

## Environment Overview

| Environment | URL | Manifest | Build Command |
|-------------|-----|----------|---------------|
| **Local Development** | https://localhost:3000 | manifest.xml | `npm run build:dev` |
| **Staging** | https://server-vs81t.intranet.local:9443/excellence/ | manifest-staging.xml | `npm run build:staging` |
| **Production** | https://server-vs84.intranet.local:9443/excellence/ | manifest-prod.xml | `npm run build:prod` |

## Features Included

All environments support:
- ✅ Task Pane functionality
- ✅ Custom Functions: `PC.AGGIRR()` and `PC.JOINCELLS()`
- ✅ React-based UI with Material-UI components
- ✅ Backend API integration
- ✅ Asset serving (icons, etc.)

---

## 🖥️ Local Development Deployment

### Prerequisites
- Node.js 16+ installed
- Git
- Office Add-in SSL certificates (required for local development)

### Setup & Run

#### Step 1: Clone and Install Dependencies
```bash
# 1. Clone and install
git clone <repository>
cd ExcelAddin
npm install
```

#### Step 2: Generate Office Add-in SSL Certificates
Office Add-ins require HTTPS connections, even for local development. Generate the required SSL certificates:

```bash
# Install certificates (creates ~/.office-addin-dev-certs/ directory)
npm run cert:install

# Verify certificates are properly installed
npm run cert:verify
```

**Certificate Details:**
- **Location:** `~/.office-addin-dev-certs/`
- **Files Created:**
  - `localhost.crt` - SSL certificate for localhost
  - `localhost.key` - Private key for the certificate
- **Validity:** These certificates are trusted by your system for Office Add-in development

**Troubleshooting Certificate Issues:**
If you encounter certificate problems:
```bash
# Uninstall existing certificates
npm run cert:uninstall

# Reinstall fresh certificates
npm run cert:install

# Verify installation
npm run cert:verify
```

#### Step 3: Start Development Server
```bash
# Start development server with HTTPS
npm run dev
# or
npm start
```

The development server will automatically use the SSL certificates from `~/.office-addin-dev-certs/` for secure HTTPS connections required by Office Add-ins.

### Development URLs
- **Frontend:** https://localhost:3000
- **Task Pane:** https://localhost:3000/taskpane.html
- **Commands:** https://localhost:3000/commands.html
- **Functions Metadata:** https://localhost:3000/functions.json

### Testing Custom Functions Locally
1. Open Excel
2. Go to Insert > My Add-ins > Upload My Add-in
3. Select `manifest.xml` from the project root
4. Use functions in Excel: `=PC.AGGIRR(100, 50)` or `=PC.JOINCELLS(A1:A3, "; ")`

### Local Development Checklist
- [ ] `npm install` completed successfully
- [ ] SSL certificates installed (`npm run cert:verify`)
- [ ] Development server starts without errors (`npm run dev`)
- [ ] All pages load: taskpane.html, commands.html
- [ ] functions.json is accessible
- [ ] Custom functions work in Excel
- [ ] Backend API responds (if applicable)

---

## 🧪 Staging Environment Deployment

### Prerequisites
- Access to staging server: `server-vs81t.intranet.local`
- IIS with SSL configured on port 9443
- PowerShell execution privileges
- SSL certificates properly configured on staging server

**Important:** Staging builds use the production webpack configuration (`webpack.prod.config.js`) which does not require Office Add-in development certificates. The staging server handles SSL through IIS configuration.

### Build & Deploy
```bash
# 1. Prepare staging build
npm run clean
npm install
npm run build:staging
```

### Deployment to IIS (Staging)
```powershell
# Run as Administrator on staging server
.\deployment\scripts\deploy-to-existing-iis.ps1

# Build and deploy application files  
.\deployment\scripts\build-and-deploy-iis.ps1

# Test the deployment
.\deployment\scripts\test-iis-simple.ps1
```

### Staging URLs
- **Frontend:** https://server-vs81t.intranet.local:9443/excellence/
- **Task Pane:** https://server-vs81t.intranet.local:9443/excellence/taskpane.html
- **Commands:** https://server-vs81t.intranet.local:9443/excellence/commands.html
- **Functions Metadata:** https://server-vs81t.intranet.local:9443/excellence/functions.json
- **Manifest:** https://server-vs81t.intranet.local:9443/excellence/manifest.xml

### Testing Staging Custom Functions
1. Download `manifest-staging.xml` or copy from build output `dist/manifest.xml`
2. Open Excel
3. Go to Insert > My Add-ins > Upload My Add-in
4. Select the staging manifest file
5. Test functions: `=PC.AGGIRR(200, 100)` or `=PC.JOINCELLS(B1:B5, " | ")`

### Staging Deployment Checklist
- [ ] Build completed: `npm run build:staging`
- [ ] Files copied to `C:\inetpub\wwwroot\ExcelAddin\excellence\`
- [ ] IIS site "ExcelAddin" is running
- [ ] HTTPS binding on port 9443 works
- [ ] Frontend loads: https://server-vs81t.intranet.local:9443/excellence/
- [ ] manifest.xml is accessible and valid
- [ ] functions.json is accessible
- [ ] Custom functions work in Excel
- [ ] Backend API integration works
- [ ] All assets (icons) load properly

---

## 🚀 Production Environment Deployment

### Prerequisites
- Access to production server: `server-vs84.intranet.local`
- IIS with SSL configured on port 9443
- PowerShell execution privileges
- SSL certificates in `C:\Cert\server-vs84.(crt|key)`
- Change management approval (if required)

### Build & Deploy
```bash
# 1. Prepare production build
npm run clean
npm install
npm run build:prod
```

### Pre-Deployment Checklist
- [ ] Code review completed
- [ ] Testing completed in staging
- [ ] Change management approved
- [ ] Production server access confirmed
- [ ] SSL certificates verified
- [ ] Backup of current production taken

### Deployment to IIS (Production)
```powershell
# Run as Administrator on production server

# 1. Create IIS site (if first deployment)
.\deployment\scripts\deploy-to-existing-iis.ps1

# 2. Deploy application files
.\deployment\scripts\build-and-deploy-iis.ps1

# 3. Test deployment
.\deployment\scripts\test-iis-simple.ps1

# 4. Health check
.\deployment\scripts\health-check.ps1
```

### Production URLs
- **Frontend:** https://server-vs84.intranet.local:9443/excellence/
- **Task Pane:** https://server-vs84.intranet.local:9443/excellence/taskpane.html
- **Commands:** https://server-vs84.intranet.local:9443/excellence/commands.html
- **Functions Metadata:** https://server-vs84.intranet.local:9443/excellence/functions.json
- **Manifest:** https://server-vs84.intranet.local:9443/excellence/manifest.xml

### Testing Production Custom Functions
1. Download `manifest-prod.xml` or copy from build output `dist/manifest.xml`
2. Open Excel
3. Go to Insert > My Add-ins > Upload My Add-in
4. Select the production manifest file
5. Test functions: `=PC.AGGIRR(500, 250)` or `=PC.JOINCELLS(C1:C10, "; ")`

### Production Deployment Checklist
- [ ] Build completed: `npm run build:prod`
- [ ] Files copied to `C:\inetpub\wwwroot\ExcelAddin\excellence\`
- [ ] IIS site "ExcelAddin" is running
- [ ] HTTPS binding on port 9443 works
- [ ] Frontend loads: https://server-vs84.intranet.local:9443/excellence/
- [ ] manifest.xml is accessible and valid
- [ ] functions.json is accessible
- [ ] Custom functions work in Excel
- [ ] Backend API integration works
- [ ] All assets (icons) load properly
- [ ] Health check passes
- [ ] User acceptance testing completed

### Post-Deployment Verification
- [ ] End-user testing completed
- [ ] Performance monitoring active
- [ ] Logs reviewed for errors
- [ ] Rollback plan ready (if issues found)

---

## 🛠️ Custom Functions Guide

### Available Functions
1. **PC.AGGIRR(expectedFutureValue, originalBeginningValue)**
   - Calculates aggregate IRR by dividing future value by beginning value
   - Example: `=PC.AGGIRR(150, 100)` returns `1.5`

2. **PC.JOINCELLS(range, [delimiter])**
   - Joins cells from a range with specified delimiter (default: comma)
   - Example: `=PC.JOINCELLS(A1:A5, " | ")` returns cells joined with " | "

### Troubleshooting Custom Functions
- **Functions not visible:** Check manifest.xml has CustomFunctions ExtensionPoint
- **Functions not working:** Verify functions.json is accessible
- **Namespace issues:** Confirm namespace "PC" in manifest matches usage

---

## 🔧 Troubleshooting

### Common Issues & Solutions

**Build Failures**
- Clear cache: `npm run clean && npm install`
- Check Node.js version (16+ required)
- Verify TypeScript compilation: `npm run lint`

**Manifest Issues**
- Validate manifest: Office Add-in Validator
- Check URLs are accessible from Excel client
- Verify SSL certificates

**Custom Functions Not Loading**
- Check functions.json is generated and accessible
- Verify CustomFunctions ExtensionPoint in manifest
- Confirm commands.html loads properly

**IIS Deployment Issues**
- Check IIS site is running
- Verify SSL certificate binding
- Check file permissions in `C:\inetpub\wwwroot\ExcelAddin\`
- Review IIS logs for errors

**Network/SSL Issues**
```powershell
# Test connectivity
Test-NetConnection server-vs81t.intranet.local -Port 9443
Test-NetConnection server-vs84.intranet.local -Port 9443

# Test SSL certificate
$cert = Invoke-WebRequest -Uri "https://server-vs81t.intranet.local:9443" -SkipCertificateCheck
```

---

## 📝 Build Commands Reference

```bash
# Development
npm run dev                    # Start dev server
npm run build:dev              # Build for development
npm run build:functions        # Generate functions.json

# Staging
npm run build:staging          # Build for staging
npm run validate:staging       # Validate staging manifest

# Production  
npm run build:prod             # Build for production
npm run validate:prod          # Validate production manifest

# Utilities
npm run clean                  # Clean build artifacts
npm run lint                   # TypeScript validation
npm run analyze                # Bundle analysis
```

---

## 🔧 Certificate & HTTPS Troubleshooting

### Common Certificate Issues

#### Issue: "Office Add-in certificates not found" warning
**Solution:**
```bash
npm run cert:install
npm run cert:verify
```

#### Issue: "Failed to load webpack.config.js - ENOENT: localhost.key"
This indicates the development configuration is trying to use certificates that don't exist.
**Solution:**
1. Install Office Add-in certificates: `npm run cert:install`
2. Or use the fallback HTTPS configuration (automatic in updated webpack config)

#### Issue: Excel doesn't trust the development server
**Symptoms:** 
- Excel shows security warnings
- Add-in won't load from localhost
- "This add-in is not secure" messages

**Solution:**
```bash
# Uninstall and reinstall certificates
npm run cert:uninstall
npm run cert:install

# Verify certificates are trusted
npm run cert:verify
```

#### Issue: Different certificate requirements per environment
**Understanding:**
- **Local Development:** Uses Office Add-in development certificates in `~/.office-addin-dev-certs/`
- **Staging/Production:** Uses server-managed SSL certificates through IIS
- **Webpack configs:** Development vs production configs handle certificates differently

**Configuration:**
- Development webpack config auto-detects and uses Office Add-in certificates
- Production/staging webpack configs rely on server SSL configuration
- No cross-configuration conflicts

---

## 📞 Support Contacts

**Development Issues:** Contact development team
**Staging Environment:** Contact staging admin
**Production Environment:** Contact production admin
**Excel Add-in Issues:** Reference TROUBLESHOOTING_GUIDE.md

---

## 📄 Related Documentation

- [README.md](README.md) - Project overview
- [DEPLOYMENT_GUIDE.md](DEPLOYMENT_GUIDE.md) - IIS deployment details  
- [TROUBLESHOOTING_GUIDE.md](TROUBLESHOOTING_GUIDE.md) - Detailed troubleshooting
- [APPLICATION_GUIDE.md](APPLICATION_GUIDE.md) - User application guide