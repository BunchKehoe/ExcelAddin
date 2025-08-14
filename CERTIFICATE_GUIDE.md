# Excel Add-in Certificate Management Guide

## Overview

Excel Add-ins require HTTPS connections for security. In local development, the add-in uses self-signed certificates installed by the `office-addin-dev-certs` package. When these certificates are missing, expired, or not trusted by the system, you'll see certificate errors in Excel.

## Common Certificate Error

**Error Message:** "The content is blocked because it isn't signed by a valid security certificate."

This error occurs when:
- Certificates are not installed
- Certificates have expired
- Certificates are not trusted by Windows
- Certificate CA has changed or been revoked

## Solution Steps

### 1. Install Office Add-in Development Certificates

```bash
# Install certificates (this may require administrator privileges)
npm run cert:install

# OR run directly
npx office-addin-dev-certs install
```

**What this does:**
- Creates `~/.office-addin-dev-certs/` directory
- Generates `localhost.crt` and `localhost.key` files
- Installs the CA certificate in Windows certificate store
- Makes localhost HTTPS trusted for Office Add-ins

### 2. Verify Certificate Installation

```bash
# Check if certificates are properly installed
npm run cert:verify

# OR run directly
npx office-addin-dev-certs verify
```

**Expected Output:**
```
Certificates are already installed.
```

### 3. If Certificates Expired or Corrupted

```bash
# Uninstall old certificates
npm run cert:uninstall

# Clean install new certificates
npm run cert:install
```

### 4. Manual Certificate Validation (Windows)

1. **Check Certificate Store:**
   - Press `Win + R`, type `certlm.msc`, press Enter
   - Navigate to: **Trusted Root Certification Authorities** → **Certificates**
   - Look for "office-addin-dev-ca" certificate
   - If missing or expired, run `npm run cert:install`

2. **Check File System:**
   - Verify files exist at: `%USERPROFILE%\.office-addin-dev-certs\`
   - Should contain: `localhost.crt`, `localhost.key`, and CA files
   - If missing, run `npm run cert:install`

### 5. Alternative: Use HTTP (Development Only)

If certificates continue to cause issues, you can temporarily modify the manifest for development:

⚠️ **WARNING: This is for debugging only and reduces security**

```xml
<!-- In manifest.xml, change HTTPS URLs to HTTP -->
<SourceLocation DefaultValue="http://localhost:3000/taskpane.html"/>
<!-- Note: This may not work with all Office versions -->
```

## Certificate Lifecycle Management

### When to Renew Certificates

- **Automatic expiry:** Office dev certificates expire periodically
- **Windows updates:** Major Windows updates may affect certificate store
- **Office updates:** New Office versions may require certificate refresh
- **Developer environment changes:** New user account or development machine

### Recommended Schedule

```bash
# Monthly certificate health check
npm run cert:verify

# If verification fails, refresh certificates
npm run cert:uninstall && npm run cert:install
```

## Troubleshooting Advanced Issues

### Issue: "Administrator privileges required"

**Solution:**
```powershell
# Run as Administrator
Start-Process powershell -Verb runAs
cd C:\path\to\ExcelAddin
npm run cert:install
```

### Issue: Corporate Firewall/Proxy

**Solution:**
- Contact IT to whitelist `localhost:3000` for HTTPS
- Ensure Windows certificate store changes are allowed
- Consider using corporate development certificate if available

### Issue: Antivirus Software Interference

**Solution:**
- Temporarily disable real-time protection during certificate installation
- Add `.office-addin-dev-certs` folder to antivirus exclusions
- Re-enable protection after successful installation

### Issue: Multiple Node.js/npm Versions

**Solution:**
```bash
# Ensure using same npm version for installation and verification
which npm
npm --version

# Clear npm cache if needed
npm cache clean --force
```

## Development Workflow Integration

### Automated Setup Script

Add this to your development setup:

```bash
#!/bin/bash
# setup-development.sh

echo "Setting up Excel Add-in development environment..."

# Install dependencies
npm install

# Install/verify certificates
echo "Checking certificates..."
if ! npm run cert:verify > /dev/null 2>&1; then
    echo "Installing Office Add-in development certificates..."
    npm run cert:install
fi

# Start development server
echo "Starting development server..."
npm run dev
```

### Integration with Build Process

The Vite configuration automatically handles certificates:

```javascript
// vite.config.mjs handles certificate detection
const certPath = homedir() + '/.office-addin-dev-certs/localhost.crt';
const keyPath = homedir() + '/.office-addin-dev-certs/localhost.key';

if (existsSync(certPath) && existsSync(keyPath)) {
    // Use proper certificates
} else {
    // Fallback with warning
    console.warn('⚠️  Office Add-in certificates not found. Run "npm run cert:install"');
}
```

## Environment-Specific Notes

### Local Development
- Uses `https://localhost:3000` with self-signed certificates
- Requires trusted certificate authority in Windows store
- Certificates managed by `office-addin-dev-certs` package

### Staging/Production
- Uses proper SSL certificates from IT infrastructure
- No self-signed certificates required
- Standard HTTPS certificate validation

## Quick Fix Summary

**If you see certificate errors in Excel:**

1. Run: `npm run cert:install`
2. Restart Excel completely
3. Reload the add-in
4. If still failing, run: `npm run cert:uninstall && npm run cert:install`

**Commands to remember:**
```bash
npm run cert:install   # Install certificates
npm run cert:verify    # Check certificate status
npm run cert:uninstall # Remove certificates
```

**Files to check:**
- `%USERPROFILE%\.office-addin-dev-certs\localhost.crt`
- `%USERPROFILE%\.office-addin-dev-certs\localhost.key`
- Windows Certificate Store: `certlm.msc` → Trusted Root CA → office-addin-dev-ca