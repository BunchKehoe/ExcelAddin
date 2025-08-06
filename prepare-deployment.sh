#!/bin/bash

# Excel Add-in Deployment Preparation Script
# This script prepares the application for Windows deployment

set -e  # Exit on any error

echo "================================================================"
echo "Excel Add-in Deployment Preparation"
echo "================================================================"

# Configuration
DEPLOY_DIR="./deployment-package"
TIMESTAMP=$(date +"%Y%m%d_%H%M%S")
PACKAGE_NAME="excel-addin-${TIMESTAMP}.tar.gz"

# Clean previous builds
echo "Cleaning previous builds..."
npm run clean
rm -rf "${DEPLOY_DIR}"

# Create deployment directory
echo "Creating deployment package directory..."
mkdir -p "${DEPLOY_DIR}"

# Install dependencies
echo "Installing dependencies..."
npm install

# Build frontend for production
echo "Building frontend for production..."
npm run build:staging

# Prepare deployment package
echo "Preparing deployment package..."

# Copy built frontend
cp -r dist "${DEPLOY_DIR}/"

# Copy backend
cp -r backend "${DEPLOY_DIR}/"

# Copy manifests
cp manifest-staging.xml "${DEPLOY_DIR}/"
cp manifest.xml "${DEPLOY_DIR}/" 2>/dev/null || echo "Note: manifest.xml not found, using staging manifest"

# Copy deployment configurations
cp -r deployment "${DEPLOY_DIR}/"

# Copy documentation
cp README.md "${DEPLOY_DIR}/"
cp DEPLOYMENT.md "${DEPLOY_DIR}/" 2>/dev/null || echo "Note: DEPLOYMENT.md not found"

# Create deployment info file
cat > "${DEPLOY_DIR}/DEPLOYMENT_INFO.txt" << EOF
Excel Add-in Deployment Package
Generated: $(date)
Build Type: Production/Staging
Frontend Build: Yes (webpack production mode)
Backend Included: Yes

Contents:
- dist/                 - Frontend build output
- backend/              - Python backend application
- deployment/           - Deployment configurations and scripts
- manifest-staging.xml  - Excel add-in manifest for staging

Deployment Instructions:
1. Copy this package to your Windows Server
2. Follow instructions in deployment/WINDOWS_DEPLOYMENT.md
3. Run deployment/scripts/deploy-windows.ps1 as Administrator

Required Services:
- nginx (reverse proxy)
- Python 3.8+ (backend runtime)
- NSSM (Windows service manager)

Default Ports:
- 443 (HTTPS frontend)
- 80 (HTTP redirect)
- 5000 (Backend API - internal)

Notes:
- Update domain names in configuration files before deployment
- Ensure SSL certificates are properly configured
- Backend service will run as Windows Service
- Frontend served via nginx with appropriate CORS headers for Excel
EOF

# Create package archive
echo "Creating deployment package archive..."
tar -czf "${PACKAGE_NAME}" -C "${DEPLOY_DIR}" .

# Display package info
echo ""
echo "================================================================"
echo "Deployment Package Created Successfully!"
echo "================================================================"
echo "Package: ${PACKAGE_NAME}"
echo "Size: $(du -h ${PACKAGE_NAME} | cut -f1)"
echo "Contents: $(tar -tzf ${PACKAGE_NAME} | wc -l) files"
echo ""
echo "Next Steps:"
echo "1. Transfer ${PACKAGE_NAME} to your Windows Server"
echo "2. Extract the package"
echo "3. Follow deployment/WINDOWS_DEPLOYMENT.md"
echo "4. Run deployment/scripts/deploy-windows.ps1 as Administrator"
echo ""
echo "Quick deployment on Windows Server:"
echo "  tar -xzf ${PACKAGE_NAME}"
echo "  PowerShell -ExecutionPolicy Bypass -File deployment/scripts/deploy-windows.ps1"
echo ""

# Validate the package
echo "Validating package contents..."
REQUIRED_FILES=(
    "dist/taskpane.html"
    "dist/taskpane"  # Match any taskpane.*.js file
    "backend/app.py"
    "backend/requirements.txt"
    "deployment/nginx/excel-addin.conf"
    "deployment/scripts/deploy-windows.ps1"
    "deployment/WINDOWS_DEPLOYMENT.md"
    "manifest-staging.xml"
)

ALL_VALID=true
for file in "${REQUIRED_FILES[@]}"; do
    if [[ "$file" == *"taskpane"* && "$file" != *".html" ]]; then
        # Special handling for taskpane.*.js files
        if tar -tzf "${PACKAGE_NAME}" | grep -q "taskpane.*\.js"; then
            echo "✓ ${file}*.js (webpack generated)"
        else
            echo "✗ Missing: ${file}*.js"
            ALL_VALID=false
        fi
    elif tar -tzf "${PACKAGE_NAME}" | grep -q "${file}"; then
        echo "✓ ${file}"
    else
        echo "✗ Missing: ${file}"
        ALL_VALID=false
    fi
done

if [ "$ALL_VALID" = true ]; then
    echo ""
    echo "✓ Package validation successful - all required files present"
    echo "Package is ready for deployment!"
else
    echo ""
    echo "✗ Package validation failed - some required files are missing"
    echo "Please check the build process and try again"
    exit 1
fi