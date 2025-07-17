#!/bin/bash

# PrimeExcelence Excel Addin Setup Script
# This script helps set up the development environment for the Excel addin

set -e

echo "🚀 Setting up PrimeExcelence Excel Addin Development Environment"
echo "=============================================================="

# Check if Node.js is installed
if ! command -v node &> /dev/null; then
    echo "❌ Node.js is not installed. Please install Node.js from https://nodejs.org/"
    exit 1
fi

# Check Node.js version
NODE_VERSION=$(node --version | cut -d'v' -f2)
MIN_VERSION="16.0.0"

if [ "$(printf '%s\n' "$MIN_VERSION" "$NODE_VERSION" | sort -V | head -n1)" != "$MIN_VERSION" ]; then
    echo "❌ Node.js version $NODE_VERSION is too old. Please upgrade to Node.js 16 or higher."
    exit 1
fi

echo "✅ Node.js version $NODE_VERSION is compatible"

# Install npm dependencies
echo "📦 Installing dependencies..."
npm install

# Install office-addin-dev-certs globally if not already installed
if ! npm list -g office-addin-dev-certs &> /dev/null; then
    echo "🔒 Installing office-addin-dev-certs for HTTPS support..."
    npm install -g office-addin-dev-certs
fi

# Install development certificates
echo "🔐 Installing development certificates..."
office-addin-dev-certs install --machine

# Verify certificate installation
if office-addin-dev-certs verify; then
    echo "✅ Development certificates installed successfully"
else
    echo "⚠️  Certificate verification failed. You may need to manually trust the certificates."
fi

# Build the project to ensure everything works
echo "🔨 Building the project..."
npm run build

echo ""
echo "✅ Setup complete! You can now:"
echo "   1. Start the development server: npm start"
echo "   2. Open Excel and sideload the manifest.xml file"
echo "   3. Access the addin at https://localhost:3000"
echo ""
echo "📖 For detailed instructions, see the README.md file"
echo "🛠️  For troubleshooting, check the Troubleshooting section in README.md"