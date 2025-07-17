#!/bin/bash

# Certificate generation script for PrimeExcelence Excel Addin
# This script generates development certificates for local testing

echo "🔐 Generating development certificates for Excel Addin..."

# Check if openssl is available
if ! command -v openssl &> /dev/null; then
    echo "❌ OpenSSL is not installed. Please install OpenSSL first."
    exit 1
fi

# Create certificates directory
mkdir -p certs

# Generate private key
echo "📄 Generating private key..."
openssl genrsa -out certs/localhost-key.pem 2048

# Generate certificate signing request
echo "📋 Generating certificate signing request..."
openssl req -new -key certs/localhost-key.pem -out certs/localhost.csr -subj "/C=US/ST=CA/L=San Francisco/O=PrimeExcelence/OU=IT/CN=localhost"

# Generate self-signed certificate
echo "🔒 Generating self-signed certificate..."
openssl x509 -req -in certs/localhost.csr -signkey certs/localhost-key.pem -out certs/localhost.pem -days 365 -extensions v3_req -extfile <(cat <<EOF
[req]
distinguished_name = req_distinguished_name
req_extensions = v3_req
prompt = no

[req_distinguished_name]
C = US
ST = CA
L = San Francisco
O = PrimeExcelence
OU = IT
CN = localhost

[v3_req]
keyUsage = keyEncipherment, dataEncipherment
extendedKeyUsage = serverAuth
subjectAltName = @alt_names

[alt_names]
DNS.1 = localhost
DNS.2 = 127.0.0.1
IP.1 = 127.0.0.1
IP.2 = ::1
EOF
)

# Clean up CSR file
rm certs/localhost.csr

echo "✅ Certificates generated successfully!"
echo "📁 Certificate files:"
echo "   - Private key: certs/localhost-key.pem"
echo "   - Certificate: certs/localhost.pem"
echo ""
echo "⚠️  Note: You'll need to trust these certificates in your browser and system"
echo "💡 For easier setup, use: npm run cert:install"