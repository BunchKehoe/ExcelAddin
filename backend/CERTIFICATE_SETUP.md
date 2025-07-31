# HTTPS Certificate Configuration Guide

This guide explains how to configure HTTPS certificates for secure communication with the NiFi endpoint.

## Quick Setup

1. **Copy environment template:**
   ```bash
   cd backend
   cp .env.example .env
   ```

2. **Place your certificates in the certificates directory:**
   ```bash
   # CA certificate (required for SSL verification)
   cp your-nifi-ca-cert.pem certificates/nifi-ca-cert.pem
   
   # Client certificates (optional, only if NiFi requires mutual TLS)
   cp your-client-cert.pem certificates/nifi-client-cert.pem
   cp your-client-key.pem certificates/nifi-client-key.pem
   ```

3. **Configure environment variables in .env:**
   ```bash
   # Enable SSL verification
   NIFI_VERIFY_SSL=true
   
   # Specify certificate files (optional, defaults to nifi-ca-cert.pem)
   NIFI_CA_CERT_PATH=nifi-ca-cert.pem
   ```

## Configuration Options

### Option 1: Use Default Certificate Paths (Recommended)

Place your CA certificate as `backend/certificates/nifi-ca-cert.pem` and set:
```bash
NIFI_VERIFY_SSL=true
```

The application will automatically use the CA certificate for server verification.

### Option 2: Custom Certificate Paths

If you have certificates with different names, specify them explicitly:
```bash
NIFI_VERIFY_SSL=true
NIFI_CA_CERT_PATH=your-custom-ca.pem
NIFI_CLIENT_CERT_PATH=your-client.pem
NIFI_CLIENT_KEY_PATH=your-client.key
```

### Option 3: Disable SSL Verification (Development Only)

```bash
NIFI_VERIFY_SSL=false
```

**Warning:** Only use this for development. Never disable SSL verification in production.

## Certificate Types and Scenarios

### Scenario 1: Self-Signed NiFi Certificate

If NiFi uses a self-signed certificate:
1. Export the NiFi server certificate
2. Save it as `certificates/nifi-ca-cert.pem`
3. Configure: `NIFI_VERIFY_SSL=true`

### Scenario 2: Corporate CA-Signed Certificate

If NiFi uses a certificate signed by your corporate CA:
1. Obtain the corporate CA certificate bundle
2. Save it as `certificates/nifi-ca-cert.pem`
3. Configure: `NIFI_VERIFY_SSL=true`

### Scenario 3: Mutual TLS Authentication

If NiFi requires client certificate authentication:
1. Place CA certificate: `certificates/nifi-ca-cert.pem`
2. Place client certificate: `certificates/nifi-client-cert.pem`
3. Place client private key: `certificates/nifi-client-key.pem`
4. Configure:
   ```bash
   NIFI_VERIFY_SSL=true
   NIFI_CA_CERT_PATH=nifi-ca-cert.pem
   NIFI_CLIENT_CERT_PATH=nifi-client-cert.pem
   NIFI_CLIENT_KEY_PATH=nifi-client-key.pem
   ```

## Certificate File Formats

All certificates must be in PEM format. If you have certificates in other formats:

### Convert DER to PEM:
```bash
openssl x509 -inform der -in certificate.der -out certificate.pem
```

### Convert PKCS12/PFX to PEM:
```bash
# Extract certificate
openssl pkcs12 -in certificate.p12 -nokeys -out certificate.pem

# Extract private key
openssl pkcs12 -in certificate.p12 -nocerts -nodes -out private-key.pem
```

### Extract CA certificate from server:
```bash
# Get server certificate chain
openssl s_client -showcerts -connect server-vs81t.intranet.local:8443 < /dev/null 2>/dev/null | openssl x509 -outform PEM > nifi-server-cert.pem
```

## File Permissions

Set appropriate permissions for security:
```bash
# Make certificates readable
chmod 644 certificates/*.pem

# Make private keys secure (if any)
chmod 600 certificates/*-key.pem
```

## Troubleshooting

### Check Certificate Validity
```bash
# Verify certificate
openssl x509 -in certificates/nifi-ca-cert.pem -text -noout

# Check certificate expiration
openssl x509 -in certificates/nifi-ca-cert.pem -enddate -noout
```

### Test Connection
```bash
# Test SSL connection to NiFi
openssl s_client -connect server-vs81t.intranet.local:8443 -CAfile certificates/nifi-ca-cert.pem
```

### Application Logs

The application logs SSL configuration on startup:
```
INFO - NiFi endpoint: https://server-vs81t.intranet.local:8443/nifi/api/excel-addin-upload
INFO - NiFi SSL verification: True
INFO - NiFi SSL verify setting: /path/to/certificates/nifi-ca-cert.pem
INFO - CA certificate found: /path/to/certificates/nifi-ca-cert.pem
```

### Common Errors

1. **Certificate not found:** Check file paths and permissions
2. **SSL verification failed:** Verify CA certificate is correct for the server
3. **Hostname mismatch:** Ensure certificate matches server hostname
4. **Certificate expired:** Check certificate validity dates

## Security Best Practices

1. **Never commit private keys to version control**
2. **Use strong file permissions (600) for private keys**
3. **Regularly update certificates before expiration**
4. **Use proper CA certificates instead of disabling verification**
5. **Monitor certificate expiration dates**
6. **Use environment variables for sensitive configuration**

## Environment Variables Reference

| Variable | Default | Description |
|----------|---------|-------------|
| `NIFI_ENDPOINT` | `https://server-vs81t.intranet.local:8443/nifi/api/excel-addin-upload` | NiFi upload endpoint URL |
| `NIFI_VERIFY_SSL` | `true` | Enable/disable SSL certificate verification |
| `NIFI_CA_CERT_PATH` | `nifi-ca-cert.pem` | CA certificate file (relative to certificates/) |
| `NIFI_CLIENT_CERT_PATH` | None | Client certificate file (optional) |
| `NIFI_CLIENT_KEY_PATH` | None | Client private key file (optional) |