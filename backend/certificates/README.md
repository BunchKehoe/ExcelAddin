# SSL Certificates Configuration

This directory contains SSL certificates for secure HTTPS connections to external services.

## Certificate Files Required

Place the following certificate files in this directory:

### For NiFi HTTPS Connection

1. **`nifi-ca-cert.pem`** - CA certificate or certificate bundle for the NiFi server
   - This should contain the Certificate Authority (CA) certificate that signed the NiFi server's SSL certificate
   - If using a self-signed certificate, this would be the server's certificate itself

2. **`nifi-client-cert.pem`** (Optional) - Client certificate for mutual TLS authentication
   - Only needed if NiFi requires client certificate authentication
   - Contains the client certificate in PEM format

3. **`nifi-client-key.pem`** (Optional) - Private key for client certificate
   - Only needed if using client certificate authentication
   - Contains the private key corresponding to the client certificate
   - **Keep this file secure and never commit it to version control**

## Environment Variables

Configure these environment variables to specify certificate paths:

```bash
# CA certificate for server verification (required)
NIFI_CA_CERT_PATH=/path/to/backend/certificates/nifi-ca-cert.pem

# Client certificate and key for mutual TLS (optional)
NIFI_CLIENT_CERT_PATH=/path/to/backend/certificates/nifi-client-cert.pem
NIFI_CLIENT_KEY_PATH=/path/to/backend/certificates/nifi-client-key.pem

# Alternative: disable SSL verification (not recommended for production)
NIFI_VERIFY_SSL=false
```

## Certificate Format

All certificates should be in PEM format. If you have certificates in other formats:

- **Convert DER to PEM**: `openssl x509 -inform der -in certificate.der -out certificate.pem`
- **Convert P12/PFX to PEM**: `openssl pkcs12 -in certificate.p12 -out certificate.pem -nodes`

## Security Notes

- Never commit private keys (*.key files) to version control
- Set appropriate file permissions: `chmod 600` for private keys, `chmod 644` for certificates
- Use `.gitignore` to exclude sensitive certificate files
- Consider using environment variables or secret management systems for production deployments

## Testing Certificate Configuration

The application will log certificate loading status during startup. Check the logs for:
- Certificate file existence
- SSL verification settings
- Connection status to NiFi endpoint