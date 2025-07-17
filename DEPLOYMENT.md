# PrimeExcelence Excel Addin - Deployment Guide

This guide provides detailed instructions for deploying the PrimeExcelence Excel Addin to production environments.

## Table of Contents

- [Prerequisites](#prerequisites)
- [SSL Certificate Setup](#ssl-certificate-setup)
- [Deployment Options](#deployment-options)
- [Azure Deployment](#azure-deployment)
- [AWS Deployment](#aws-deployment)
- [IIS Deployment](#iis-deployment)
- [Nginx Deployment](#nginx-deployment)
- [Manifest Configuration](#manifest-configuration)
- [Distribution Methods](#distribution-methods)
- [Monitoring and Maintenance](#monitoring-and-maintenance)

## Prerequisites

Before deploying to production, ensure you have:

- A registered domain name
- SSL certificate (Let's Encrypt or commercial CA)
- Web hosting service or server
- Access to DNS management
- Excel admin rights (for organization-wide deployment)

## SSL Certificate Setup

### Option 1: Let's Encrypt (Free)

```bash
# Install Certbot
sudo apt-get update
sudo apt-get install certbot python3-certbot-nginx

# Generate certificate
sudo certbot --nginx -d yourdomain.com

# Auto-renewal setup
sudo crontab -e
# Add line: 0 12 * * * /usr/bin/certbot renew --quiet
```

### Option 2: Commercial Certificate

1. Purchase SSL certificate from a CA (GoDaddy, DigiCert, etc.)
2. Generate Certificate Signing Request (CSR)
3. Submit CSR to CA
4. Download and install certificate

### Option 3: Azure App Service (Managed)

Azure App Service provides free SSL certificates for custom domains.

## Deployment Options

### Option 1: Static Web Hosting

Suitable for: Simple deployments, CDN integration

**Steps:**
1. Build the project: `npm run build`
2. Upload `dist` folder to your web host
3. Configure domain and SSL
4. Update manifest with production URLs

### Option 2: Node.js Server

Suitable for: Dynamic content, API integration

**Steps:**
1. Set up Node.js on your server
2. Install dependencies: `npm install --production`
3. Configure reverse proxy (Nginx/Apache)
4. Set up SSL termination
5. Configure domain

### Option 3: Container Deployment

Suitable for: Scalable deployments, microservices

**Steps:**
1. Create Dockerfile
2. Build container image
3. Deploy to container platform (Docker, Kubernetes)
4. Configure ingress for HTTPS

## Azure Deployment

### Azure Static Web Apps

1. **Create Azure Static Web App**
   ```bash
   az staticwebapp create \
     --name primeexcelence-addin \
     --resource-group myResourceGroup \
     --source https://github.com/yourusername/ExcelAddin \
     --location "Central US" \
     --branch main \
     --app-location "/" \
     --output-location "dist"
   ```

2. **Configure Build Settings**
   Create `.github/workflows/azure-static-web-apps-<name>.yml`:
   ```yaml
   name: Azure Static Web Apps CI/CD

   on:
     push:
       branches:
         - main
     pull_request:
       types: [opened, synchronize, reopened, closed]
       branches:
         - main

   jobs:
     build_and_deploy_job:
       runs-on: ubuntu-latest
       name: Build and Deploy Job
       steps:
         - uses: actions/checkout@v3
         - name: Setup Node.js
           uses: actions/setup-node@v3
           with:
             node-version: '16'
         - name: Install dependencies
           run: npm install
         - name: Build
           run: npm run build
         - name: Deploy to Azure Static Web Apps
           uses: Azure/static-web-apps-deploy@v1
           with:
             azure_static_web_apps_api_token: ${{ secrets.AZURE_STATIC_WEB_APPS_API_TOKEN }}
             repo_token: ${{ secrets.GITHUB_TOKEN }}
             action: "upload"
             app_location: "/"
             output_location: "dist"
   ```

### Azure App Service

1. **Create App Service**
   ```bash
   az webapp create \
     --resource-group myResourceGroup \
     --plan myAppServicePlan \
     --name primeexcelence-addin \
     --runtime "NODE|16-lts"
   ```

2. **Configure SSL**
   - Add custom domain in Azure Portal
   - Enable SSL certificate (free or bring your own)

3. **Deploy Code**
   ```bash
   # Using Azure CLI
   az webapp deployment source config \
     --name primeexcelence-addin \
     --resource-group myResourceGroup \
     --repo-url https://github.com/yourusername/ExcelAddin \
     --branch main \
     --manual-integration
   ```

## AWS Deployment

### AWS S3 + CloudFront

1. **Create S3 Bucket**
   ```bash
   aws s3 mb s3://primeexcelence-addin-bucket
   ```

2. **Configure Static Website Hosting**
   ```bash
   aws s3 website s3://primeexcelence-addin-bucket \
     --index-document taskpane.html \
     --error-document error.html
   ```

3. **Create CloudFront Distribution**
   ```bash
   aws cloudfront create-distribution \
     --distribution-config file://cloudfront-config.json
   ```

4. **Configure SSL Certificate**
   - Request certificate via AWS Certificate Manager
   - Associate with CloudFront distribution

### AWS Elastic Beanstalk

1. **Install EB CLI**
   ```bash
   pip install awsebcli
   ```

2. **Initialize Application**
   ```bash
   eb init primeexcelence-addin
   ```

3. **Deploy**
   ```bash
   eb create production
   eb deploy
   ```

## IIS Deployment

### Windows Server Setup

1. **Install IIS and Node.js**
   ```powershell
   # Install IIS
   Enable-WindowsOptionalFeature -Online -FeatureName IIS-WebServerRole
   Enable-WindowsOptionalFeature -Online -FeatureName IIS-WebServer
   Enable-WindowsOptionalFeature -Online -FeatureName IIS-CommonHttpFeatures
   Enable-WindowsOptionalFeature -Online -FeatureName IIS-HttpErrors
   Enable-WindowsOptionalFeature -Online -FeatureName IIS-HttpRedirect
   Enable-WindowsOptionalFeature -Online -FeatureName IIS-ApplicationDevelopment
   Enable-WindowsOptionalFeature -Online -FeatureName IIS-NetFxExtensibility45
   Enable-WindowsOptionalFeature -Online -FeatureName IIS-HealthAndDiagnostics
   Enable-WindowsOptionalFeature -Online -FeatureName IIS-HttpLogging
   Enable-WindowsOptionalFeature -Online -FeatureName IIS-Security
   Enable-WindowsOptionalFeature -Online -FeatureName IIS-RequestFiltering
   Enable-WindowsOptionalFeature -Online -FeatureName IIS-Performance
   Enable-WindowsOptionalFeature -Online -FeatureName IIS-WebServerManagementTools
   Enable-WindowsOptionalFeature -Online -FeatureName IIS-ManagementConsole
   Enable-WindowsOptionalFeature -Online -FeatureName IIS-IIS6ManagementCompatibility
   Enable-WindowsOptionalFeature -Online -FeatureName IIS-Metabase
   ```

2. **Configure IIS Site**
   ```xml
   <!-- web.config -->
   <?xml version="1.0" encoding="utf-8"?>
   <configuration>
     <system.webServer>
       <staticContent>
         <mimeMap fileExtension=".json" mimeType="application/json" />
         <mimeMap fileExtension=".xml" mimeType="application/xml" />
       </staticContent>
       <httpProtocol>
         <customHeaders>
           <add name="Access-Control-Allow-Origin" value="*" />
           <add name="Access-Control-Allow-Methods" value="GET, POST, PUT, DELETE, OPTIONS" />
           <add name="Access-Control-Allow-Headers" value="Content-Type, Authorization" />
         </customHeaders>
       </httpProtocol>
       <rewrite>
         <rules>
           <rule name="React Routes" stopProcessing="true">
             <match url=".*" />
             <conditions logicalGrouping="MatchAll">
               <add input="{REQUEST_FILENAME}" matchType="IsFile" negate="true" />
               <add input="{REQUEST_FILENAME}" matchType="IsDirectory" negate="true" />
             </conditions>
             <action type="Rewrite" url="/taskpane.html" />
           </rule>
         </rules>
       </rewrite>
     </system.webServer>
   </configuration>
   ```

3. **Install SSL Certificate**
   - Use IIS Manager to import certificate
   - Configure HTTPS binding

## Nginx Deployment

### Ubuntu/Debian Setup

1. **Install Nginx**
   ```bash
   sudo apt update
   sudo apt install nginx
   ```

2. **Configure Site**
   ```nginx
   # /etc/nginx/sites-available/primeexcelence-addin
   server {
       listen 80;
       server_name yourdomain.com;
       return 301 https://$server_name$request_uri;
   }

   server {
       listen 443 ssl http2;
       server_name yourdomain.com;

       ssl_certificate /etc/letsencrypt/live/yourdomain.com/fullchain.pem;
       ssl_certificate_key /etc/letsencrypt/live/yourdomain.com/privkey.pem;
       ssl_protocols TLSv1.2 TLSv1.3;
       ssl_ciphers ECDHE-RSA-AES256-GCM-SHA512:DHE-RSA-AES256-GCM-SHA512:ECDHE-RSA-AES256-GCM-SHA384:DHE-RSA-AES256-GCM-SHA384;
       ssl_prefer_server_ciphers off;
       ssl_session_cache shared:SSL:10m;

       root /var/www/primeexcelence-addin/dist;
       index taskpane.html;

       location / {
           try_files $uri $uri/ /taskpane.html;
       }

       location ~* \.(js|css|png|jpg|jpeg|gif|ico|svg)$ {
           expires 1y;
           add_header Cache-Control "public, immutable";
       }

       # Security headers
       add_header X-Frame-Options "SAMEORIGIN" always;
       add_header X-Content-Type-Options "nosniff" always;
       add_header X-XSS-Protection "1; mode=block" always;
       add_header Strict-Transport-Security "max-age=31536000; includeSubDomains" always;

       # CORS headers for Office.js
       add_header Access-Control-Allow-Origin "*" always;
       add_header Access-Control-Allow-Methods "GET, POST, PUT, DELETE, OPTIONS" always;
       add_header Access-Control-Allow-Headers "Content-Type, Authorization" always;
   }
   ```

3. **Enable Site**
   ```bash
   sudo ln -s /etc/nginx/sites-available/primeexcelence-addin /etc/nginx/sites-enabled/
   sudo nginx -t
   sudo systemctl restart nginx
   ```

## Manifest Configuration

### Production Manifest Updates

1. **Update manifest-prod.xml**
   ```xml
   <Id>generate-unique-guid-here</Id>
   <ProviderName>Your Company Name</ProviderName>
   <IconUrl DefaultValue="https://yourdomain.com/assets/icon-32.png"/>
   <SupportUrl DefaultValue="https://yourdomain.com/support"/>
   <AppDomains>
     <AppDomain>https://yourdomain.com</AppDomain>
   </AppDomains>
   <SourceLocation DefaultValue="https://yourdomain.com/taskpane.html"/>
   ```

2. **Generate Unique GUID**
   ```bash
   # Using PowerShell
   [System.Guid]::NewGuid()

   # Using Python
   python -c "import uuid; print(uuid.uuid4())"

   # Online generator
   # https://www.guidgenerator.com/
   ```

## Distribution Methods

### Microsoft AppSource

1. **Prepare for AppSource**
   - Complete manifest validation
   - Prepare marketing materials
   - Create privacy policy and terms of service
   - Set up support channels

2. **Submit to AppSource**
   - Create Partner Center account
   - Upload manifest and materials
   - Complete certification process

### SharePoint App Catalog

1. **Upload to App Catalog**
   ```powershell
   # PowerShell with SharePoint Online Management Shell
   Connect-SPOService -Url https://yourtenant-admin.sharepoint.com
   Add-SPOAppCatalogSite -Site https://yourtenant.sharepoint.com/sites/appcatalog
   ```

2. **Deploy Organization-Wide**
   - Upload manifest to App Catalog
   - Configure deployment settings
   - Users can access from Office Store

### Network Share Distribution

1. **Set up Network Share**
   ```bash
   # Windows
   net share AddinManifests=C:\AddinManifests /grant:everyone,read

   # Linux (Samba)
   sudo apt install samba
   # Configure smb.conf
   ```

2. **Configure Excel Trust Center**
   - File → Options → Trust Center → Trust Center Settings
   - Trusted Add-in Catalogs
   - Add network path

## Monitoring and Maintenance

### Health Checks

1. **Application Health**
   ```bash
   # Simple health check script
   #!/bin/bash
   DOMAIN="https://yourdomain.com"
   
   if curl -f -s $DOMAIN/taskpane.html > /dev/null; then
       echo "Application is healthy"
   else
       echo "Application is down"
       # Send alert
   fi
   ```

2. **SSL Certificate Monitoring**
   ```bash
   # Check SSL expiration
   openssl s_client -connect yourdomain.com:443 -servername yourdomain.com < /dev/null | openssl x509 -noout -dates
   ```

### Log Monitoring

1. **Application Logs**
   - Configure structured logging
   - Set up log aggregation (ELK stack, Splunk)
   - Monitor for errors and performance issues

2. **Web Server Logs**
   - Monitor access logs for usage patterns
   - Set up alerts for 4xx/5xx errors
   - Track performance metrics

### Updates and Versioning

1. **Version Management**
   - Update manifest version for each release
   - Use semantic versioning
   - Test thoroughly before deployment

2. **Rollback Strategy**
   - Keep previous versions available
   - Plan rollback procedures
   - Test rollback process

### Security Considerations

1. **Regular Updates**
   - Keep dependencies updated
   - Monitor for security vulnerabilities
   - Apply security patches promptly

2. **Security Headers**
   - Implement CSP headers
   - Use HTTPS everywhere
   - Regular security audits

## Troubleshooting

### Common Production Issues

1. **SSL Certificate Issues**
   - Check certificate validity
   - Verify certificate chain
   - Ensure proper domain configuration

2. **CORS Problems**
   - Verify CORS headers
   - Check domain whitelist
   - Test from different origins

3. **Performance Issues**
   - Enable compression
   - Optimize asset delivery
   - Use CDN for static assets

### Performance Optimization

1. **Asset Optimization**
   ```bash
   # Compress assets
   npm install -g gzip-cli
   gzip -k dist/*.js dist/*.css
   ```

2. **CDN Integration**
   - Use CDN for static assets
   - Configure proper cache headers
   - Implement asset versioning

---

This deployment guide provides comprehensive instructions for deploying the PrimeExcelence Excel Addin to production environments. Choose the deployment method that best fits your infrastructure and requirements.