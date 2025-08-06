# Deploying Excel Add-in at /excellence Path

This document outlines the configuration changes needed to deploy the Excel Add-in at `https://server01.intranet.local:8443/excellence` instead of the root path.

## Configuration Files Updated

### 1. nginx Configuration (`deployment/nginx/excel-addin.conf`)

The nginx configuration has been updated to:
- Serve the application at `/excellence/` path
- Proxy API requests from `/excellence/api/` to the backend
- Handle static assets with proper caching
- Set up health check endpoints at both `/health` and `/excellence/health`
- Redirect root requests to `/excellence/`
- Use port 8443 instead of 443

### 2. Manifest Files

**manifest-staging.xml** has been updated with:
- All URLs changed to use `https://server01.intranet.local:8443/excellence/`
- Icons, taskpane, and command URLs updated
- AppDomain updated to match the server

### 3. Frontend API Client (`src/components/api/apiClient.ts`)

The API base URL has been changed from:
```typescript
const API_BASE_URL = 'http://localhost:5000/api';
```
to:
```typescript
const API_BASE_URL = 'https://server01.intranet.local:8443/excellence/api';
```

### 4. Webpack Configuration (`webpack.prod.config.js`)

The public path has been updated to:
```javascript
publicPath: '/excellence/'
```

### 5. Backend Environment Configuration (`backend/.env.production`)

CORS origins updated to include:
```
CORS_ORIGINS=https://server01.intranet.local:8443,https://localhost:3000
```

## Key URLs

After deployment, the Excel Add-in will be accessible at:
- **Main Application**: https://server01.intranet.local:8443/excellence/
- **Task Pane**: https://server01.intranet.local:8443/excellence/taskpane.html
- **Commands**: https://server01.intranet.local:8443/excellence/commands.html
- **API Endpoints**: https://server01.intranet.local:8443/excellence/api/
- **Health Check**: https://server01.intranet.local:8443/excellence/health
- **Assets**: https://server01.intranet.local:8443/excellence/assets/

## How nginx Routes Work

1. **API Requests**: Requests to `/excellence/api/*` are rewritten to `/api/*` and proxied to the backend Flask service at `127.0.0.1:5000`
2. **Static Files**: Requests to `/excellence/*` are served from `C:/inetpub/wwwroot/ExcelAddin/dist/`
3. **Root Redirects**: Any request to `/` or other paths not matching `/excellence/` or `/health` are redirected to `/excellence/`

## Deployment Steps

1. Use the updated configuration files
2. Ensure nginx is configured to listen on port 8443 with SSL
3. Build the frontend with the updated webpack configuration
4. Deploy the backend with the updated environment configuration
5. Install the updated manifest file in Excel

## Testing

After deployment, verify:
- [ ] https://server01.intranet.local:8443/excellence/ loads the task pane
- [ ] https://server01.intranet.local:8443/excellence/api/health returns healthy status
- [ ] Excel Add-in loads correctly using the updated manifest
- [ ] API calls from the frontend work correctly
- [ ] Static assets (CSS, JS, images) load properly

## Troubleshooting

If the add-in doesn't load:
1. Check nginx access/error logs at `C:/Logs/nginx/`
2. Verify SSL certificate is valid for `server01.intranet.local`
3. Ensure the manifest file is correctly installed in Excel
4. Check that the backend service is running on port 5000
5. Verify CORS configuration allows requests from the domain