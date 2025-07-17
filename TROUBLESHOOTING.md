# Troubleshooting Guide - PrimeExcelence Excel Addin

This guide covers common issues and solutions when developing and deploying the PrimeExcelence Excel Addin.

## Table of Contents

- [Development Environment Issues](#development-environment-issues)
- [SSL Certificate Problems](#ssl-certificate-problems)
- [Manifest Issues](#manifest-issues)
- [Excel Integration Problems](#excel-integration-problems)
- [Build and Dependencies](#build-and-dependencies)
- [Production Deployment Issues](#production-deployment-issues)
- [Performance Issues](#performance-issues)
- [Debugging Tips](#debugging-tips)

## Development Environment Issues

### Issue: "npm start" fails with EADDRINUSE error

**Problem**: Port 3000 is already in use

**Solution**:
```bash
# Find process using port 3000
lsof -i :3000

# Kill the process
kill -9 <PID>

# Or use a different port
npm start -- --port 3001
```

### Issue: "Cannot resolve module" errors

**Problem**: Missing dependencies or incorrect import paths

**Solution**:
```bash
# Clear npm cache
npm cache clean --force

# Delete node_modules and reinstall
rm -rf node_modules package-lock.json
npm install

# Check TypeScript configuration
npx tsc --noEmit
```

### Issue: Hot reload not working

**Problem**: Webpack dev server configuration

**Solution**:
```javascript
// webpack.config.js
devServer: {
  hot: true,
  liveReload: true,
  watchFiles: ['src/**/*'],
  // ... other config
}
```

## SSL Certificate Problems

### Issue: "Your connection is not private" in browser

**Problem**: Self-signed certificate not trusted

**Solution**:
```bash
# Install office-addin-dev-certs
npm install -g office-addin-dev-certs

# Install development certificate
office-addin-dev-certs install

# Verify installation
office-addin-dev-certs verify
```

### Issue: Certificate expired

**Problem**: Development certificate has expired

**Solution**:
```bash
# Uninstall old certificate
office-addin-dev-certs uninstall

# Install new certificate
office-addin-dev-certs install

# Restart development server
npm start
```

### Issue: Certificate not trusted in Excel

**Problem**: Excel doesn't trust the development certificate

**Solution**:
1. Close Excel completely
2. Install certificate for machine (not just user):
   ```bash
   office-addin-dev-certs install --machine
   ```
3. Restart Excel
4. Clear Excel cache:
   - Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
   - Mac: `~/Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`

## Manifest Issues

### Issue: "Manifest validation failed"

**Problem**: Invalid XML or missing required elements

**Solution**:
```bash
# Install manifest validator
npm install -g office-addin-validator

# Validate manifest
office-addin-validator manifest.xml

# Common fixes:
# - Ensure unique GUID for <Id>
# - Check all URLs are accessible
# - Verify XML syntax
```

### Issue: "Add-in won't load" after sideloading

**Problem**: Manifest URLs are not accessible

**Solution**:
1. Check that development server is running
2. Verify URLs in manifest are accessible:
   ```bash
   curl -I https://localhost:3000/taskpane.html
   curl -I https://localhost:3000/commands.html
   curl -I https://localhost:3000/assets/icon-32.png
   ```
3. Update manifest URLs if needed

### Issue: Icons not displaying

**Problem**: Icon URLs are not accessible or wrong format

**Solution**:
1. Verify icon files exist in assets folder
2. Check icon URLs in manifest:
   ```xml
   <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
   <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
   <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
   ```
3. Ensure icons are PNG format and correct dimensions

## Excel Integration Problems

### Issue: "Office is not defined" error

**Problem**: Office.js not loaded or initialized

**Solution**:
```typescript
// Check if Office.js is available
if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
      // Initialize app
      initializeApp();
    }
  });
} else {
  // Fallback for development
  document.addEventListener('DOMContentLoaded', initializeApp);
}
```

### Issue: "Excel.run is not a function" error

**Problem**: Excel API not available

**Solution**:
```typescript
// Check if Excel API is available
if (typeof Excel !== 'undefined') {
  Excel.run(async (context) => {
    // Your Excel code here
  });
} else {
  console.log('Excel API not available - running in browser mode');
  // Fallback behavior
}
```

### Issue: Data not inserting into Excel

**Problem**: Incorrect range selection or API usage

**Solution**:
```typescript
// Correct way to insert data
Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const range = sheet.getRange("A1:C10");
  
  range.values = [
    ["Header1", "Header2", "Header3"],
    ["Data1", "Data2", "Data3"]
  ];
  
  await context.sync();
});
```

## Build and Dependencies

### Issue: TypeScript compilation errors

**Problem**: Type mismatches or missing declarations

**Solution**:
```bash
# Check TypeScript configuration
npx tsc --noEmit

# Install missing type definitions
npm install --save-dev @types/office-js @types/react @types/react-dom

# Common fixes in tsconfig.json:
{
  "compilerOptions": {
    "skipLibCheck": true,
    "allowSyntheticDefaultImports": true,
    "esModuleInterop": true
  }
}
```

### Issue: Webpack build fails

**Problem**: Incorrect webpack configuration or missing loaders

**Solution**:
```bash
# Clear webpack cache
rm -rf node_modules/.cache

# Check webpack configuration
npx webpack --config webpack.config.js --mode development

# Common fixes:
# - Ensure all loaders are installed
# - Check file extensions in resolve
# - Verify entry points exist
```

### Issue: Material UI styling issues

**Problem**: Incorrect theme or component usage

**Solution**:
```typescript
// Ensure proper theme provider
import { ThemeProvider, createTheme } from '@mui/material/styles';

const theme = createTheme({
  palette: {
    primary: {
      main: '#1976d2',
    },
  },
});

// Wrap app with theme provider
<ThemeProvider theme={theme}>
  <App />
</ThemeProvider>
```

## Production Deployment Issues

### Issue: "Mixed content" errors

**Problem**: HTTP resources loaded over HTTPS

**Solution**:
1. Ensure all resources use HTTPS URLs
2. Check for hardcoded HTTP URLs in code
3. Update API endpoints to use HTTPS

### Issue: CORS errors in production

**Problem**: Cross-origin requests blocked

**Solution**:
```javascript
// Server-side (Express.js example)
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET,PUT,POST,DELETE,OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  next();
});

// Or use cors middleware
const cors = require('cors');
app.use(cors({
  origin: '*',
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));
```

### Issue: 404 errors for assets

**Problem**: Incorrect asset paths or missing files

**Solution**:
1. Check build output directory structure
2. Verify asset paths in HTML
3. Configure web server to serve static files:
   ```nginx
   # Nginx configuration
   location /assets/ {
     alias /var/www/primeexcelence-addin/dist/assets/;
     expires 1y;
   }
   ```

## Performance Issues

### Issue: Slow loading times

**Problem**: Large bundle size or inefficient loading

**Solution**:
```bash
# Analyze bundle size
npm install --save-dev webpack-bundle-analyzer
npx webpack-bundle-analyzer dist/

# Optimize bundle
# - Code splitting
# - Tree shaking
# - Minification
# - Compression
```

### Issue: Memory leaks

**Problem**: Improper cleanup of event listeners or timers

**Solution**:
```typescript
// React component cleanup
useEffect(() => {
  const handleResize = () => {
    // Handle resize
  };
  
  window.addEventListener('resize', handleResize);
  
  return () => {
    window.removeEventListener('resize', handleResize);
  };
}, []);

// Office.js cleanup
Office.onReady(() => {
  // Set up event handlers
  
  // Clean up on unload
  window.addEventListener('beforeunload', () => {
    // Clean up resources
  });
});
```

## Debugging Tips

### Enable Debug Mode

1. **Browser DevTools**
   - Open F12 Developer Tools
   - Check Console, Network, and Sources tabs
   - Set breakpoints in TypeScript code

2. **Office.js Debugging**
   ```typescript
   // Enable Office.js runtime logging
   Office.context.diagnostics.log = true;
   
   // Log Office.js events
   Office.onReady((info) => {
     console.log('Office.js ready:', info);
   });
   ```

3. **Webpack Dev Server Debugging**
   ```bash
   # Enable verbose logging
   npm start -- --log-level verbose
   
   # Debug webpack configuration
   npm start -- --debug
   ```

### Common Debug Commands

```bash
# Check if development server is running
curl -I https://localhost:3000

# Test manifest accessibility
curl -I https://localhost:3000/manifest.xml

# Validate manifest
office-addin-validator manifest.xml

# Check certificate
openssl s_client -connect localhost:3000 -servername localhost

# Test Office.js loading
curl https://localhost:3000/taskpane.html | grep -i office

# Check for JavaScript errors
npm run build 2>&1 | grep -i error
```

### Log Collection

1. **Browser Console Logs**
   - Right-click in console → "Save as..."
   - Or copy logs to clipboard

2. **Excel Application Logs**
   - Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
   - Mac: `~/Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`

3. **Web Server Logs**
   ```bash
   # Nginx
   tail -f /var/log/nginx/access.log
   tail -f /var/log/nginx/error.log
   
   # Apache
   tail -f /var/log/apache2/access.log
   tail -f /var/log/apache2/error.log
   ```

### Testing Checklist

Before deploying to production:

- [ ] Manifest validates successfully
- [ ] All URLs are accessible over HTTPS
- [ ] SSL certificate is valid and trusted
- [ ] Icons display correctly
- [ ] Addin loads in Excel desktop
- [ ] Addin loads in Excel Online
- [ ] All features work as expected
- [ ] No console errors
- [ ] Performance is acceptable
- [ ] CORS is properly configured
- [ ] Error handling works correctly

### Getting Help

If you're still experiencing issues:

1. Check the GitHub repository issues
2. Review Office.js documentation
3. Use Stack Overflow with tags: `office-js`, `excel-web-addin`
4. Contact Microsoft support for Office.js issues
5. Check browser compatibility

## Additional Resources

- [Office.js Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- [Excel JavaScript API Reference](https://docs.microsoft.com/en-us/javascript/api/excel)
- [Office Add-ins Troubleshooting](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-development-errors)
- [Webpack Documentation](https://webpack.js.org/concepts/)
- [React Documentation](https://reactjs.org/docs/getting-started.html)
- [Material-UI Documentation](https://mui.com/getting-started/installation/)

---

This troubleshooting guide should help resolve most common issues. If you encounter a problem not covered here, please consider contributing to improve this guide.