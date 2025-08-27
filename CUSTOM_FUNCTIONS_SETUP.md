# Excel Custom Functions Setup and Usage Guide

## Problem Resolved
The Excel Add-in custom functions were completely non-functional due to critical infrastructure issues:
- Functions weren't properly registered with Excel's Custom Functions runtime
- Incorrect manifest configuration pointing Script and Page to the same URL
- Missing dedicated custom functions files
- Complex Vite build pipeline interfering with Excel's runtime requirements

## Solution Implemented

### 1. Dedicated Custom Functions Architecture
Created a clean, dedicated custom functions setup that follows Microsoft's best practices:

- **`/public/customfunctions.js`** - Pure JavaScript functions without module complexity
- **`/public/customfunctions.html`** - Simple HTML page that loads the functions
- **Separate manifest URLs** - Script and Page now point to different resources as required

### 2. Fixed Manifest Configuration
Updated all manifest files (local, staging, production) with proper Custom Functions extension points:

```xml
<ExtensionPoint xsi:type="CustomFunctions">
  <Script>
    <SourceLocation resid="CustomFunctions.Script.Url"/>
  </Script>
  <Page>
    <SourceLocation resid="CustomFunctions.Page.Url"/>
  </Page>
  <Metadata>
    <SourceLocation resid="Functions.Metadata.Url"/>
  </Metadata>
  <Namespace resid="CustomFunctions.Namespace"/>
</ExtensionPoint>
```

### 3. Proper Function Registration
Functions are now registered with Excel's runtime using the correct API:

```javascript
// Register functions with CustomFunctions API
CustomFunctions.associate("PC.IRR", IRR);
CustomFunctions.associate("PC.JOINCELLS", JOINCELLS);
```

### 4. HTTPS Certificates Installed
Development certificates are now properly installed for Excel Add-in compatibility:
```bash
npm run cert:install
```

## Available Custom Functions

### PC.IRR Function
- **Purpose**: Takes five cell values and returns their sum
- **Usage**: `=PC.IRR(1,2,3,4,5)` → Returns `15`
- **Parameters**: Five numeric cell values
- **Validation**: Checks that all inputs are valid numbers

### PC.JOINCELLS Function  
- **Purpose**: Joins cells from a range with a specified delimiter
- **Usage**: `=PC.JOINCELLS(A1:A3,"; ")` → Joins A1, A2, A3 with semicolon
- **Parameters**: 
  - `range`: Cell range to join
  - `delimiter`: String to use as separator (optional, defaults to ", ")

## Development Setup

### 1. Install Dependencies and Certificates
```bash
npm install
npm run cert:install
```

### 2. Start Development Server
```bash
npm run dev
```
Server runs on `https://localhost:3000` with proper SSL certificates.

### 3. Load Add-in in Excel
1. Open Excel
2. Go to Insert → Add-ins → Upload My Add-in
3. Select `public/manifest-local.xml`
4. Functions will be available as `PC.IRR` and `PC.JOINCELLS`

## File Structure

### Custom Functions Files
- `/public/customfunctions.js` - Function implementations and registration
- `/public/customfunctions.html` - HTML page that loads the functions  
- `/public/functions.json` - Function metadata for Excel

### Manifest Files
- `/public/manifest-local.xml` - Local development (https://localhost:3000)
- `/public/manifest-staging.xml` - Staging environment
- `/public/manifest-prod.xml` - Production environment

## Verification Steps

### 1. Check Development Server
All endpoints should be accessible:
- `https://localhost:3000/customfunctions.js` - Function code
- `https://localhost:3000/customfunctions.html` - Function page
- `https://localhost:3000/functions.json` - Function metadata

### 2. Test Functions in Excel
```excel
=PC.IRR(1,2,3,4,5)        // Should return 15
=PC.JOINCELLS(A1:A3,"; ") // Should join cells with semicolon
```

### 3. Debug Console
Functions include console logging for debugging:
- Open Excel Developer Tools (F12)
- Check Console tab for function execution logs

## Technical Notes

- **No ES Modules**: Custom functions use traditional JavaScript to ensure Excel compatibility
- **Direct Registration**: Functions are registered immediately with `CustomFunctions.associate()`
- **Manifest Separation**: Script and Page URLs are now properly separated
- **Public Assets**: Functions are served as static assets via Vite's public directory

This implementation provides a solid foundation for Excel custom functions development with proper debugging and deployment support.
```
This installs SSL certificates required for HTTPS, which is mandatory for Excel Add-ins.

### 3. Build the Project
```bash
# Development build
npm run build:dev

# Or for production
npm run build:prod
```

### 4. Start Development Server
```bash
npm run dev
```
The server will start at `https://localhost:3000`

## Available Custom Functions

### PC.IRR
**Syntax**: `=PC.IRR(cell1, cell2, cell3, cell4, cell5)`
**Description**: Takes five numeric values and returns their sum
**Example**: `=PC.IRR(1, 2, 3, 4, 5)` returns `15`

### PC.JOINCELLS  
**Syntax**: `=PC.JOINCELLS(range, [delimiter])`
**Description**: Joins cells from a range into a single string with specified delimiter
**Example**: `=PC.JOINCELLS(A1:A3, "; ")` joins cells A1, A2, A3 with semicolon separator

## Testing the Add-in

### Local Development Testing

1. **Start the development server**:
   ```bash
   npm run dev
   ```

2. **Verify endpoints are working**:
   - Main taskpane: `https://localhost:3000/taskpane.html`
   - Custom functions: `https://localhost:3000/commands.html`
   - Functions metadata: `https://localhost:3000/functions.json`

3. **Load the manifest in Excel**:
   - Use `public/manifest-local.xml` for local development
   - Open Excel → Insert → Add-ins → Upload My Add-in
   - Select the manifest file

### Production Testing

1. **Build for production**:
   ```bash
   npm run build:prod
   ```

2. **Deploy to your server** and update manifest URLs accordingly

## Troubleshooting

### Common Issues

1. **HTTPS Certificate Errors**:
   - Run `npm run cert:install`
   - Restart browser after certificate installation

2. **Functions not appearing in Excel**:
   - Check console for JavaScript errors in the commands.html page
   - Verify functions.json is accessible
   - Ensure manifest URLs point to correct server

3. **"Function not found" errors**:
   - Verify the function names match the IDs in functions.json
   - Check that functions are properly exported to global scope
   - Ensure Office.js is loaded before custom functions

### Debug Console Output

The custom functions include console logging. Open Developer Tools in Excel to see:
- Function execution logs
- Parameter values
- Error messages

## File Structure

```
src/commands/
├── commands.ts          # Custom function implementations
├── functions.json       # Function metadata for Excel
public/
├── functions.json       # Copy of functions metadata
├── manifest-local.xml   # Development manifest
├── manifest-staging.xml # Staging manifest
└── manifest-prod.xml    # Production manifest
```

## Development Notes

- Functions are exported to global scope using `(globalThis as any).IRR = IRR`
- TypeScript validation ensures type safety
- All functions include input validation and error handling
- Console logging helps with debugging during development

## Next Steps

1. Test functions in Excel with the local manifest
2. Modify function implementations as needed
3. Update functions.json metadata if adding new functions
4. Deploy to staging/production servers for broader testing