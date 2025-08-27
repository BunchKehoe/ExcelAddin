# Excel Add-in Custom Functions Setup Guide

This guide provides step-by-step instructions for setting up and testing the Excel Add-in with custom functions.

## Prerequisites

1. Node.js (version 16 or higher)
2. npm
3. Excel (Desktop or Online)

## Quick Setup

### 1. Install Dependencies
```bash
npm install
```

### 2. Install Development Certificates
```bash
npm run cert:install
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