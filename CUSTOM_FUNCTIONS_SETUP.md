# Excel Custom Functions Setup Guide

## Overview
This guide covers the setup and architecture of Excel Custom Functions for the Prime Capital add-in with shared runtime support.

## Updated Architecture (Shared Runtime)

### The Solution
Custom functions now support both standalone and shared runtime execution:

1. **Exported TypeScript Functions**: Located in `src/functions/functions.ts` with proper exports
2. **Shared Runtime Support**: Functions are available in both custom functions runtime AND taskpane runtime  
3. **Build Process**: Automatically generates plain JavaScript versions for Excel compatibility

## File Structure
```
./src/functions/
├── functions.ts             # TypeScript source with exports (shared runtime)
├── functions.html           # HTML page for Custom Functions runtime
└── functions.json           # Function metadata for Excel

./public/ (generated)
├── customfunctions.js       # Auto-generated plain JavaScript (standalone runtime)
├── customfunctions.html     # HTML page for Custom Functions runtime
├── functions.js             # Alternative plain JavaScript version
├── functions.html           # Generated HTML page
└── functions.json           # Generated function metadata
```

## Key Changes

### 1. TypeScript Source with Exports
Custom functions are now properly exported from TypeScript:
```typescript
export function IRR(cell1: number, cell2: number, ...): number {
  // Implementation
}

export function JOINCELLS(range: any[][], delimiter: string = ", "): string {
  // Implementation  
}
```

### 2. Shared Runtime Support
- Functions are available in both custom functions runtime AND taskpane context
- CustomFunctions.associate() is called in both runtimes for maximum compatibility
- Taskpane can access custom functions through imports

### 3. Automatic Build Process
- Build script removes exports and type annotations for standalone runtime
- Generates plain JavaScript in `/public` automatically
- Maintains TypeScript source for shared runtime functionality

## Function Registration

### Shared Runtime (TypeScript)
```typescript
// In functions.ts and taskpane.tsx
if (typeof CustomFunctions !== 'undefined') {
  CustomFunctions.associate("PC.IRR", IRR);
  CustomFunctions.associate("PC.JOINCELLS", JOINCELLS);
}
```

### Standalone Runtime (Generated JavaScript)
```javascript
// In generated customfunctions.js
if (typeof CustomFunctions !== 'undefined') {
  CustomFunctions.associate("PC.IRR", IRR);
  CustomFunctions.associate("PC.JOINCELLS", JOINCELLS);
}
```

## Available Functions

### PC.IRR
Takes five cell values and returns their sum.
```excel
=PC.IRR(1,2,3,4,5)  # Returns 15
```

### PC.JOINCELLS
Joins cell ranges with a specified delimiter.
```excel
=PC.JOINCELLS(A1:A3,"; ")  # Joins cells A1-A3 with semicolon
```

## Development Setup

### 1. Install Dependencies and Certificates
```bash
npm install
npm run cert:install
```

### 2. Build Custom Functions
```bash
node build-customfunctions.cjs
# OR
npm run build:dev
```

### 3. Start Development Server
```bash
npm run dev
```
Server runs at `https://localhost:3000`

### 4. Load Add-in in Excel
1. Open Excel
2. Insert → Add-ins → Upload My Add-in
3. Select `public/manifest-local.xml`
4. Functions will be available as `=PC.IRR()` and `=PC.JOINCELLS()`

## Build Process

### Custom Functions Build
The `build-customfunctions.cjs` script:
1. Reads TypeScript source from `src/functions/functions.ts`
2. Removes exports and type annotations
3. Generates plain JavaScript in `public/customfunctions.js`
4. Copies metadata and HTML files

### Development Build
```bash
npm run build:dev  # Includes custom functions build
```

### Production Build  
```bash
npm run build:prod  # Includes custom functions build
```

## Verification

### Check Endpoints
All endpoints should be accessible in browser:
- `https://localhost:3000/customfunctions.html` - Custom Functions runtime page
- `https://localhost:3000/customfunctions.js` - Generated function implementations
- `https://localhost:3000/functions.json` - Function metadata

### Test in Excel
```excel
=PC.IRR(1,2,3,4,5)        // Should return 15
=PC.JOINCELLS(A1:A3,"; ") // Should join cells with semicolon
```

### Debug Console
Open Excel Developer Tools (F12) to see:
- Function execution logs
- Registration success/error messages
- Parameter validation results
- Shared runtime registration confirmations

## Manifest Configuration

### Local Development
```xml
<bt:Url id="CustomFunctions.Script.Url" DefaultValue="https://localhost:3000/customfunctions.js"/>
<bt:Url id="CustomFunctions.Page.Url" DefaultValue="https://localhost:3000/customfunctions.html"/>
<bt:Url id="Functions.Metadata.Url" DefaultValue="https://localhost:3000/functions.json"/>
```

### Shared Runtime Configuration
```xml
<Runtimes>
  <Runtime resid="CustomFunctions.Page.Url" lifetime="long" />
</Runtimes>
```

## Troubleshooting

### Functions Not Loading
1. Check Excel Developer Console for errors
2. Verify custom functions endpoints are accessible in browser
3. Ensure build process completed successfully
4. Check that `CustomFunctions.associate()` is being called in both runtimes
5. Verify shared runtime is properly configured in manifest

### Build Issues
```bash
node build-customfunctions.cjs  # Rebuild custom functions
npm run lint                    # Check TypeScript compilation
```

### SSL Certificate Issues
```bash
npm run cert:install
```

### Function Updates
1. Rebuild custom functions: `node build-customfunctions.cjs`
2. Restart development server: `npm run dev`
3. Reload the add-in in Excel
4. Clear Excel's function cache if needed

## Technical Notes

### Shared Runtime Benefits
- Functions available in both custom functions and taskpane contexts
- Better debugging and development experience
- Consistent function behavior across runtimes
- Ability to share state between taskpane and custom functions

### TypeScript to JavaScript Conversion
- Build process automatically removes type annotations
- Exports are stripped for standalone runtime compatibility
- Source remains in TypeScript for better development experience
- Generated JavaScript is pure ES5 for maximum Excel compatibility