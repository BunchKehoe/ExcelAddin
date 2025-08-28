# Excel Custom Functions Setup Guide

## Overview
This guide covers the setup and architecture of Excel Custom Functions for the Prime Capital add-in.

## Key Architecture Fix

### The Problem
Excel Custom Functions were not loading because:
1. **ES Module Incompatibility**: Excel Custom Functions runtime cannot load ES modules
2. **Dual Implementation Conflict**: Functions were defined in both TypeScript (ES modules) and plain JavaScript
3. **Incorrect Runtime Context**: Functions need to run in Excel's separate Custom Functions runtime, not the task pane context

### The Solution
Created dedicated plain JavaScript custom functions that are compatible with Excel's Custom Functions runtime:

## File Structure
```
/public/
├── customfunctions.js      # Plain JavaScript functions (NO ES modules)
├── customfunctions.html    # HTML page for Custom Functions runtime
└── functions.json          # Function metadata for Excel

/src/commands/
└── commands.ts            # Ribbon commands only (NOT custom functions)
```

## Key Requirements

### 1. Plain JavaScript Only
Excel Custom Functions **cannot use ES modules**:
- ✅ `function IRR() { ... }`
- ✅ `window.IRR = IRR`
- ✅ `CustomFunctions.associate('PC.IRR', IRR)`
- ❌ `export function IRR() { ... }`
- ❌ `import { ... } from '...'`

### 2. Separate Runtime Context
- **Custom Functions**: Run in Excel's separate Custom Functions runtime
- **Ribbon Commands**: Run in task pane context with full ES module support
- These are completely separate and cannot share code directly

### 3. Function Registration
Functions must be registered with Excel's Custom Functions runtime:
```javascript
Office.onReady(() => {
  CustomFunctions.associate('PC.IRR', IRR);
  CustomFunctions.associate('PC.JOINCELLS', JOINCELLS);
});
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

### 2. Start Development Server
```bash
npm run dev
```
Server runs at `https://localhost:3000`

### 3. Load Add-in in Excel
1. Open Excel
2. Insert → Add-ins → Upload My Add-in
3. Select `public/manifest-local.xml`
4. Functions will be available as `=PC.IRR()` and `=PC.JOINCELLS()`

## Verification

### Check Endpoints
All endpoints should be accessible in browser:
- `https://localhost:3000/customfunctions.html` - Custom Functions runtime page
- `https://localhost:3000/customfunctions.js` - Function implementations
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

## Manifest Configuration

### Local Development
```xml
<bt:Url id="CustomFunctions.Script.Url" DefaultValue="https://localhost:3000/customfunctions.js"/>
<bt:Url id="CustomFunctions.Page.Url" DefaultValue="https://localhost:3000/customfunctions.html"/>
```

### Production/Staging
Manifests automatically point to correct server URLs with `/excellence/` base path.

## Building for Production

```bash
npm run build
```

Build process copies custom functions files from `/public` to `/dist`:
- `dist/customfunctions.js`
- `dist/customfunctions.html`
- `dist/functions.json`

## Troubleshooting

### Functions Not Loading
1. Check Excel Developer Console for errors
2. Verify custom functions endpoints are accessible in browser
3. Ensure no ES module syntax in `customfunctions.js`
4. Check that `CustomFunctions.associate()` is being called

### SSL Certificate Issues
```bash
npm run cert:install
```

### Function Updates
1. Restart Excel completely
2. Reload the add-in
3. Clear Excel's function cache if needed

## Technical Notes

### Why Plain JavaScript?
- Excel Custom Functions runtime is separate from task pane context
- This runtime cannot load ES modules
- Plain JavaScript ensures maximum compatibility
- Vite handles task pane complexity separately

### Function Isolation
- Custom Functions: `/public/customfunctions.js` (plain JS)
- Ribbon Commands: `/src/commands/commands.ts` (TypeScript with ES modules)
- These run in completely separate contexts and cannot share code directly