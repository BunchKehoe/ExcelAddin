# Environment-Specific Asset Serving

This document explains how the Excel Add-in handles conditional image references based on deployment environment.

## Overview

The application automatically serves images from different paths depending on the deployment environment:

- **Development**: Images served from `/assets/images/`
- **Staging/Production**: Images served from `/excellence/assets/images/`

## Implementation

### 1. Environment Detection (`src/config/environment.ts`)

The system automatically detects the current environment based on hostname:

```typescript
// Development: localhost or 127.0.0.1
// Staging: server-vs81t.intranet.local  
// Production: server-vs84.intranet.local
```

### 2. Asset Base URL Configuration

Each environment has its own `assetBaseUrl`:

```typescript
development: { assetBaseUrl: '/assets' }
staging: { assetBaseUrl: '/excellence/assets' }
production: { assetBaseUrl: '/excellence/assets' }
```

### 3. Asset Utility Functions (`src/utils/assetUtils.ts`)

Use these utility functions for consistent asset URL generation:

```typescript
import { getLogoUrl, getIconUrl, getAssetUrl } from '../utils/assetUtils';

// Generate logo URLs
const logoUrl = getLogoUrl('PCAG_white_trans.svg');

// Generate icon URLs  
const iconUrl = getIconUrl('icon-32.png');

// Generate any asset URL
const customUrl = getAssetUrl('images/custom/my-image.png');
```

### 4. Server Configuration (`server.cjs`)

The Express server serves assets at both paths:

```javascript
// Development path: /assets -> dist/assets
app.use('/assets', express.static(path.join(__dirname, 'dist/assets')));

// Staging/Production path: /excellence -> dist (includes assets)
app.use('/excellence', express.static(path.join(__dirname, 'dist')));
```

## Development Tools

### Debug Panel (Development Only)

In development mode, a debug panel appears in the bottom-right corner showing:
- Current environment detection
- Asset base URL being used
- Test results for key assets
- Any loading errors

### Asset Verification Test

Run the test script to verify asset serving works in all environments:

```bash
node test-assets.cjs
```

Expected output shows ✅ SUCCESS for all asset paths.

## Usage Examples

### In React Components

```typescript
import { getLogoUrl } from '../utils/assetUtils';

// ✅ Recommended: Use utility functions
<img src={getLogoUrl('company-logo.svg')} alt="Logo" />

// ❌ Avoid: Manual path construction
<img src={`${assetBaseUrl}/images/logos/company-logo.svg`} alt="Logo" />
```

### Asset Organization

Place assets in the organized structure:

```
public/assets/images/
├── logos/           # Company branding
│   ├── PCAG_white_trans.svg
│   └── company-logo.png
├── icons/           # Excel add-in icons  
│   ├── icon-16.png
│   ├── icon-32.png
│   └── icon-80.png
└── ui/              # UI elements (if needed)
```

## Environment URLs

| Environment | Application URL | Asset URL Pattern |
|-------------|----------------|-------------------|
| **Development** | `https://localhost:3000` | `/assets/images/...` |
| **Staging** | `https://server-vs81t:9443/excellence/` | `/excellence/assets/images/...` |
| **Production** | `https://server-vs84:9443/excellence/` | `/excellence/assets/images/...` |

## Troubleshooting

### Images Not Loading

1. **Check Environment Detection**: Look at browser console for environment logs
2. **Verify Asset Paths**: Use the debug panel (development) or test script
3. **Check Build Output**: Ensure `npm run build` completed successfully
4. **Verify Server Configuration**: Test asset URLs directly in browser

### Debug Commands

```bash
# Verify assets are accessible
curl -I http://localhost:3000/assets/images/logos/PCAG_white_trans.svg
curl -I http://localhost:3000/excellence/assets/images/logos/PCAG_white_trans.svg

# Run comprehensive asset test
node test-assets.cjs
```

## Benefits

- ✅ **Automatic Environment Detection**: No manual configuration needed
- ✅ **Consistent API**: Same utility functions work in all environments  
- ✅ **Development Tools**: Built-in debugging and verification
- ✅ **Error Prevention**: Type-safe asset URL generation
- ✅ **Future-Proof**: Easy to add new environments or asset types