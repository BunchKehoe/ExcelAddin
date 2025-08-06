# Production Build Guide

This guide explains how to handle production builds for the Excel Add-in, particularly resolving issues with `custom-functions-metadata` command not found during production deployments.

## The Issue

When running `npm install --production` followed by `npm run build`, the build fails with:
```
Der Befehl "custom-functions-metadata" ist entweder falsch geschrieben oder konnte nicht gefunden werden.
```

This occurs because `custom-functions-metadata` is in `devDependencies`, but the build process requires it to generate `functions.json` from TypeScript source.

## Solution: Pre-Generated Functions Metadata

The solution is to pre-generate the `functions.json` file during development and include it in the repository, eliminating the need for `custom-functions-metadata` during production builds.

## Build Strategies

### Strategy 1: Full Build Environment (Recommended)
For production deployments, maintain a build environment with full dependencies:

```bash
# On build server or CI/CD pipeline
npm install                    # Install all dependencies including devDependencies
npm run build:staging          # Build with function generation
# Deploy dist/ folder to production server
```

### Strategy 2: Pre-Built Artifacts
For environments where you cannot install devDependencies:

1. **Development/Build Phase:**
   ```bash
   npm install                           # Full install with devDependencies
   npm run build:functions              # Generate functions.json
   git add src/commands/functions.json  # Commit the generated file
   npm run build:staging                # Create production build
   ```

2. **Production Deployment:**
   - Copy the pre-built `dist/` folder to your production server
   - No npm install needed on production server

### Strategy 3: Hybrid Approach
For cases where you must install packages on production server:

```bash
# Production server
npm install --production        # Install only production dependencies
# Copy pre-built dist/ folder from build environment
# Or include dist/ in your repository (not recommended for large builds)
```

## Updated Build Scripts

The package.json has been updated with production-friendly build commands:

```json
{
  "scripts": {
    "build": "npm run build:functions && webpack --mode production",
    "build:production": "webpack --config webpack.prod.config.js --mode production",
    "build:production:staging": "webpack --config webpack.prod.config.js --mode production --env staging",
    "deploy:prepare:production": "npm run clean && npm install --production && npm run build:production"
  }
}
```

**Note:** The `build:production` commands will fail without devDependencies because webpack itself is a devDependency. These are provided for environments where devDependencies might be selectively installed.

## Production Deployment Workflow

### Option A: CI/CD Pipeline (Recommended)
```yaml
# Example GitHub Actions or similar
- name: Install dependencies
  run: npm install
- name: Build application  
  run: npm run build:staging
- name: Deploy artifacts
  run: rsync -av dist/ user@server:/path/to/deployment/
```

### Option B: Local Build, Remote Deploy
```bash
# Local development machine
npm install
npm run build:staging
scp -r dist/* user@server:/path/to/deployment/

# No npm install needed on production server
```

### Option C: Self-Contained Deployment Package
```bash
# Create deployment package with pre-built assets
npm install
npm run build:staging
tar -czf excel-addin-deployment.tar.gz dist/
# Transfer and extract tar.gz on production server
```

## Functions.json Management

The `src/commands/functions.json` file is:
- ✅ Pre-generated from TypeScript source during development
- ✅ Committed to the git repository  
- ✅ Used at runtime by Excel to register custom functions
- ✅ Copied to `dist/functions.json` during webpack build

### When to Regenerate
Run `npm run build:functions` when you:
- Add new custom functions to `src/commands/commands.ts`
- Modify function parameters or return types
- Change function descriptions or metadata

## Troubleshooting

### Error: "custom-functions-metadata command not found"
**Cause:** Trying to run builds that require devDependencies with production-only install.

**Solutions:**
1. Use a proper CI/CD pipeline with full dependency install
2. Pre-build on development machine and deploy artifacts
3. Install specific devDependencies needed: `npm install custom-functions-metadata webpack webpack-cli --save-dev`

### Error: "Cannot find module 'html-webpack-plugin'"
**Cause:** Webpack and its plugins are devDependencies, not available in production-only install.

**Solutions:**
1. Deploy pre-built `dist/` folder (recommended)
2. Use full `npm install` on build environment
3. Consider using webpack as a production dependency if builds must happen on production server (not recommended)

### Functions not working in Excel
1. Verify `functions.json` exists in `dist/` folder
2. Check manifest.xml references correct functions.json URL
3. Ensure functions.json is accessible via web server
4. Verify custom function registration in Excel developer tools

## Best Practices

1. **Separation of Concerns:** Build on dedicated build environment, deploy artifacts to production
2. **Version Control:** Include `functions.json` in repository for reproducible builds
3. **Automation:** Use CI/CD pipelines for consistent builds and deployments
4. **Validation:** Test built artifacts before deployment
5. **Rollback:** Keep previous build artifacts for quick rollback if needed

## File Structure

```
src/commands/
├── commands.ts        # TypeScript custom function definitions  
├── functions.json     # Generated metadata (committed to git)
└── commands.html      # Custom functions page

dist/                  # Build output (generated)
├── functions.json     # Copied from src/commands/
├── taskpane.html      # Main add-in page
├── commands.html      # Custom functions page  
└── *.js              # Compiled and bundled JavaScript
```