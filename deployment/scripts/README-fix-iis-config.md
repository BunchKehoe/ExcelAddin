# IIS Configuration Fix

This script fixes common IIS web.config configuration issues that cause HTTP 500.19 errors.

## Problem
The most common issue when deploying to IIS is:
```
HTTP Error 500.19 - Internal Server Error
Config Error: This configuration section cannot be used at this path.
Error Code: 0x80070021
```

This happens when the web.config contains a `<handlers>` section, which is locked at the server level in IIS by default.

## Solution

Run the fix script as Administrator:
```powershell
.\deployment\scripts\fix-iis-config.ps1
```

## What it does

1. **Backs up** your current web.config to `web.config.backup`
2. **Detects** problematic configurations:
   - `<handlers>` section (causes 500.19 error)
   - Duplicate MIME type entries for .js, .json files
3. **Deploys** the latest compatible web.config
4. **Tests** the IIS configuration
5. **Restarts** the IIS site

## Options

```powershell
# Use custom site name
.\fix-iis-config.ps1 -SiteName "MyExcelApp"

# Force replacement even if no issues detected
.\fix-iis-config.ps1 -Force

# Both options
.\fix-iis-config.ps1 -SiteName "MyExcelApp" -Force
```

## Recovery

If something goes wrong, restore the backup:
```powershell
Copy-Item "C:\inetpub\wwwroot\ExcelAddin\web.config.backup" "C:\inetpub\wwwroot\ExcelAddin\web.config" -Force
```