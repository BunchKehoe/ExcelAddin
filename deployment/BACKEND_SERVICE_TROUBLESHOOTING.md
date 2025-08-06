# Backend Service Startup Issue - Quick Fix Guide

## Problem: Service Fails to Start with No Logs

If your Excel Add-in Backend service fails to start with no error messages in logs, this is typically caused by:

1. **Incorrect Python executable path in NSSM**
2. **Missing Python dependencies** 
3. **Wrong working directory or file paths**
4. **Permission issues**

## Quick Fix Steps

### 1. Use the New Setup Script

Run the enhanced backend service setup script:

```powershell
# Run as Administrator
cd C:\path\to\ExcelAddin\deployment\scripts
.\setup-backend-service.ps1 -Force
```

This script will:
- Auto-detect the correct Python executable path
- Test the Python environment before installing the service
- Create comprehensive logging and debugging tools
- Set proper NSSM configuration with full paths

### 2. Run Diagnostics

If the service still fails, run the diagnostic script:

```powershell
.\diagnose-backend-service.ps1 -FixCommonIssues
```

This will:
- Check all service configuration
- Identify common issues
- Apply automatic fixes
- Show detailed error information

### 3. Manual Testing

Test the backend manually to isolate the issue:

```powershell
cd C:\inetpub\wwwroot\ExcelAddin\backend
debug-service.bat
```

This will run the service wrapper directly and show any Python import errors.

## Common Issues and Solutions

### Issue 1: Python Path Problem

**Symptom:** Service fails immediately, no logs
**Cause:** NSSM configured with "python" instead of full path

**Fix:**
```powershell
# Find correct Python path
where python
# Update NSSM (replace with your actual path)
nssm set ExcelAddinBackend Application "C:\Python39\python.exe"
```

### Issue 2: Missing Dependencies 

**Symptom:** Import errors in logs
**Cause:** Python packages not installed

**Fix:**
```powershell
cd C:\inetpub\wwwroot\ExcelAddin\backend
python -m pip install -r requirements.txt
```

### Issue 3: Permission Issues

**Symptom:** Access denied errors
**Cause:** Service account lacks permissions

**Fix:**
```powershell
# Grant permissions to log directory
icacls "C:\Logs\ExcelAddin" /grant "Network Service:(OI)(CI)F" /T
# Grant read permissions to backend directory  
icacls "C:\inetpub\wwwroot\ExcelAddin" /grant "Network Service:(OI)(CI)R" /T
```

### Issue 4: Working Directory Problems

**Symptom:** File not found errors
**Cause:** Service can't find app.py or other files

**Fix:**
```powershell
# Ensure working directory is set correctly
nssm set ExcelAddinBackend AppDirectory "C:\inetpub\wwwroot\ExcelAddin\backend"
```

## Debugging Commands

```powershell
# Check service status
Get-Service ExcelAddinBackend

# View NSSM configuration
nssm dump ExcelAddinBackend

# Check recent logs
Get-Content "C:\Logs\ExcelAddin\backend-service-stderr.log" -Tail 20

# Test Python environment
cd C:\inetpub\wwwroot\ExcelAddin\backend
python -c "import app; print('Success')"

# Check if port 5000 is in use
Get-NetTCPConnection -LocalPort 5000

# Start service with verbose logging
nssm set ExcelAddinBackend AppStdout "C:\Logs\ExcelAddin\debug-stdout.log"
nssm set ExcelAddinBackend AppStderr "C:\Logs\ExcelAddin\debug-stderr.log"  
Start-Service ExcelAddinBackend
```

## Service Management

```powershell
# Start service
Start-Service ExcelAddinBackend

# Stop service  
Stop-Service ExcelAddinBackend

# Restart service
Restart-Service ExcelAddinBackend

# Remove and reinstall service
nssm remove ExcelAddinBackend confirm
.\setup-backend-service.ps1 -Force

# Edit service configuration
nssm edit ExcelAddinBackend
```

## Log File Locations

- **Service Output:** `C:\Logs\ExcelAddin\backend-service-stdout.log`
- **Service Errors:** `C:\Logs\ExcelAddin\backend-service-stderr.log`  
- **Application Log:** `C:\Logs\ExcelAddin\backend-service.log`
- **Windows Event Log:** Event Viewer → Windows Logs → System (filter by source: Service Control Manager)

## Next Steps

If the issue persists after trying these fixes:

1. **Run full diagnostics:** `.\diagnose-backend-service.ps1 -TestManually`
2. **Check Windows Event Viewer** for additional service errors
3. **Verify Python installation** and dependencies
4. **Test backend independently** using `python run.py`
5. **Consider using a different Python installation** or virtual environment

The enhanced setup and diagnostic scripts provide much better error reporting and should help identify the root cause of startup failures.