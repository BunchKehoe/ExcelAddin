# Add Windows Firewall rule for nginx HTTPS port
# Usage: .\add-firewall-rule.ps1
# Requires administrator privileges

param(
    [int]$Port = 9443,
    [string]$RuleName = "Excel Add-in nginx HTTPS"
)

Write-Host "🔧 Adding Windows Firewall rule for port $Port..." -ForegroundColor Cyan

# Check if running as administrator
$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]$currentUser
$isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "❌ This script must be run as Administrator!" -ForegroundColor Red
    Write-Host "💡 Right-click PowerShell and select 'Run as Administrator'" -ForegroundColor Yellow
    exit 1
}

try {
    # Check if rule already exists
    $existingRule = Get-NetFirewallRule -DisplayName $RuleName -ErrorAction SilentlyContinue
    
    if ($existingRule) {
        Write-Host "⚠️  Firewall rule '$RuleName' already exists" -ForegroundColor Yellow
        Write-Host "Removing existing rule..." -ForegroundColor Yellow
        Remove-NetFirewallRule -DisplayName $RuleName
    }
    
    # Create new firewall rule
    New-NetFirewallRule -DisplayName $RuleName -Direction Inbound -Protocol TCP -LocalPort $Port -Action Allow
    
    Write-Host "✅ Successfully added firewall rule:" -ForegroundColor Green
    Write-Host "   Name: $RuleName" -ForegroundColor White
    Write-Host "   Port: $Port" -ForegroundColor White
    Write-Host "   Direction: Inbound" -ForegroundColor White
    Write-Host "   Action: Allow" -ForegroundColor White
    
    # Verify the rule was created
    $newRule = Get-NetFirewallRule -DisplayName $RuleName -ErrorAction SilentlyContinue
    if ($newRule) {
        Write-Host "✅ Firewall rule verified successfully!" -ForegroundColor Green
    }
    
} catch {
    Write-Host "❌ Error adding firewall rule: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "💡 Try running this command manually:" -ForegroundColor Yellow
    Write-Host "New-NetFirewallRule -DisplayName '$RuleName' -Direction Inbound -Protocol TCP -LocalPort $Port -Action Allow" -ForegroundColor Cyan
    exit 1
}

Write-Host ""
Write-Host "🎯 Firewall configuration complete!" -ForegroundColor Cyan
Write-Host "You can now test connectivity to port $Port" -ForegroundColor Green