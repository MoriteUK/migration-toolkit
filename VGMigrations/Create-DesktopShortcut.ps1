#Requires -Version 7.0
<#
.SYNOPSIS
    Creates a desktop shortcut to launch Migration Toolkit
.DESCRIPTION
    Creates a Windows shortcut on the desktop that launches the Migration Toolkit main menu
#>

$ErrorActionPreference = 'Stop'

# Determine paths
$toolkitPath = $PSScriptRoot
$mainMenuScript = Join-Path $toolkitPath 'main-menu.ps1'

if (-not (Test-Path $mainMenuScript)) {
    Write-Host "ERROR: main-menu.ps1 not found at: $mainMenuScript" -ForegroundColor Red
    exit 1
}

# Get desktop path
$desktopPath = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktopPath 'Migration Toolkit.lnk'

# Find PowerShell 7+ executable
$pwshPath = (Get-Command pwsh -ErrorAction SilentlyContinue).Source
if (-not $pwshPath) {
    Write-Host "ERROR: PowerShell 7+ (pwsh.exe) not found. Please install PowerShell 7 or later." -ForegroundColor Red
    exit 1
}

# Create the shortcut
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($shortcutPath)
$Shortcut.TargetPath = $pwshPath
$Shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$mainMenuScript`""
$Shortcut.WorkingDirectory = $toolkitPath
$Shortcut.Description = "Migration Toolkit - M365 Migration Tools"

# Set icon if available
$iconPath = Join-Path $toolkitPath 'FlyMigration.ico'
if (Test-Path $iconPath) {
    $Shortcut.IconLocation = $iconPath
}

$Shortcut.Save()

Write-Host ""
Write-Host "Desktop shortcut created successfully!" -ForegroundColor Green
Write-Host ""
Write-Host "Shortcut location: $shortcutPath" -ForegroundColor Cyan
Write-Host "Target: $mainMenuScript" -ForegroundColor Cyan
Write-Host ""
Write-Host "You can now launch Migration Toolkit from your desktop." -ForegroundColor Yellow
Write-Host ""

# Offer to launch now
$response = Read-Host "Would you like to launch Migration Toolkit now? (Y/N)"
if ($response -eq 'Y' -or $response -eq 'y') {
    & $shortcutPath
}
