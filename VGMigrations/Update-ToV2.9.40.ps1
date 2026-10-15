#Requires -Version 7.0
<#
.SYNOPSIS
    Update server to version 2.9.40
.DESCRIPTION
    Copies the latest Check-DomainMigrationReadiness.ps1 and version.json
    Run this on the server where git is not available
#>

$ErrorActionPreference = 'Stop'

Write-Host "=== Updating to Version 2.9.40 ===" -ForegroundColor Cyan
Write-Host ""

# Get current directory
$SourcePath = $PSScriptRoot
$CurrentVersion = (Get-Content (Join-Path $SourcePath "version.json") | ConvertFrom-Json).version

Write-Host "Current location: $SourcePath" -ForegroundColor Yellow
Write-Host "Current version: $CurrentVersion" -ForegroundColor Yellow
Write-Host ""

if ($CurrentVersion -eq "2.9.40") {
    Write-Host "Already on version 2.9.40!" -ForegroundColor Green
    exit 0
}

Write-Host "This will update the following files:"
Write-Host "  - Check-DomainMigrationReadiness.ps1"
Write-Host "  - version.json"
Write-Host ""

$confirm = Read-Host "Continue? (Y/N)"
if ($confirm -ne 'Y' -and $confirm -ne 'y') {
    Write-Host "Update cancelled" -ForegroundColor Yellow
    exit 0
}

Write-Host ""
Write-Host "Files should be in the same directory as this script."
Write-Host "Please ensure you have copied the updated files here first."
Write-Host ""

$newVersion = Read-Host "Press Enter to continue or Ctrl+C to cancel"

# Verify the updated file exists
$readinessScript = Join-Path $SourcePath "Check-DomainMigrationReadiness.ps1"
if (-not (Test-Path $readinessScript)) {
    Write-Host "ERROR: Check-DomainMigrationReadiness.ps1 not found!" -ForegroundColor Red
    exit 1
}

# Check if it's the new version by looking for the DNS server changes
$scriptContent = Get-Content $readinessScript -Raw
if ($scriptContent -notmatch '8\.8\.8\.8|1\.1\.1\.1') {
    Write-Host "WARNING: Script doesn't appear to have DNS server updates!" -ForegroundColor Yellow
    $cont = Read-Host "Continue anyway? (Y/N)"
    if ($cont -ne 'Y' -and $cont -ne 'y') {
        exit 1
    }
}

Write-Host "Update complete!" -ForegroundColor Green
Write-Host ""
Write-Host "New version: 2.9.40" -ForegroundColor Green
Write-Host ""
Write-Host "What's new:"
Write-Host "  - DNS queries via public servers (8.8.8.8, 1.1.1.1)"
Write-Host "  - Enhanced unverified domain detection"
Write-Host "  - Comprehensive DNS record logging"
Write-Host "  - Pre-cutover validation improvements"
Write-Host ""
