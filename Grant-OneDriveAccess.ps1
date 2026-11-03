<#
.SYNOPSIS
    Grant yourself permanent admin access to a user's OneDrive

.DESCRIPTION
    This script grants site collection admin rights to a specified admin account
    for a user's OneDrive site. After running this, wait 5-10 minutes before
    attempting to access the OneDrive.

.PARAMETER UserEmail
    The email address of the user whose OneDrive you want to access

.PARAMETER AdminEmail
    The admin email address to grant access to

.PARAMETER TenantName
    Your SharePoint tenant name (e.g., 'theivycollection')

.EXAMPLE
    .\Grant-OneDriveAccess.ps1 -UserEmail "user@domain.com" -AdminEmail "admin@domain.com" -TenantName "theivycollection"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$UserEmail,

    [Parameter(Mandatory=$true)]
    [string]$AdminEmail,

    [Parameter(Mandatory=$true)]
    [string]$TenantName
)

# Check if PnP.PowerShell module is installed
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "Installing PnP.PowerShell module..." -ForegroundColor Yellow
    Install-Module -Name PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
}

Import-Module PnP.PowerShell

# Clean up TenantName
if ($TenantName -match 'https?://(.+?)(?:-admin)?\.sharepoint\.com') {
    $TenantName = $Matches[1]
}
$TenantName = $TenantName.Trim()

# Convert email to OneDrive URL format
$userFormatted = $UserEmail -replace '@', '_' -replace '\.', '_'
$oneDriveUrl = "https://$TenantName-my.sharepoint.com/personal/$userFormatted"
$adminUrl = "https://$TenantName-admin.sharepoint.com"

Write-Host "=== Grant OneDrive Access ===" -ForegroundColor Cyan
Write-Host "User Email: $UserEmail" -ForegroundColor White
Write-Host "Admin Email: $AdminEmail" -ForegroundColor White
Write-Host "OneDrive URL: $oneDriveUrl" -ForegroundColor Yellow
Write-Host ""

try {
    Write-Host "Connecting to SharePoint Admin Center..." -ForegroundColor Cyan
    Connect-PnPOnline -Url $adminUrl -Interactive

    Write-Host "Granting site collection admin access to $AdminEmail..." -ForegroundColor Cyan
    Set-PnPTenantSite -Identity $oneDriveUrl -Owners $AdminEmail

    Write-Host ""
    Write-Host "SUCCESS! Admin access granted." -ForegroundColor Green
    Write-Host ""
    Write-Host "IMPORTANT: Please wait 5-10 minutes for permissions to fully propagate" -ForegroundColor Yellow
    Write-Host "across all SharePoint services before attempting to access the OneDrive." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "After waiting, you can access the OneDrive at:" -ForegroundColor White
    Write-Host $oneDriveUrl -ForegroundColor Cyan
    Write-Host ""
    Write-Host "To download files, run:" -ForegroundColor White
    Write-Host ".\Download-OneDrive-Simple.ps1 -OneDriveUrl '$oneDriveUrl' -DestinationPath 'C:\temp'" -ForegroundColor Cyan

    Disconnect-PnPOnline
}
catch {
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
    exit 1
}
