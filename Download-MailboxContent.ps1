<#
.SYNOPSIS
    Grant admin access to a user's mailbox and export all content to PST file

.DESCRIPTION
    This script connects to Exchange Online, grants necessary permissions,
    and exports a user's mailbox to a PST file on a network share or local path.

.PARAMETER UserEmail
    The email address of the user whose mailbox you want to export

.PARAMETER AdminEmail
    The admin email address to grant access to (typically your admin account)

.PARAMETER ExportPath
    UNC path where PST file will be saved (e.g., \\server\share\mailboxbackups)
    Note: Must be a network share accessible by Exchange Online

.PARAMETER PSTFileName
    Name for the PST file (without extension)

.EXAMPLE
    .\Download-MailboxContent.ps1 -UserEmail "user@domain.com" -AdminEmail "admin@domain.com" -ExportPath "\\fileserver\mailboxexports" -PSTFileName "user_backup"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$UserEmail,

    [Parameter(Mandatory=$true)]
    [string]$AdminEmail,

    [Parameter(Mandatory=$true)]
    [string]$ExportPath,

    [Parameter(Mandatory=$true)]
    [string]$PSTFileName
)

# Check if Exchange Online Management module is installed
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Installing ExchangeOnlineManagement module..." -ForegroundColor Yellow
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
}

Import-Module ExchangeOnlineManagement

Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan

try {
    # Connect to Exchange Online
    Connect-ExchangeOnline -UserPrincipalName $AdminEmail -ShowBanner:$false

    # Grant Full Access permission to the mailbox
    Write-Host "Granting Full Access permission to $AdminEmail on mailbox: $UserEmail" -ForegroundColor Cyan
    Add-MailboxPermission -Identity $UserEmail -User $AdminEmail -AccessRights FullAccess -InheritanceType All -AutoMapping $false

    # Check if admin has Mailbox Import Export role
    Write-Host "Checking role assignments..." -ForegroundColor Cyan
    $roleAssignment = Get-ManagementRoleAssignment -RoleAssignee $AdminEmail -Role "Mailbox Import Export" -ErrorAction SilentlyContinue

    if (-not $roleAssignment) {
        Write-Host "WARNING: You need 'Mailbox Import Export' role to perform exports." -ForegroundColor Yellow
        Write-Host "Attempting to assign role..." -ForegroundColor Yellow

        try {
            New-ManagementRoleAssignment -Role "Mailbox Import Export" -User $AdminEmail
            Write-Host "Role assigned successfully. Please wait 15-30 minutes for the role to propagate, then run this script again." -ForegroundColor Green
            Write-Host "You may also need to close and reopen your PowerShell session." -ForegroundColor Yellow
            Disconnect-ExchangeOnline -Confirm:$false
            exit 0
        }
        catch {
            Write-Host "ERROR: Failed to assign role. Please have a Global Admin assign the 'Mailbox Import Export' role manually." -ForegroundColor Red
            Write-Host "Command: New-ManagementRoleAssignment -Role 'Mailbox Import Export' -User $AdminEmail" -ForegroundColor Yellow
            Disconnect-ExchangeOnline -Confirm:$false
            exit 1
        }
    }

    # Ensure export path ends with backslash
    if (-not $ExportPath.EndsWith('\')) {
        $ExportPath += '\'
    }

    # Create full PST path
    $fullPSTPath = "$ExportPath$PSTFileName.pst"

    Write-Host "`nStarting mailbox export..." -ForegroundColor Cyan
    Write-Host "Source Mailbox: $UserEmail" -ForegroundColor Green
    Write-Host "Destination: $fullPSTPath" -ForegroundColor Green
    Write-Host "`nNOTE: The export path must be a network share (UNC path) accessible by Exchange Online." -ForegroundColor Yellow

    # Create the export request
    $exportRequest = New-MailboxExportRequest -Mailbox $UserEmail -FilePath $fullPSTPath -Name "$PSTFileName-$(Get-Date -Format 'yyyyMMdd-HHmmss')"

    Write-Host "`nExport request created successfully!" -ForegroundColor Green
    Write-Host "Request Name: $($exportRequest.Name)" -ForegroundColor Cyan
    Write-Host "Status: $($exportRequest.Status)" -ForegroundColor Cyan

    Write-Host "`nMonitoring export progress..." -ForegroundColor Yellow
    Write-Host "This may take several minutes to hours depending on mailbox size." -ForegroundColor Yellow
    Write-Host "Press Ctrl+C to stop monitoring (export will continue in background)`n" -ForegroundColor Yellow

    # Monitor progress
    do {
        Start-Sleep -Seconds 30
        $status = Get-MailboxExportRequest -Identity $exportRequest.Identity

        $percentComplete = if ($status.PercentComplete) { $status.PercentComplete } else { 0 }
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Status: $($status.Status) | Progress: $percentComplete% | Items Transferred: $($status.ItemsTransferred)" -ForegroundColor Cyan

        if ($status.Status -eq "Failed") {
            Write-Host "`nExport FAILED!" -ForegroundColor Red
            Write-Host "Error: $($status.Message)" -ForegroundColor Red
            break
        }

    } while ($status.Status -ne "Completed")

    if ($status.Status -eq "Completed") {
        Write-Host "`nExport completed successfully!" -ForegroundColor Green
        Write-Host "Items Transferred: $($status.ItemsTransferred)" -ForegroundColor Cyan
        Write-Host "Bytes Transferred: $([math]::Round($status.BytesTransferred / 1MB, 2)) MB" -ForegroundColor Cyan
        Write-Host "PST File Location: $fullPSTPath" -ForegroundColor Green

        # Clean up export request
        $cleanup = Read-Host "`nDo you want to remove the export request from Exchange Online? (Y/N)"
        if ($cleanup -eq 'Y') {
            Remove-MailboxExportRequest -Identity $exportRequest.Identity -Confirm:$false
            Write-Host "Export request removed." -ForegroundColor Green
        }
    }

    # Optional: Remove mailbox permissions
    $removeAccess = Read-Host "`nDo you want to remove Full Access permission from $AdminEmail? (Y/N)"
    if ($removeAccess -eq 'Y') {
        Remove-MailboxPermission -Identity $UserEmail -User $AdminEmail -AccessRights FullAccess -Confirm:$false
        Write-Host "Mailbox permissions removed." -ForegroundColor Green
    }

    Disconnect-ExchangeOnline -Confirm:$false
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 1
}
