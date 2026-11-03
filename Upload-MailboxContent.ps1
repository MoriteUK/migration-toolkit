<#
.SYNOPSIS
    Upload PST file content to a target mailbox after migration

.DESCRIPTION
    This script connects to Exchange Online in the target tenant and imports
    a PST file into a user's mailbox.

.PARAMETER TargetUserEmail
    The email address of the target user in the destination tenant

.PARAMETER AdminEmail
    The admin email address in the destination tenant

.PARAMETER PSTPath
    UNC path to the PST file to import (e.g., \\server\share\mailboxbackups\user.pst)
    Note: Must be a network share accessible by Exchange Online

.PARAMETER ImportToArchive
    If specified, will import to the user's archive mailbox instead of primary

.EXAMPLE
    .\Upload-MailboxContent.ps1 -TargetUserEmail "user@newdomain.com" -AdminEmail "admin@newdomain.com" -PSTPath "\\fileserver\mailboxexports\user_backup.pst"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$TargetUserEmail,

    [Parameter(Mandatory=$true)]
    [string]$AdminEmail,

    [Parameter(Mandatory=$true)]
    [string]$PSTPath,

    [switch]$ImportToArchive
)

# Check if Exchange Online Management module is installed
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Installing ExchangeOnlineManagement module..." -ForegroundColor Yellow
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
}

Import-Module ExchangeOnlineManagement

# Verify PST path format (must be UNC path for Exchange Online)
if (-not $PSTPath.StartsWith('\\')) {
    Write-Host "ERROR: PST path must be a UNC path (e.g., \\server\share\file.pst)" -ForegroundColor Red
    Write-Host "Local paths are not supported for Exchange Online imports." -ForegroundColor Yellow
    exit 1
}

Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan

try {
    # Connect to Exchange Online
    Connect-ExchangeOnline -UserPrincipalName $AdminEmail -ShowBanner:$false

    # Verify target mailbox exists
    Write-Host "Verifying target mailbox exists: $TargetUserEmail" -ForegroundColor Cyan
    $mailbox = Get-Mailbox -Identity $TargetUserEmail -ErrorAction Stop
    Write-Host "Target mailbox found: $($mailbox.DisplayName)" -ForegroundColor Green

    # Check if admin has Mailbox Import Export role
    Write-Host "Checking role assignments..." -ForegroundColor Cyan
    $roleAssignment = Get-ManagementRoleAssignment -RoleAssignee $AdminEmail -Role "Mailbox Import Export" -ErrorAction SilentlyContinue

    if (-not $roleAssignment) {
        Write-Host "WARNING: You need 'Mailbox Import Export' role to perform imports." -ForegroundColor Yellow
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

    Write-Host "`nStarting mailbox import..." -ForegroundColor Cyan
    Write-Host "Source PST: $PSTPath" -ForegroundColor Green
    Write-Host "Target Mailbox: $TargetUserEmail" -ForegroundColor Green

    if ($ImportToArchive) {
        Write-Host "Destination: Archive Mailbox" -ForegroundColor Green
    } else {
        Write-Host "Destination: Primary Mailbox" -ForegroundColor Green
    }

    Write-Host "`nNOTE: The PST path must be a network share (UNC path) accessible by Exchange Online." -ForegroundColor Yellow

    # Create the import request
    $importParams = @{
        Mailbox = $TargetUserEmail
        FilePath = $PSTPath
        Name = "Import-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    }

    if ($ImportToArchive) {
        $importParams.Add("IsArchive", $true)
    }

    $importRequest = New-MailboxImportRequest @importParams

    Write-Host "`nImport request created successfully!" -ForegroundColor Green
    Write-Host "Request Name: $($importRequest.Name)" -ForegroundColor Cyan
    Write-Host "Status: $($importRequest.Status)" -ForegroundColor Cyan

    Write-Host "`nMonitoring import progress..." -ForegroundColor Yellow
    Write-Host "This may take several minutes to hours depending on PST size." -ForegroundColor Yellow
    Write-Host "Press Ctrl+C to stop monitoring (import will continue in background)`n" -ForegroundColor Yellow

    # Monitor progress
    do {
        Start-Sleep -Seconds 30
        $status = Get-MailboxImportRequest -Identity $importRequest.Identity

        $percentComplete = if ($status.PercentComplete) { $status.PercentComplete } else { 0 }
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Status: $($status.Status) | Progress: $percentComplete% | Items Transferred: $($status.ItemsTransferred)" -ForegroundColor Cyan

        if ($status.Status -eq "Failed") {
            Write-Host "`nImport FAILED!" -ForegroundColor Red
            Write-Host "Error: $($status.Message)" -ForegroundColor Red
            break
        }

    } while ($status.Status -ne "Completed")

    if ($status.Status -eq "Completed") {
        Write-Host "`nImport completed successfully!" -ForegroundColor Green
        Write-Host "Items Transferred: $($status.ItemsTransferred)" -ForegroundColor Cyan
        Write-Host "Bytes Transferred: $([math]::Round($status.BytesTransferred / 1MB, 2)) MB" -ForegroundColor Cyan

        # Clean up import request
        $cleanup = Read-Host "`nDo you want to remove the import request from Exchange Online? (Y/N)"
        if ($cleanup -eq 'Y') {
            Remove-MailboxImportRequest -Identity $importRequest.Identity -Confirm:$false
            Write-Host "Import request removed." -ForegroundColor Green
        }
    }

    Disconnect-ExchangeOnline -Confirm:$false
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    exit 1
}
