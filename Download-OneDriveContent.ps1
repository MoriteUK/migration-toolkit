<#
.SYNOPSIS
    Grant admin access to a user's OneDrive and download all content to local drive

.DESCRIPTION
    This script connects to SharePoint Online, grants site collection admin access
    to a user's OneDrive, and downloads all files to a local directory.

.PARAMETER UserEmail
    The email address of the user whose OneDrive you want to access

.PARAMETER AdminEmail
    The admin email address to grant access to (typically your admin account)

.PARAMETER TenantName
    Your SharePoint tenant name ONLY (e.g., 'theivycollection' NOT the full URL)
    This is the part before '-admin.sharepoint.com'

.PARAMETER DestinationPath
    Local path where files will be downloaded

.EXAMPLE
    .\Download-OneDriveContent.ps1 -UserEmail "user@domain.com" -AdminEmail "admin@domain.com" -TenantName "theivycollection" -DestinationPath "C:\ODBackup"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$UserEmail,

    [Parameter(Mandatory=$true)]
    [string]$AdminEmail,

    [Parameter(Mandatory=$true)]
    [string]$TenantName,

    [Parameter(Mandatory=$true)]
    [string]$DestinationPath
)

# Set up logging
$LogFile = Join-Path $DestinationPath "OneDrive-Download-Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARNING','ERROR','SUCCESS')]
        [string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $logMessage

    switch ($Level) {
        'INFO'    { Write-Host $Message -ForegroundColor Cyan }
        'WARNING' { Write-Host $Message -ForegroundColor Yellow }
        'ERROR'   { Write-Host $Message -ForegroundColor Red }
        'SUCCESS' { Write-Host $Message -ForegroundColor Green }
    }
}

# Create destination folder if it doesn't exist (for log file)
if (-not (Test-Path $DestinationPath)) {
    New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
}

Write-Log "=== OneDrive Download Script Started ===" -Level INFO
Write-Log "User Email: $UserEmail" -Level INFO
Write-Log "Admin Email: $AdminEmail" -Level INFO
Write-Log "Tenant Name (input): $TenantName" -Level INFO
Write-Log "Destination Path: $DestinationPath" -Level INFO
Write-Log "Log File: $LogFile" -Level INFO

# Clean up TenantName if user provided full URL
if ($TenantName -match 'https?://(.+?)(?:-admin)?\.sharepoint\.com') {
    $TenantName = $Matches[1]
    Write-Log "Extracted tenant name: $TenantName" -Level WARNING
}
$TenantName = $TenantName.Trim()
Write-Log "Tenant Name (cleaned): $TenantName" -Level INFO

# Check if PnP.PowerShell module is installed
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Log "Installing PnP.PowerShell module..." -Level WARNING
    Install-Module -Name PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
}

Write-Log "Importing PnP.PowerShell module..." -Level INFO
Import-Module PnP.PowerShell
$moduleVersion = (Get-Module PnP.PowerShell).Version
Write-Log "PnP.PowerShell version: $moduleVersion" -Level INFO

# Convert email to OneDrive URL format (replace @ and . with _, keep - in username)
$userFormatted = $UserEmail -replace '@', '_' -replace '\.', '_'
$oneDriveUrl = "https://$TenantName-my.sharepoint.com/personal/$userFormatted"

Write-Log "Formatted user path: $userFormatted" -Level INFO
Write-Log "OneDrive URL: $oneDriveUrl" -Level INFO

$adminUrl = "https://$TenantName-admin.sharepoint.com"
Write-Log "Admin URL: $adminUrl" -Level INFO
Write-Log "Connecting to SharePoint Online admin center..." -Level INFO

try {
    # Connect to SharePoint Online Admin
    # Using DeviceLogin to avoid app registration issues
    Write-Log "Attempting connection to admin center..." -Level INFO
    Connect-PnPOnline -Url $adminUrl -UseWebLogin
    Write-Log "Successfully connected to admin center" -Level SUCCESS

    # Check current connection
    $web = Get-PnPWeb
    Write-Log "Connected to web: $($web.Title) - $($web.Url)" -Level INFO

    Write-Log "Granting site collection admin access to $AdminEmail on OneDrive: $oneDriveUrl" -Level INFO
    try {
        Set-PnPTenantSite -Identity $oneDriveUrl -Owners $AdminEmail
        Write-Log "Successfully granted admin access to $AdminEmail" -Level SUCCESS
    }
    catch {
        Write-Log "Failed to grant admin access: $($_.Exception.Message)" -Level ERROR
        Write-Log "This might still work if you already have access..." -Level WARNING
    }

    Write-Log "Waiting 30 seconds for permissions to propagate..." -Level INFO
    Start-Sleep -Seconds 30

    # Disconnect and reconnect to the OneDrive site
    Write-Log "Disconnecting from admin center..." -Level INFO
    Disconnect-PnPOnline

    Write-Log "Connecting to user's OneDrive: $oneDriveUrl" -Level INFO
    Write-Log "Using Interactive authentication method..." -Level INFO
    Connect-PnPOnline -Url $oneDriveUrl -Interactive
    Write-Log "Successfully connected to OneDrive site" -Level SUCCESS

    # Verify connection and permissions (with error handling)
    Write-Log "Attempting to retrieve web properties..." -Level INFO
    try {
        $web = Get-PnPWeb -Includes Title,Url
        Write-Log "Connected to OneDrive web: $($web.Title) - $($web.Url)" -Level SUCCESS
    }
    catch {
        Write-Log "Warning: Could not retrieve web properties: $($_.Exception.Message)" -Level WARNING
        Write-Log "Continuing anyway - this may be a permission timing issue..." -Level WARNING
    }

    # Check current user (with error handling)
    Write-Log "Attempting to retrieve current user..." -Level INFO
    try {
        $context = Get-PnPContext
        $currentUser = $context.Web.CurrentUser
        $context.Load($currentUser)
        $context.ExecuteQuery()
        Write-Log "Current user: $($currentUser.Title) ($($currentUser.LoginName))" -Level SUCCESS
    }
    catch {
        Write-Log "Warning: Could not retrieve current user: $($_.Exception.Message)" -Level WARNING
    }

    # Test site permissions with retry logic
    Write-Log "Testing site collection administrators..." -Level INFO
    $retryCount = 0
    $maxRetries = 3
    $permissionSuccess = $false

    while (-not $permissionSuccess -and $retryCount -lt $maxRetries) {
        try {
            $retryCount++
            Write-Log "Permission check attempt $retryCount of $maxRetries..." -Level INFO

            $site = Get-PnPSite
            $owners = Get-PnPSiteCollectionAdmin

            if ($owners) {
                Write-Log "Site Collection Admins:" -Level SUCCESS
                foreach ($owner in $owners) {
                    Write-Log "  - $($owner.Title) ($($owner.Email))" -Level INFO
                }
                $permissionSuccess = $true
            }
        }
        catch {
            Write-Log "Could not retrieve site admins (attempt $retryCount): $($_.Exception.Message)" -Level WARNING
            if ($retryCount -lt $maxRetries) {
                Write-Log "Waiting 10 seconds before retry..." -Level INFO
                Start-Sleep -Seconds 10
            }
        }
    }

    if (-not $permissionSuccess) {
        Write-Log "Could not verify admin permissions after $maxRetries attempts" -Level WARNING
        Write-Log "This may indicate permissions haven't fully propagated yet" -Level WARNING
    }

    Write-Log "Starting download of all files from OneDrive..." -Level INFO
    Write-Log "Destination: $DestinationPath" -Level SUCCESS

    # Test access to Documents library with retry
    Write-Log "Testing access to Documents library..." -Level INFO
    $retryCount = 0
    $maxRetries = 3
    $librarySuccess = $false
    $list = $null

    while (-not $librarySuccess -and $retryCount -lt $maxRetries) {
        try {
            $retryCount++
            Write-Log "Library access attempt $retryCount of $maxRetries..." -Level INFO

            # Try different methods to access the library
            try {
                $list = Get-PnPList -Identity "Documents"
                Write-Log "Documents library found: $($list.Title) - Items: $($list.ItemCount)" -Level SUCCESS
                $librarySuccess = $true
            }
            catch {
                Write-Log "Standard method failed, trying alternative..." -Level WARNING
                # Alternative: Try accessing via URL
                $list = Get-PnPList -Identity "/personal/$userFormatted/Documents"
                Write-Log "Documents library found via URL path" -Level SUCCESS
                $librarySuccess = $true
            }
        }
        catch {
            Write-Log "Library access failed (attempt $retryCount): $($_.Exception.Message)" -Level ERROR
            if ($retryCount -lt $maxRetries) {
                Write-Log "Waiting 15 seconds before retry..." -Level INFO
                Start-Sleep -Seconds 15
            }
        }
    }

    if (-not $librarySuccess) {
        Write-Log "CRITICAL: Could not access Documents library after $maxRetries attempts" -Level ERROR
        Write-Log "This usually means permissions haven't propagated or the OneDrive site is restricted" -Level ERROR

        # Diagnostic: Try to list all available lists
        Write-Log "Attempting to list all available libraries for diagnostics..." -Level INFO
        try {
            $allLists = Get-PnPList
            Write-Log "Available lists/libraries:" -Level INFO
            foreach ($availList in $allLists) {
                Write-Log "  - $($availList.Title) (Type: $($availList.BaseTemplate)) - URL: $($availList.RootFolder.ServerRelativeUrl)" -Level INFO
            }
        }
        catch {
            Write-Log "Could not retrieve list of libraries: $($_.Exception.Message)" -Level ERROR
        }

        throw "Unable to access Documents library - permissions may not have propagated"
    }

    # Download all files recursively from the Documents library
    Write-Log "Retrieving file list from Documents library..." -Level INFO
    $files = Get-PnPListItem -List "Documents" -PageSize 1000
    Write-Log "Retrieved $($files.Count) items from Documents library" -Level INFO
    $totalFiles = $files.Count
    $current = 0
    $successCount = 0
    $errorCount = 0
    $folderCount = 0

    Write-Log "Processing $totalFiles items..." -Level INFO

    foreach ($file in $files) {
        $current++

        if ($file.FileSystemObjectType -eq "File") {
            $serverRelativeUrl = $file.FieldValues.FileRef
            $fileName = $file.FieldValues.FileLeafRef
            $filePath = $serverRelativeUrl -replace "^/personal/$userFormatted/Documents/", ""

            $localPath = Join-Path $DestinationPath $filePath
            $localDir = Split-Path $localPath -Parent

            if (-not (Test-Path $localDir)) {
                New-Item -ItemType Directory -Path $localDir -Force | Out-Null
            }

            Write-Progress -Activity "Downloading files" -Status "[$current/$totalFiles] $fileName" -PercentComplete (($current / $totalFiles) * 100)

            try {
                Write-Log "Downloading [$current/$totalFiles]: $filePath" -Level INFO
                Get-PnPFile -Url $serverRelativeUrl -Path (Split-Path $localPath -Parent) -FileName (Split-Path $localPath -Leaf) -AsFile -Force
                Write-Log "  SUCCESS: Downloaded to $localPath" -Level SUCCESS
                $successCount++
            }
            catch {
                Write-Log "  ERROR: Failed to download $filePath" -Level ERROR
                Write-Log "  Error details: $($_.Exception.Message)" -Level ERROR
                Write-Log "  Server URL: $serverRelativeUrl" -Level ERROR
                $errorCount++
            }
        }
        else {
            Write-Log "Folder [$current/$totalFiles]: $($file.FieldValues.FileRef)" -Level INFO
            $folderCount++
        }
    }

    Write-Log "`n=== Download Complete ===" -Level SUCCESS
    Write-Log "Total items processed: $totalFiles" -Level INFO
    Write-Log "Files downloaded: $successCount" -Level SUCCESS
    Write-Log "Files failed: $errorCount" -Level $(if($errorCount -gt 0){'ERROR'}else{'INFO'})
    Write-Log "Folders: $folderCount" -Level INFO
    Write-Log "Files saved to: $DestinationPath" -Level SUCCESS
    Write-Log "Log file saved to: $LogFile" -Level INFO

    # Optional: Remove admin access after completion
    $removeAccess = Read-Host "`nDo you want to remove admin access from $AdminEmail? (Y/N)"
    if ($removeAccess -eq 'Y') {
        Write-Log "Removing admin access from $AdminEmail..." -Level INFO
        Disconnect-PnPOnline
        Connect-PnPOnline -Url $adminUrl -UseWebLogin
        Set-PnPTenantSite -Identity $oneDriveUrl -Owners "" -RemoveOwners $AdminEmail
        Write-Log "Admin access removed." -Level SUCCESS
    }

    Disconnect-PnPOnline
    Write-Log "=== Script completed successfully ===" -Level SUCCESS
}
catch {
    Write-Log "=== FATAL ERROR ===" -Level ERROR
    Write-Log "Error: $($_.Exception.Message)" -Level ERROR
    Write-Log "Error Type: $($_.Exception.GetType().FullName)" -Level ERROR
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level ERROR
    Write-Log "Script failed. Check log file: $LogFile" -Level ERROR
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
    exit 1
}
