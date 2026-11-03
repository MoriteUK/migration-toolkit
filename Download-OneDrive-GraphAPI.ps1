<#
.SYNOPSIS
    Download OneDrive content using Microsoft Graph API (alternative method)

.DESCRIPTION
    This script uses Microsoft Graph API instead of PnP PowerShell to access OneDrive.
    It may work better with GUID-based user accounts.

.PARAMETER UserEmail
    The email address of the user whose OneDrive you want to access

.PARAMETER DestinationPath
    Local path where files will be downloaded

.EXAMPLE
    .\Download-OneDrive-GraphAPI.ps1 -UserEmail "user@domain.com" -DestinationPath "C:\ODBackup"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$UserEmail,

    [Parameter(Mandatory=$true)]
    [string]$DestinationPath
)

# Set up logging
$LogFile = Join-Path $DestinationPath "OneDrive-GraphAPI-Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
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

# Create destination folder if it doesn't exist
if (-not (Test-Path $DestinationPath)) {
    New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
}

Write-Log "=== OneDrive Download Script (Graph API) Started ===" -Level INFO
Write-Log "User Email: $UserEmail" -Level INFO
Write-Log "Destination Path: $DestinationPath" -Level INFO
Write-Log "Log File: $LogFile" -Level INFO

# Check if Microsoft.Graph modules are installed
$requiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Files', 'Microsoft.Graph.Users')
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Log "Installing $module module..." -Level WARNING
        Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser
    }
}

Write-Log "Importing Microsoft.Graph modules..." -Level INFO
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Files
Import-Module Microsoft.Graph.Users

try {
    # Connect to Microsoft Graph with required scopes
    Write-Log "Connecting to Microsoft Graph..." -Level INFO
    Write-Log "You will be prompted to sign in with admin credentials..." -Level WARNING

    $scopes = @(
        "Files.Read.All",
        "Sites.Read.All",
        "User.Read.All"
    )

    Connect-MgGraph -Scopes $scopes
    Write-Log "Successfully connected to Microsoft Graph" -Level SUCCESS

    # Get the current context
    $context = Get-MgContext
    Write-Log "Connected as: $($context.Account)" -Level INFO
    Write-Log "Tenant ID: $($context.TenantId)" -Level INFO

    # Get user information - try multiple methods
    Write-Log "Looking up user: $UserEmail" -Level INFO
    $user = $null

    # Method 1: Try direct lookup
    try {
        $user = Get-MgUser -UserId $UserEmail -Property Id,DisplayName,Mail,UserPrincipalName -ErrorAction Stop
        Write-Log "Found user via direct lookup: $($user.DisplayName) ($($user.UserPrincipalName))" -Level SUCCESS
    }
    catch {
        Write-Log "Direct lookup failed: $($_.Exception.Message)" -Level WARNING
    }

    # Method 2: If email looks like a GUID, try using the GUID directly
    if (-not $user -and $UserEmail -match '^([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})@') {
        $guid = $matches[1]
        Write-Log "Email appears to contain a GUID: $guid" -Level INFO
        Write-Log "Attempting lookup by GUID..." -Level INFO
        try {
            $user = Get-MgUser -UserId $guid -Property Id,DisplayName,Mail,UserPrincipalName -ErrorAction Stop
            Write-Log "Found user via GUID lookup: $($user.DisplayName) ($($user.UserPrincipalName))" -Level SUCCESS
        }
        catch {
            Write-Log "GUID lookup failed: $($_.Exception.Message)" -Level WARNING
        }
    }

    # Method 3: Search for the user
    if (-not $user) {
        Write-Log "Attempting to search for user..." -Level INFO
        try {
            $users = Get-MgUser -Filter "userPrincipalName eq '$UserEmail' or mail eq '$UserEmail'" -Property Id,DisplayName,Mail,UserPrincipalName
            if ($users) {
                $user = $users[0]
                Write-Log "Found user via search: $($user.DisplayName) ($($user.UserPrincipalName))" -Level SUCCESS
            }
        }
        catch {
            Write-Log "Search failed: $($_.Exception.Message)" -Level WARNING
        }
    }

    # Method 4: List all users and try to find a match
    if (-not $user) {
        Write-Log "Attempting to list users to find a match..." -Level WARNING
        Write-Log "This may take a moment for large organizations..." -Level INFO
        try {
            # Get all users and search for the email
            $allUsers = Get-MgUser -All -Property Id,DisplayName,Mail,UserPrincipalName | Where-Object {
                $_.UserPrincipalName -eq $UserEmail -or $_.Mail -eq $UserEmail
            }
            if ($allUsers) {
                $user = $allUsers[0]
                Write-Log "Found user in user list: $($user.DisplayName) ($($user.UserPrincipalName))" -Level SUCCESS
            }
        }
        catch {
            Write-Log "User listing failed: $($_.Exception.Message)" -Level WARNING
        }
    }

    if (-not $user) {
        Write-Log "CRITICAL: Could not find user with email: $UserEmail" -Level ERROR
        Write-Log "Attempting to list first 10 users for diagnostic purposes..." -Level INFO
        try {
            $sampleUsers = Get-MgUser -Top 10 -Property Id,DisplayName,Mail,UserPrincipalName
            Write-Log "Sample of users in your tenant:" -Level INFO
            foreach ($sampleUser in $sampleUsers) {
                Write-Log "  - $($sampleUser.DisplayName) | UPN: $($sampleUser.UserPrincipalName) | Mail: $($sampleUser.Mail) | ID: $($sampleUser.Id)" -Level INFO
            }
        }
        catch {
            Write-Log "Could not retrieve sample users: $($_.Exception.Message)" -Level ERROR
        }
        throw "User not found: $UserEmail"
    }

    Write-Log "User ID: $($user.Id)" -Level INFO
    Write-Log "User Display Name: $($user.DisplayName)" -Level INFO
    Write-Log "User Principal Name: $($user.UserPrincipalName)" -Level INFO
    Write-Log "User Mail: $($user.Mail)" -Level INFO

    # Get OneDrive information
    Write-Log "Accessing user's OneDrive..." -Level INFO
    $drive = Get-MgUserDrive -UserId $user.Id
    Write-Log "OneDrive found: $($drive.Name)" -Level SUCCESS
    Write-Log "Drive ID: $($drive.Id)" -Level INFO

    # Function to download files recursively
    function Download-DriveItems {
        param(
            [string]$UserId,
            [string]$DriveId,
            [string]$ItemId = "root",
            [string]$LocalPath,
            [ref]$Stats
        )

        Write-Log "Scanning folder: $ItemId" -Level INFO

        # Get items in current folder
        if ($ItemId -eq "root") {
            $items = Get-MgUserDriveRootChild -UserId $UserId -DriveId $DriveId -All
        }
        else {
            $items = Get-MgUserDriveItemChild -UserId $UserId -DriveId $DriveId -DriveItemId $ItemId -All
        }

        Write-Log "Found $($items.Count) items" -Level INFO

        foreach ($item in $items) {
            $itemPath = Join-Path $LocalPath $item.Name

            if ($item.Folder) {
                # It's a folder - create it and recurse
                Write-Log "Creating folder: $($item.Name)" -Level INFO
                if (-not (Test-Path $itemPath)) {
                    New-Item -ItemType Directory -Path $itemPath -Force | Out-Null
                }
                $Stats.Value.FolderCount++

                # Recurse into folder
                Download-DriveItems -UserId $UserId -DriveId $DriveId -ItemId $item.Id -LocalPath $itemPath -Stats $Stats
            }
            else {
                # It's a file - download it
                try {
                    $Stats.Value.TotalFiles++
                    Write-Log "Downloading file [$($Stats.Value.TotalFiles)]: $($item.Name) ($([math]::Round($item.Size / 1MB, 2)) MB)" -Level INFO

                    # Download the file
                    $fileContent = Get-MgUserDriveItemContent -UserId $UserId -DriveId $DriveId -DriveItemId $item.Id
                    [System.IO.File]::WriteAllBytes($itemPath, $fileContent)

                    Write-Log "  SUCCESS: Downloaded to $itemPath" -Level SUCCESS
                    $Stats.Value.SuccessCount++
                }
                catch {
                    Write-Log "  ERROR: Failed to download $($item.Name)" -Level ERROR
                    Write-Log "  Error details: $($_.Exception.Message)" -Level ERROR
                    $Stats.Value.ErrorCount++
                }
            }
        }
    }

    # Initialize statistics
    $stats = @{
        TotalFiles = 0
        SuccessCount = 0
        ErrorCount = 0
        FolderCount = 0
    }

    Write-Log "Starting recursive download..." -Level INFO
    Download-DriveItems -UserId $user.Id -DriveId $drive.Id -ItemId "root" -LocalPath $DestinationPath -Stats ([ref]$stats)

    Write-Log "`n=== Download Complete ===" -Level SUCCESS
    Write-Log "Total files: $($stats.TotalFiles)" -Level INFO
    Write-Log "Files downloaded: $($stats.SuccessCount)" -Level SUCCESS
    Write-Log "Files failed: $($stats.ErrorCount)" -Level $(if($stats.ErrorCount -gt 0){'ERROR'}else{'INFO'})
    Write-Log "Folders created: $($stats.FolderCount)" -Level INFO
    Write-Log "Files saved to: $DestinationPath" -Level SUCCESS
    Write-Log "Log file saved to: $LogFile" -Level INFO

    # Disconnect
    Disconnect-MgGraph
    Write-Log "=== Script completed successfully ===" -Level SUCCESS
}
catch {
    Write-Log "=== FATAL ERROR ===" -Level ERROR
    Write-Log "Error: $($_.Exception.Message)" -Level ERROR
    Write-Log "Error Type: $($_.Exception.GetType().FullName)" -Level ERROR
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level ERROR
    Write-Log "Script failed. Check log file: $LogFile" -Level ERROR

    Disconnect-MgGraph -ErrorAction SilentlyContinue
    exit 1
}
