<#
.SYNOPSIS
    Simple OneDrive downloader using PnP with better GUID account handling

.DESCRIPTION
    Simplified approach that connects directly without complex permission checks

.PARAMETER OneDriveUrl
    The full OneDrive URL (e.g., https://tenant-my.sharepoint.com/personal/user_path)

.PARAMETER DestinationPath
    Local path where files will be downloaded

.EXAMPLE
    .\Download-OneDrive-Simple.ps1 -OneDriveUrl "https://theivycollection-my.sharepoint.com/personal/1b26a836-e5ef-4408-9407-0617613467a9_theivycollection_onmicrosoft_com" -DestinationPath "C:\temp"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$OneDriveUrl,

    [Parameter(Mandatory=$true)]
    [string]$DestinationPath
)

# Set up logging
$LogFile = Join-Path $DestinationPath "OneDrive-Simple-Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
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

Write-Log "=== Simple OneDrive Download Started ===" -Level INFO
Write-Log "OneDrive URL: $OneDriveUrl" -Level INFO
Write-Log "Destination Path: $DestinationPath" -Level INFO
Write-Log "Log File: $LogFile" -Level INFO

# Check if PnP.PowerShell module is installed
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Log "Installing PnP.PowerShell module..." -Level WARNING
    Install-Module -Name PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
}

Import-Module PnP.PowerShell

try {
    # Connect with Interactive authentication (most reliable)
    Write-Log "Connecting to OneDrive..." -Level INFO
    Write-Log "You will be prompted to authenticate..." -Level WARNING

    Connect-PnPOnline -Url $OneDriveUrl -Interactive
    Write-Log "Successfully connected!" -Level SUCCESS

    # Get Web title to confirm connection
    Write-Log "Verifying connection..." -Level INFO
    $ctx = Get-PnPContext
    $web = $ctx.Web
    $ctx.Load($web)
    try {
        $ctx.ExecuteQuery()
        Write-Log "Connected to: $($web.Title)" -Level SUCCESS
    }
    catch {
        Write-Log "Note: Could not load web title, but continuing..." -Level WARNING
    }

    # List all available lists/libraries
    Write-Log "Discovering available libraries..." -Level INFO
    $allLists = Get-PnPList
    Write-Log "Found $($allLists.Count) lists/libraries:" -Level INFO

    $documentsLib = $null
    foreach ($list in $allLists) {
        $listInfo = "  - $($list.Title) (Template: $($list.BaseTemplate))"
        if ($list.BaseTemplate -eq 101) {
            $listInfo += " [DOCUMENT LIBRARY]"
            if ($list.Title -eq "Documents") {
                $documentsLib = $list
                $listInfo += " <-- TARGET"
            }
        }
        Write-Log $listInfo -Level INFO
    }

    if (-not $documentsLib) {
        Write-Log "WARNING: 'Documents' library not found, using first document library..." -Level WARNING
        $documentsLib = $allLists | Where-Object { $_.BaseTemplate -eq 101 } | Select-Object -First 1
    }

    if (-not $documentsLib) {
        throw "No document library found in this OneDrive site"
    }

    Write-Log "Using library: $($documentsLib.Title)" -Level SUCCESS
    Write-Log "Library URL: $($documentsLib.RootFolder.ServerRelativeUrl)" -Level INFO

    # Get all items
    Write-Log "Retrieving all items from library..." -Level INFO
    $items = Get-PnPListItem -List $documentsLib -PageSize 5000
    Write-Log "Found $($items.Count) total items" -Level SUCCESS

    # Download files
    $filesProcessed = 0
    $foldersProcessed = 0
    $successCount = 0
    $errorCount = 0

    Write-Log "Starting download process..." -Level INFO

    foreach ($item in $items) {
        if ($item.FileSystemObjectType -eq "File") {
            $filesProcessed++

            $serverRelUrl = $item.FieldValues.FileRef
            $fileName = $item.FieldValues.FileLeafRef
            $fileSize = $item.FieldValues.File_x0020_Size

            # Calculate relative path
            $libraryPath = $documentsLib.RootFolder.ServerRelativeUrl
            $relativePath = $serverRelUrl -replace [regex]::Escape($libraryPath), ""
            $relativePath = $relativePath.TrimStart('/')

            $localPath = Join-Path $DestinationPath $relativePath
            $localDir = Split-Path $localPath -Parent

            # Create directory if needed
            if (-not (Test-Path $localDir)) {
                New-Item -ItemType Directory -Path $localDir -Force | Out-Null
            }

            Write-Progress -Activity "Downloading files" -Status "[$filesProcessed] $fileName" -PercentComplete (($filesProcessed / $items.Count) * 100)

            try {
                Write-Log "[$filesProcessed/$($items.Count)] Downloading: $relativePath ($('{0:N2}' -f ($fileSize/1KB)) KB)" -Level INFO

                # Download file
                Get-PnPFile -Url $serverRelUrl -Path (Split-Path $localPath -Parent) -FileName (Split-Path $localPath -Leaf) -AsFile -Force

                Write-Log "  SUCCESS: $localPath" -Level SUCCESS
                $successCount++
            }
            catch {
                Write-Log "  ERROR: $($_.Exception.Message)" -Level ERROR
                $errorCount++
            }
        }
        else {
            $foldersProcessed++
        }
    }

    Write-Progress -Activity "Downloading files" -Completed

    Write-Log "`n=== Download Complete ===" -Level SUCCESS
    Write-Log "Total items: $($items.Count)" -Level INFO
    Write-Log "Files processed: $filesProcessed" -Level INFO
    Write-Log "Files downloaded: $successCount" -Level SUCCESS
    Write-Log "Files failed: $errorCount" -Level $(if($errorCount -gt 0){'ERROR'}else{'INFO'})
    Write-Log "Folders: $foldersProcessed" -Level INFO
    Write-Log "Destination: $DestinationPath" -Level SUCCESS
    Write-Log "Log file: $LogFile" -Level INFO

    Disconnect-PnPOnline
    Write-Log "=== Script completed successfully ===" -Level SUCCESS
}
catch {
    Write-Log "=== FATAL ERROR ===" -Level ERROR
    Write-Log "Error: $($_.Exception.Message)" -Level ERROR
    Write-Log "Error Type: $($_.Exception.GetType().FullName)" -Level ERROR
    if ($_.ScriptStackTrace) {
        Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level ERROR
    }
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
    exit 1
}
