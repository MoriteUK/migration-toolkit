<#
.SYNOPSIS
    Upload local files to a target OneDrive after migration

.DESCRIPTION
    This script uploads files from a local directory to a target user's OneDrive
    in a destination tenant, preserving folder structure.

.PARAMETER TargetUserEmail
    The email address of the target user in the destination tenant

.PARAMETER AdminEmail
    The admin email address in the destination tenant

.PARAMETER TargetTenantName
    The destination SharePoint tenant name ONLY (e.g., 'theivycollection' NOT the full URL)
    This is the part before '-admin.sharepoint.com'

.PARAMETER SourcePath
    Local path containing the files to upload

.PARAMETER OverwriteExisting
    If specified, will overwrite existing files in the destination

.EXAMPLE
    .\Upload-OneDriveContent.ps1 -TargetUserEmail "user@newdomain.com" -AdminEmail "admin@newdomain.com" -TargetTenantName "theivycollection" -SourcePath "C:\ODBackup"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$TargetUserEmail,

    [Parameter(Mandatory=$true)]
    [string]$AdminEmail,

    [Parameter(Mandatory=$true)]
    [string]$TargetTenantName,

    [Parameter(Mandatory=$true)]
    [string]$SourcePath,

    [switch]$OverwriteExisting
)

# Clean up TenantName if user provided full URL
if ($TargetTenantName -match 'https?://(.+?)(?:-admin)?\.sharepoint\.com') {
    $TargetTenantName = $Matches[1]
    Write-Host "Extracted tenant name: $TargetTenantName" -ForegroundColor Yellow
}
$TargetTenantName = $TargetTenantName.Trim()

# Check if PnP.PowerShell module is installed
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "Installing PnP.PowerShell module..." -ForegroundColor Yellow
    Install-Module -Name PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
}

Import-Module PnP.PowerShell

# Verify source path exists
if (-not (Test-Path $SourcePath)) {
    Write-Host "Error: Source path does not exist: $SourcePath" -ForegroundColor Red
    exit 1
}

# Convert email to OneDrive URL format (replace @ and . with _, keep - in username)
$userFormatted = $TargetUserEmail -replace '@', '_' -replace '\.', '_'
$oneDriveUrl = "https://$TargetTenantName-my.sharepoint.com/personal/$userFormatted"

Write-Host "Connecting to target SharePoint Online admin center..." -ForegroundColor Cyan
$adminUrl = "https://$TargetTenantName-admin.sharepoint.com"

try {
    # Connect to SharePoint Online Admin
    # Using DeviceLogin to avoid app registration issues
    Connect-PnPOnline -Url $adminUrl -UseWebLogin

    Write-Host "Granting site collection admin access to $AdminEmail on target OneDrive: $oneDriveUrl" -ForegroundColor Cyan
    Set-PnPTenantSite -Identity $oneDriveUrl -Owners $AdminEmail

    Write-Host "Waiting 10 seconds for permissions to propagate..." -ForegroundColor Yellow
    Start-Sleep -Seconds 10

    # Disconnect and reconnect to the OneDrive site
    Disconnect-PnPOnline

    Write-Host "Connecting to target user's OneDrive: $oneDriveUrl" -ForegroundColor Cyan
    Connect-PnPOnline -Url $oneDriveUrl -UseWebLogin

    # Get all files from source directory
    $allFiles = Get-ChildItem -Path $SourcePath -Recurse -File
    $totalFiles = $allFiles.Count
    $current = 0
    $successCount = 0
    $failCount = 0

    Write-Host "Starting upload of $totalFiles files to target OneDrive..." -ForegroundColor Cyan

    foreach ($file in $allFiles) {
        $current++

        # Calculate relative path and SharePoint path
        $relativePath = $file.FullName.Substring($SourcePath.Length).TrimStart('\', '/')
        $relativePath = $relativePath -replace '\\', '/'
        $targetFolder = Split-Path $relativePath -Parent

        if ([string]::IsNullOrEmpty($targetFolder)) {
            $targetFolder = "/"
        } else {
            $targetFolder = "/$targetFolder"
        }

        Write-Progress -Activity "Uploading files" -Status "[$current/$totalFiles] $relativePath" -PercentComplete (($current / $totalFiles) * 100)

        try {
            # Ensure target folder exists
            if ($targetFolder -ne "/") {
                $folderParts = $targetFolder.Split('/', [StringSplitOptions]::RemoveEmptyEntries)
                $currentPath = ""

                foreach ($part in $folderParts) {
                    $currentPath += "/$part"
                    try {
                        $null = Get-PnPFolder -Url "Documents$currentPath" -ErrorAction Stop
                    }
                    catch {
                        $parentPath = $currentPath -replace "/$part$", ""
                        $folderBase = if ([string]::IsNullOrEmpty($parentPath)) { "Documents" } else { "Documents$parentPath" }
                        Add-PnPFolder -Name $part -Folder $folderBase | Out-Null
                    }
                }
            }

            # Upload file
            $uploadParams = @{
                Path = $file.FullName
                Folder = "Documents$targetFolder"
            }

            if ($OverwriteExisting) {
                $uploadParams.Add("Force", $true)
            }

            Add-PnPFile @uploadParams | Out-Null
            Write-Host "Uploaded: $relativePath" -ForegroundColor Green
            $successCount++
        }
        catch {
            Write-Host "Failed to upload: $relativePath - $($_.Exception.Message)" -ForegroundColor Red
            $failCount++
        }
    }

    Write-Host "`nUpload complete!" -ForegroundColor Green
    Write-Host "Total files: $totalFiles" -ForegroundColor Cyan
    Write-Host "Successful: $successCount" -ForegroundColor Green
    Write-Host "Failed: $failCount" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Green" })

    # Optional: Remove admin access after completion
    $removeAccess = Read-Host "`nDo you want to remove admin access from $AdminEmail? (Y/N)"
    if ($removeAccess -eq 'Y') {
        Disconnect-PnPOnline
        Connect-PnPOnline -Url $adminUrl -UseWebLogin
        Set-PnPTenantSite -Identity $oneDriveUrl -Owners "" -RemoveOwners $AdminEmail
        Write-Host "Admin access removed." -ForegroundColor Green
    }

    Disconnect-PnPOnline
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
    exit 1
}
