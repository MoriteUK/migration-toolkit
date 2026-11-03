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
    Your SharePoint tenant name (e.g., 'contoso' for contoso.sharepoint.com)

.PARAMETER DestinationPath
    Local path where files will be downloaded

.EXAMPLE
    .\Download-OneDriveContent.ps1 -UserEmail "user@domain.com" -AdminEmail "admin@domain.com" -TenantName "contoso" -DestinationPath "C:\ODBackup"
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

# Check if SharePoint Online Management Shell module is installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Online.SharePoint.PowerShell)) {
    Write-Host "Installing Microsoft.Online.SharePoint.PowerShell module..." -ForegroundColor Yellow
    Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force -AllowClobber -Scope CurrentUser
}

# Check if SharePoint CSOM is available
$packagesPath = Join-Path $PSScriptRoot "packages"
$csomPath = $null

# Try to find existing installation
$nugetBase = "$env:USERPROFILE\.nuget\packages\microsoft.sharepoint.client"
if (Test-Path $nugetBase) {
    $versions = Get-ChildItem $nugetBase -Directory | Sort-Object Name -Descending
    if ($versions.Count -gt 0) {
        $csomPath = Join-Path $versions[0].FullName "lib\net45"
    }
}

# If not found, try local packages folder
if (-not $csomPath -or -not (Test-Path "$csomPath\Microsoft.SharePoint.Client.dll")) {
    $localPackage = Get-ChildItem -Path $packagesPath -Filter "Microsoft.SharePoint.Client.*" -Directory -ErrorAction SilentlyContinue |
                    Sort-Object Name -Descending | Select-Object -First 1

    if ($localPackage) {
        $csomPath = Join-Path $localPackage.FullName "lib\net45"
    }
}

# If still not found, download it
if (-not $csomPath -or -not (Test-Path "$csomPath\Microsoft.SharePoint.Client.dll")) {
    Write-Host "SharePoint CSOM libraries not found. Downloading..." -ForegroundColor Yellow

    if (-not (Test-Path $packagesPath)) {
        New-Item -ItemType Directory -Path $packagesPath -Force | Out-Null
    }

    # Download using NuGet.exe
    $nugetUrl = "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe"
    $nugetPath = Join-Path $packagesPath "nuget.exe"

    if (-not (Test-Path $nugetPath)) {
        Write-Host "Downloading NuGet.exe..." -ForegroundColor Yellow
        Invoke-WebRequest -Uri $nugetUrl -OutFile $nugetPath
    }

    Write-Host "Installing SharePoint CSOM package..." -ForegroundColor Yellow
    & $nugetPath install Microsoft.SharePoint.Client -OutputDirectory $packagesPath -NonInteractive

    # Find the downloaded package
    $localPackage = Get-ChildItem -Path $packagesPath -Filter "Microsoft.SharePoint.Client.*" -Directory |
                    Sort-Object Name -Descending | Select-Object -First 1

    if ($localPackage) {
        $csomPath = Join-Path $localPackage.FullName "lib\net45"
    } else {
        Write-Host "Failed to download CSOM package" -ForegroundColor Red
        exit 1
    }
}

Write-Host "Using CSOM from: $csomPath" -ForegroundColor Green

# Load CSOM assemblies
Add-Type -Path "$csomPath\Microsoft.SharePoint.Client.dll"
Add-Type -Path "$csomPath\Microsoft.SharePoint.Client.Runtime.dll"

Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking

# Convert email to OneDrive URL format
$userFormatted = $UserEmail -replace '@', '_' -replace '\.', '_'
$oneDriveUrl = "https://$TenantName-my.sharepoint.com/personal/$userFormatted"

Write-Host "Connecting to SharePoint Online admin center..." -ForegroundColor Cyan
$adminUrl = "https://$TenantName-admin.sharepoint.com"

# Get credentials once
$credentials = Get-Credential -Message "Enter your SharePoint admin credentials"

try {
    # Connect to SharePoint Online Admin using SPO module
    Connect-SPOService -Url $adminUrl -Credential $credentials

    Write-Host "Granting site collection admin access to $AdminEmail on OneDrive: $oneDriveUrl" -ForegroundColor Cyan
    Set-SPOUser -Site $oneDriveUrl -LoginName $AdminEmail -IsSiteCollectionAdmin $true

    Write-Host "Waiting 10 seconds for permissions to propagate..." -ForegroundColor Yellow
    Start-Sleep -Seconds 10

    # Disconnect from SPO
    Disconnect-SPOService

    # Create destination folder if it doesn't exist
    if (-not (Test-Path $DestinationPath)) {
        New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
    }

    Write-Host "Connecting to user's OneDrive using CSOM: $oneDriveUrl" -ForegroundColor Cyan

    # Create SharePoint context using CSOM
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($oneDriveUrl)
    $context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($credentials.UserName, $credentials.Password)

    # Get the Documents library
    $web = $context.Web
    $list = $web.Lists.GetByTitle("Documents")
    $context.Load($web)
    $context.Load($list)
    $context.ExecuteQuery()

    Write-Host "Starting download of all files from OneDrive..." -ForegroundColor Cyan
    Write-Host "Destination: $DestinationPath" -ForegroundColor Green

    # Function to download files recursively
    function Download-Folder {
        param(
            [Microsoft.SharePoint.Client.ClientContext]$ctx,
            [string]$folderUrl,
            [string]$localPath
        )

        $folder = $ctx.Web.GetFolderByServerRelativeUrl($folderUrl)
        $ctx.Load($folder)
        $ctx.Load($folder.Files)
        $ctx.Load($folder.Folders)
        $ctx.ExecuteQuery()

        # Download files
        foreach ($file in $folder.Files) {
            $fileName = $file.Name
            $fileRef = $file.ServerRelativeUrl

            $localFilePath = Join-Path $localPath $fileName

            Write-Host "Downloading: $fileRef" -ForegroundColor Gray

            try {
                $fileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($ctx, $fileRef)
                $ctx.ExecuteQuery()

                $fileStream = [System.IO.File]::Create($localFilePath)
                $fileInfo.Stream.CopyTo($fileStream)
                $fileStream.Close()

                Write-Host "Downloaded: $fileName" -ForegroundColor Green
            }
            catch {
                Write-Host "Failed to download: $fileName - $($_.Exception.Message)" -ForegroundColor Red
            }
        }

        # Process subfolders
        foreach ($subfolder in $folder.Folders) {
            if ($subfolder.Name -ne "Forms") {
                $subfolderPath = Join-Path $localPath $subfolder.Name
                if (-not (Test-Path $subfolderPath)) {
                    New-Item -ItemType Directory -Path $subfolderPath -Force | Out-Null
                }
                Download-Folder -ctx $ctx -folderUrl $subfolder.ServerRelativeUrl -localPath $subfolderPath
            }
        }
    }

    # Start downloading from root Documents folder
    $rootFolder = $list.RootFolder
    $context.Load($rootFolder)
    $context.ExecuteQuery()

    Download-Folder -ctx $context -folderUrl $rootFolder.ServerRelativeUrl -localPath $DestinationPath

    Write-Host "`nDownload complete!" -ForegroundColor Green
    Write-Host "Files saved to: $DestinationPath" -ForegroundColor Cyan

    # Cleanup
    $context.Dispose()

    # Optional: Remove admin access after completion
    $removeAccess = Read-Host "`nDo you want to remove admin access from $AdminEmail? (Y/N)"
    if ($removeAccess -eq 'Y') {
        Connect-SPOService -Url $adminUrl -Credential $credentials
        Set-SPOUser -Site $oneDriveUrl -LoginName $AdminEmail -IsSiteCollectionAdmin $false
        Write-Host "Admin access removed." -ForegroundColor Green
        Disconnect-SPOService
    }
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    if ($context) { $context.Dispose() }
    Disconnect-SPOService -ErrorAction SilentlyContinue
    exit 1
}
