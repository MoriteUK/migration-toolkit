<#
.SYNOPSIS
    Copies a file or folder (with all contents) from a user's OneDrive in one tenant to another tenant.

.DESCRIPTION
    Supports both individual files and entire folders (including OneNote notebooks).
    Recursively copies all files and subfolders, preserving the structure.

    Two auth modes per tenant:
      - App-only (certificate): pass -SourceClientId/-SourceThumbprint (and Target equivalents).
        Source app needs Graph application permission Files.Read.All, target app Files.ReadWrite.All
      - Interactive (delegated): omit the thumbprint. Signs in with scope Files.ReadWrite.All

    Created/Modified dates are preserved per file.

.EXAMPLE
    # Copy a OneNote notebook
    .\Copy-CrossTenantOneDriveFolder.ps1 `
        -SourceTenant contoso.onmicrosoft.com `
        -SourceUpn jbloggs@contoso.com `
        -TargetTenant fabrikam.onmicrosoft.com `
        -TargetUpn jbloggs@fabrikam.com `
        -Path "Bruno @ Volaris group"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SourceTenant,
    [string]$SourceClientId,
    [string]$SourceThumbprint,
    [Parameter(Mandatory)][string]$SourceUpn,

    [Parameter(Mandatory)][string]$TargetTenant,
    [string]$TargetClientId,
    [string]$TargetThumbprint,
    [Parameter(Mandatory)][string]$TargetUpn,

    # Path to file or folder relative to OneDrive root
    [Parameter(Mandatory)][string]$Path,
    # Target path; defaults to same path
    [string]$TargetPath,

    [string]$TempPath = "C:\Temp\xfer"
)

$ErrorActionPreference = 'Stop'

# Script-level variables for tenant info
$script:SourceTenant = $SourceTenant
$script:SourceClientId = $SourceClientId
$script:SourceThumbprint = $SourceThumbprint
$script:TargetTenant = $TargetTenant
$script:TargetClientId = $TargetClientId
$script:TargetThumbprint = $TargetThumbprint

function Connect-SourceTenant {
    try { Disconnect-MgGraph | Out-Null } catch {}
    if ($script:SourceThumbprint) {
        Connect-MgGraph -TenantId $script:SourceTenant -ClientId $script:SourceClientId `
            -CertificateThumbprint $script:SourceThumbprint -NoWelcome
    } elseif ($script:SourceClientId) {
        Connect-MgGraph -TenantId $script:SourceTenant -ClientId $script:SourceClientId `
            -Scopes 'Files.ReadWrite.All' -NoWelcome
    } else {
        Connect-MgGraph -TenantId $script:SourceTenant -Scopes 'Files.ReadWrite.All' -NoWelcome
    }
}

function Connect-TargetTenant {
    try { Disconnect-MgGraph | Out-Null } catch {}
    if ($script:TargetThumbprint) {
        Connect-MgGraph -TenantId $script:TargetTenant -ClientId $script:TargetClientId `
            -CertificateThumbprint $script:TargetThumbprint -NoWelcome
    } elseif ($script:TargetClientId) {
        Connect-MgGraph -TenantId $script:TargetTenant -ClientId $script:TargetClientId `
            -Scopes 'Files.ReadWrite.All' -NoWelcome
    } else {
        Connect-MgGraph -TenantId $script:TargetTenant -Scopes 'Files.ReadWrite.All' -NoWelcome
    }
}

function ConvertTo-GraphPath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    ($Path.Trim('/') -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
}

function Copy-File {
    param(
        [string]$SourcePath,
        [string]$TargetPath,
        [string]$SourceUpn,
        [string]$TargetUpn,
        [hashtable]$Item
    )

    $srcEnc = ConvertTo-GraphPath $SourcePath
    $dstEnc = ConvertTo-GraphPath $TargetPath
    $name = Split-Path $SourcePath -Leaf
    $local = Join-Path $TempPath $name

    Write-Host "  File: $SourcePath" -ForegroundColor Cyan

    # Download from source
    Connect-SourceTenant
    Invoke-MgGraphRequest -Method GET -Uri "v1.0/users/$SourceUpn/drive/root:/${srcEnc}:/content" -OutputFilePath $local

    # Upload to target
    Connect-TargetTenant
    $size = (Get-Item $local).Length
    $created = ([datetime]$Item.fileSystemInfo.createdDateTime).ToUniversalTime().ToString('o')
    $modified = ([datetime]$Item.fileSystemInfo.lastModifiedDateTime).ToUniversalTime().ToString('o')

    if ($size -eq 0) {
        Invoke-MgGraphRequest -Method PUT -Uri "v1.0/users/$TargetUpn/drive/root:/${dstEnc}:/content" `
            -Body ([byte[]]@()) -ContentType 'application/octet-stream' | Out-Null
    } else {
        $body = @{
            item = @{
                '@microsoft.graph.conflictBehavior' = 'replace'
                fileSystemInfo = @{ createdDateTime = $created; lastModifiedDateTime = $modified }
            }
        } | ConvertTo-Json -Depth 5

        $session = Invoke-MgGraphRequest -Method POST -Uri "v1.0/users/$TargetUpn/drive/root:/${dstEnc}:/createUploadSession" `
            -Body $body -ContentType 'application/json'

        # Upload in chunks (10MB each)
        $chunkSize = 320KB * 32
        $buffer = New-Object byte[] $chunkSize
        $stream = [IO.File]::OpenRead($local)
        try {
            $pos = 0
            while ($pos -lt $size) {
                $read = $stream.Read($buffer, 0, $chunkSize)
                $part = New-Object byte[] $read
                [Array]::Copy($buffer, $part, $read)
                $range = "bytes $pos-$($pos + $read - 1)/$size"
                Invoke-RestMethod -Method PUT -Uri $session.uploadUrl -Body $part `
                    -Headers @{ 'Content-Range' = $range } -ContentType 'application/octet-stream' | Out-Null
                $pos += $read
            }
        } finally {
            $stream.Dispose()
        }
    }

    Remove-Item $local -Force
    Write-Host "    ✓ Copied $name ($size bytes)" -ForegroundColor Green
}

function Copy-FolderRecursive {
    param(
        [string]$SourcePath,
        [string]$TargetPath,
        [string]$SourceUpn,
        [string]$TargetUpn
    )

    $srcEnc = ConvertTo-GraphPath $SourcePath

    Write-Host "Folder: $SourcePath" -ForegroundColor Yellow

    # Connect to source to get children
    Connect-SourceTenant

    # Get all children of this folder
    $uri = if ($SourcePath) {
        "v1.0/users/$SourceUpn/drive/root:/${srcEnc}:/children"
    } else {
        "v1.0/users/$SourceUpn/drive/root/children"
    }

    $children = @()
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri
    $children += $response.value

    # Handle pagination
    while ($response.'@odata.nextLink') {
        $response = Invoke-MgGraphRequest -Method GET -Uri $response.'@odata.nextLink'
        $children += $response.value
    }

    foreach ($child in $children) {
        $childSourcePath = if ($SourcePath) { "$SourcePath/$($child.name)" } else { $child.name }
        $childTargetPath = if ($TargetPath) { "$TargetPath/$($child.name)" } else { $child.name }

        if ($child.folder) {
            # It's a folder - recurse
            Copy-FolderRecursive -SourcePath $childSourcePath -TargetPath $childTargetPath `
                -SourceUpn $SourceUpn -TargetUpn $TargetUpn
        } else {
            # It's a file - copy it
            Copy-File -SourcePath $childSourcePath -TargetPath $childTargetPath `
                -SourceUpn $SourceUpn -TargetUpn $TargetUpn -Item $child
        }
    }
}

# Main script
if (-not $TargetPath) { $TargetPath = $Path }
New-Item -ItemType Directory -Force $TempPath | Out-Null

Write-Host "Source: $SourceUpn :/$Path"
Write-Host "Target: $TargetUpn :/$TargetPath"
Write-Host ""

# Connect to source tenant and check if it's a file or folder
Connect-SourceTenant

$srcEnc = ConvertTo-GraphPath $Path
try {
    $item = Invoke-MgGraphRequest -Method GET -Uri "v1.0/users/$SourceUpn/drive/root:/${srcEnc}"
} catch {
    Write-Error "Could not find '$Path' in $SourceUpn's OneDrive"
    exit 1
}

if ($item.folder) {
    # It's a folder - copy recursively
    Write-Host "Copying folder and all contents..." -ForegroundColor Cyan
    Copy-FolderRecursive -SourcePath $Path -TargetPath $TargetPath `
        -SourceUpn $SourceUpn -TargetUpn $TargetUpn
} else {
    # It's a single file
    Write-Host "Copying single file..." -ForegroundColor Cyan
    Copy-File -SourcePath $Path -TargetPath $TargetPath `
        -SourceUpn $SourceUpn -TargetUpn $TargetUpn -Item $item
}

Disconnect-MgGraph | Out-Null
Write-Host ""
Write-Host "Copy complete!" -ForegroundColor Green
