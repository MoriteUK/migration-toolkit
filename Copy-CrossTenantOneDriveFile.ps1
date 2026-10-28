<#
.SYNOPSIS
    Copies a single file from a user's OneDrive in one tenant to the same user's OneDrive in another tenant,
    using Microsoft Graph (Microsoft.Graph.Authentication module only - no PnP).
    Use for items AvePoint FLY failed/skipped. Copies the current version only (no version history).

.DESCRIPTION
    Two auth modes per tenant:
      - App-only (certificate): pass -SourceClientId/-SourceThumbprint (and Target equivalents).
        Source app needs Graph application permission Files.Read.All, target app Files.ReadWrite.All
        (admin consented). No OneDrive access rights needed for your account.
      - Interactive (delegated): omit the thumbprint. Signs in with scope Files.ReadWrite.All; your account
        must have access to the user's OneDrive (e.g. site collection admin).

    Created/Modified dates are preserved. "Modified By" will show the app (app-only) or you (interactive) -
    Graph cannot set Author/Editor.

.EXAMPLE
    .\Copy-CrossTenantOneDriveFile.ps1 `
        -SourceTenant contoso.onmicrosoft.com -SourceClientId <appId> -SourceThumbprint <thumb> `
        -SourceUpn jbloggs@contoso.com `
        -TargetTenant fabrikam.onmicrosoft.com -TargetClientId <appId> -TargetThumbprint <thumb> `
        -TargetUpn jbloggs@fabrikam.com `
        -FilePath "Finance/Budget 2026.xlsx"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SourceTenant,      # contoso.onmicrosoft.com or tenant GUID
    [string]$SourceClientId,
    [string]$SourceThumbprint,
    [Parameter(Mandatory)][string]$SourceUpn,

    [Parameter(Mandatory)][string]$TargetTenant,      # fabrikam.onmicrosoft.com or tenant GUID
    [string]$TargetClientId,
    [string]$TargetThumbprint,
    [Parameter(Mandatory)][string]$TargetUpn,

    # Path relative to the OneDrive root (no "Documents/"), e.g. "Finance/Budget 2026.xlsx"
    [Parameter(Mandatory)][string]$FilePath,
    # Target path relative to the OneDrive root; defaults to the same path. Missing folders are created.
    [string]$TargetPath,

    [string]$TempPath = "C:\Temp\xfer"
)

$ErrorActionPreference = 'Stop'

function Connect-Tenant([string]$Tenant, [string]$ClientId, [string]$Thumbprint) {
    try { Disconnect-MgGraph | Out-Null } catch {}
    if ($Thumbprint) {
        Connect-MgGraph -TenantId $Tenant -ClientId $ClientId -CertificateThumbprint $Thumbprint -NoWelcome
    } elseif ($ClientId) {
        Connect-MgGraph -TenantId $Tenant -ClientId $ClientId -Scopes 'Files.ReadWrite.All' -NoWelcome
    } else {
        Connect-MgGraph -TenantId $Tenant -Scopes 'Files.ReadWrite.All' -NoWelcome
    }
}

function ConvertTo-GraphPath([string]$Path) {
    ($Path.Trim('/') -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
}

if (-not $TargetPath) { $TargetPath = $FilePath }
$srcEnc = ConvertTo-GraphPath $FilePath
$dstEnc = ConvertTo-GraphPath $TargetPath
$name   = Split-Path $FilePath -Leaf
$local  = Join-Path $TempPath $name
New-Item -ItemType Directory -Force $TempPath | Out-Null

# --- Source tenant: read metadata and download ---
Write-Host "Source: $SourceUpn :/$FilePath"
Connect-Tenant $SourceTenant $SourceClientId $SourceThumbprint
$item = Invoke-MgGraphRequest -Method GET -Uri "v1.0/users/$SourceUpn/drive/root:/${srcEnc}"
$created  = ([datetime]$item.fileSystemInfo.createdDateTime).ToUniversalTime().ToString('o')
$modified = ([datetime]$item.fileSystemInfo.lastModifiedDateTime).ToUniversalTime().ToString('o')
Invoke-MgGraphRequest -Method GET -Uri "v1.0/users/$SourceUpn/drive/root:/${srcEnc}:/content" -OutputFilePath $local

# --- Target tenant: upload via upload session (any size) with original dates ---
Write-Host "Target: $TargetUpn :/$TargetPath"
Connect-Tenant $TargetTenant $TargetClientId $TargetThumbprint
$size = (Get-Item $local).Length

if ($size -eq 0) {
    Invoke-MgGraphRequest -Method PUT -Uri "v1.0/users/$TargetUpn/drive/root:/${dstEnc}:/content" -Body ([byte[]]@()) -ContentType 'application/octet-stream' | Out-Null
} else {
    $body = @{
        item = @{
            '@microsoft.graph.conflictBehavior' = 'replace'
            fileSystemInfo = @{ createdDateTime = $created; lastModifiedDateTime = $modified }
        }
    } | ConvertTo-Json -Depth 5
    $session = Invoke-MgGraphRequest -Method POST -Uri "v1.0/users/$TargetUpn/drive/root:/${dstEnc}:/createUploadSession" -Body $body -ContentType 'application/json'

    # Chunks must be a multiple of 320 KiB; uploadUrl is pre-authenticated (no bearer token)
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
            Invoke-RestMethod -Method PUT -Uri $session.uploadUrl -Body $part -Headers @{ 'Content-Range' = $range } -ContentType 'application/octet-stream' | Out-Null
            $pos += $read
            Write-Progress -Activity "Uploading $name" -PercentComplete ([int]($pos / $size * 100))
        }
    } finally {
        $stream.Dispose()
    }
}

Disconnect-MgGraph | Out-Null
Remove-Item $local -Force
Write-Host "Copied $name ($size bytes, Modified $modified)" -ForegroundColor Green
