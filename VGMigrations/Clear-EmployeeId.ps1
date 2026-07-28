#Requires -Version 5.1
<#
.SYNOPSIS
    Clear-EmployeeId.ps1 — Clears the on-premise Active Directory EmployeeID attribute for
    every user listed in an AvePoint mapping CSV file.

.DESCRIPTION
    Reads a source/destination mapping CSV — the same kind of file used to import mappings into
    an AvePoint Fly migration project (Exchange, OneDrive, etc.) — auto-detects the "Source"
    identity column (email/UPN), looks each user up in on-premise Active Directory by
    UserPrincipalName or mail, and clears (blanks) their EmployeeID attribute.

    This is an on-premise AD operation, not a cloud/Entra ID one — it must run on a host that
    can reach a domain controller (a Domain Controller itself, or any machine with the
    ActiveDirectory RSAT module installed).

.PARAMETER MappingFile
    Path to the AvePoint mapping CSV file. Must contain a column whose header starts with
    "Source" (e.g. "Source email address", "Source user", "Source UPN").

.PARAMETER WhatIf
    Preview changes without applying them.

.EXAMPLE
    .\Clear-EmployeeId.ps1 -MappingFile "C:\mappings\exchange.csv" -WhatIf
#>

param(
    [Parameter(Mandatory)]
    [string]$MappingFile,

    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

Write-Host "=== Clear-EmployeeId ===" -ForegroundColor Cyan
Write-Host "Mapping file: $MappingFile" -ForegroundColor White
if ($WhatIf) { Write-Host "Mode: WhatIf (preview only — no changes will be made)" -ForegroundColor Yellow }
Write-Host ""

if (-not (Test-Path $MappingFile)) {
    Write-Host "ERROR: Mapping file not found: $MappingFile" -ForegroundColor Red
    exit 1
}

if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "ERROR: ActiveDirectory PowerShell module is not installed." -ForegroundColor Red
    Write-Host "This is an on-premise AD module — install RSAT or run this from a Domain Controller." -ForegroundColor Yellow
    Write-Host "Install: Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'" -ForegroundColor Yellow
    exit 1
}
Import-Module ActiveDirectory -ErrorAction Stop

# Detect CSV encoding from BOM so non-ASCII characters are read correctly regardless of how
# AvePoint/Excel saved the file — same helper as Import-FlyMappings.ps1.
function Get-CsvEncoding([string]$Path) {
    try {
        $fs    = [System.IO.FileStream]::new($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $bytes = New-Object byte[] 4
        $read  = $fs.Read($bytes, 0, 4)
        $fs.Close()
        if ($read -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { return 'UTF8' }
        if ($read -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { return 'Unicode' }
        if ($read -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) { return 'BigEndianUnicode' }
    } catch {
        Write-Host "Could not detect CSV encoding ($($_.Exception.Message)) — defaulting to system ANSI." -ForegroundColor Yellow
    }
    return 'Default'
}

$encoding = Get-CsvEncoding $MappingFile
$rows = @(Import-Csv -Path $MappingFile -Encoding $encoding)

if ($rows.Count -eq 0) {
    Write-Host "Mapping file has no rows. Nothing to do." -ForegroundColor Yellow
    exit 0
}

# Detect the "Source" identity column dynamically — different workload exports use different
# headers (e.g. "Source email address", "Source user", "Source UPN"). Prefer identity-looking
# columns (email/upn/user) over plain name columns, same convention as Import-FlyMappings.ps1.
$headers = $rows[0].PSObject.Properties.Name
$srcCol = $headers | Where-Object { $_ -imatch '^source' -and $_ -imatch 'email|upn|user' } | Select-Object -First 1
if (-not $srcCol) { $srcCol = $headers | Where-Object { $_ -imatch '^source' } | Select-Object -First 1 }

if (-not $srcCol) {
    Write-Host "ERROR: Could not find a 'Source' column in the CSV. Columns found: $($headers -join ', ')" -ForegroundColor Red
    exit 1
}
Write-Host "Using '$srcCol' as the source identity column." -ForegroundColor Gray

$identities = @($rows | ForEach-Object { $_.$srcCol } | Where-Object { $_ } | Select-Object -Unique)
Write-Host "$($identities.Count) unique user(s) to process." -ForegroundColor Cyan
Write-Host ""

$cleared = 0; $alreadyClear = 0; $notFound = 0; $failed = 0

foreach ($identity in $identities) {
    $identity = $identity.Trim()
    try {
        $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$identity' -or mail -eq '$identity'" -Properties EmployeeID -ErrorAction Stop |
            Select-Object -First 1

        if (-not $adUser) {
            Write-Host "NOT FOUND: $identity" -ForegroundColor Yellow
            $notFound++
            continue
        }

        if ([string]::IsNullOrEmpty($adUser.EmployeeID)) {
            Write-Host "ALREADY CLEAR: $identity ($($adUser.SamAccountName))" -ForegroundColor DarkGray
            $alreadyClear++
            continue
        }

        if ($WhatIf) {
            Write-Host "WHATIF: would clear EmployeeID '$($adUser.EmployeeID)' for $identity ($($adUser.SamAccountName))" -ForegroundColor Yellow
            $cleared++
        } else {
            Set-ADUser -Identity $adUser -Clear EmployeeID -ErrorAction Stop
            Write-Host "CLEARED: $identity ($($adUser.SamAccountName)) — was '$($adUser.EmployeeID)'" -ForegroundColor Green
            $cleared++
        }
    } catch {
        Write-Host "FAILED: $identity — $($_.Exception.Message.Split([Environment]::NewLine)[0])" -ForegroundColor Red
        $failed++
    }
}

Write-Host ""
Write-Host "=== Done ===" -ForegroundColor Cyan
Write-Host "Cleared: $cleared  |  Already clear: $alreadyClear  |  Not found: $notFound  |  Failed: $failed"

if ($failed -gt 0) { exit 1 }
