<#
.SYNOPSIS
    Exports members of specific distribution groups to individual CSV files.

.DESCRIPTION
    Finds distribution groups matching a name pattern and exports each group's members
    to a separate CSV file (one file per group).

.PARAMETER NamePattern
    Pattern to match group names (e.g., 'MB CELCAT*'). Supports wildcards.

.PARAMETER OutputPath
    Folder for the CSV files. Defaults to the script folder.

.EXAMPLE
    .\export-group-members.ps1 -NamePattern 'MB CELCAT*'

.EXAMPLE
    .\export-group-members.ps1 -NamePattern 'MB *' -OutputPath 'C:\Exports'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$NamePattern,

    [string]$OutputPath = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    throw "ExchangeOnlineManagement module not found. Install with: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
}

if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:$false
}

# Ensure output directory exists
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

Write-Host "`nSearching for groups matching: $NamePattern" -ForegroundColor Cyan
$groups = @(Get-DistributionGroup -Filter "DisplayName -like '$NamePattern'" -ResultSize Unlimited)

if ($groups.Count -eq 0) {
    Write-Warning "No groups found matching pattern: $NamePattern"
    return
}

Write-Host "Found $($groups.Count) groups`n" -ForegroundColor Green

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$exportedFiles = @()

foreach ($group in $groups) {
    Write-Host "Processing: $($group.DisplayName)" -ForegroundColor Cyan

    try {
        # Get members
        $members = @(Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited)

        if ($members.Count -eq 0) {
            Write-Warning "  No members found in group: $($group.DisplayName)"
            continue
        }

        # Create safe filename from group name
        $safeFileName = $group.DisplayName -replace '[\\/:*?"<>|]', '_'
        $csvFile = Join-Path $OutputPath "${safeFileName}_Members_$timestamp.csv"

        # Export members
        $memberData = $members | Select-Object `
            @{N='GroupName'; E={$group.DisplayName}},
            @{N='GroupEmail'; E={$group.PrimarySmtpAddress}},
            DisplayName,
            PrimarySmtpAddress,
            RecipientTypeDetails

        $memberData | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8

        Write-Host "  ✓ Exported $($members.Count) members -> $csvFile" -ForegroundColor Green
        $exportedFiles += $csvFile

    } catch {
        Write-Warning "  Failed to export $($group.DisplayName): $($_.Exception.Message)"
    }
}

Write-Host "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "EXPORT SUMMARY" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "Groups found:    $($groups.Count)" -ForegroundColor White
Write-Host "Files exported:  $($exportedFiles.Count)" -ForegroundColor Green
Write-Host "`nFiles saved to:  $OutputPath" -ForegroundColor White
Write-Host ""
