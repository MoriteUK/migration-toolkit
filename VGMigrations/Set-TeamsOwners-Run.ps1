#Requires -Version 7.0
<#
.SYNOPSIS
    Adds a user as owner to Teams/M365 Groups from a CSV (headless, streaming output).
.PARAMETER CsvFile
    Path to CSV with an "address" column containing team/group email addresses.
.PARAMETER OwnerUpn
    UPN of the user to add as owner.
.PARAMETER WhatIf
    Preview only — resolve groups and report what would change without making changes.
#>
param(
    [Parameter(Mandatory=$true)]  [string]$CsvFile,
    [Parameter(Mandatory=$true)]  [string]$OwnerUpn,
    [Parameter(Mandatory=$false)] [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

Write-Host "=== Set Teams Owners ===" -ForegroundColor Cyan
Write-Host "CSV:   $CsvFile"
Write-Host "Owner: $OwnerUpn"
if ($WhatIf) { Write-Host "Mode:  WhatIf (no changes will be made)" -ForegroundColor Yellow }

# Read CSV
if (-not (Test-Path $CsvFile)) { Write-Error "CSV not found: $CsvFile"; exit 1 }
$entries = @(Import-Csv -Path $CsvFile -Encoding UTF8 -ErrorAction Stop)
if ($entries.Count -eq 0) { Write-Error "CSV is empty."; exit 1 }

$cols = @($entries[0].PSObject.Properties.Name)
Write-Host "Columns: $($cols -join ', ')"
if ($cols -notcontains 'address') {
    Write-Error "CSV must have an 'address' column. Found: $($cols -join ', ')"
    exit 1
}

$addresses = @($entries | ForEach-Object { $_.address.Trim() } | Where-Object { $_ })
if ($addresses.Count -eq 0) { Write-Error "No addresses found in CSV."; exit 1 }
Write-Host "$($addresses.Count) address(es) loaded." -ForegroundColor Green

# Check modules
Write-Host "`nChecking Microsoft.Graph modules..." -ForegroundColor Cyan
. (Join-Path $PSScriptRoot 'Ensure-GraphModules.ps1') -GraphModules @('Microsoft.Graph.Groups','Microsoft.Graph.Users')

# Connect
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Write-Host "(A sign-in window will open — please authenticate)" -ForegroundColor Yellow
Connect-MgGraph -Scopes 'Group.ReadWrite.All','GroupMember.ReadWrite.All','User.Read.All' -ErrorAction Stop
Write-Host "Connected to Microsoft Graph." -ForegroundColor Green

# Resolve owner
Write-Host "Resolving owner: $OwnerUpn" -ForegroundColor Cyan
$targetUser = Get-MgUser -Filter "userPrincipalName eq '$OwnerUpn'" -Property Id,DisplayName -ErrorAction Stop |
              Select-Object -First 1
if (-not $targetUser) { Write-Error "User not found: $OwnerUpn"; exit 1 }
Write-Host "Owner resolved: $($targetUser.DisplayName) [$($targetUser.Id)]" -ForegroundColor Green

# Process each group
$added   = 0
$skipped = 0
$failed  = 0
$total   = $addresses.Count
$i       = 0
$selectProps = 'id,displayName,mail,mailNickname,groupTypes,resourceProvisioningOptions'

foreach ($address in $addresses) {
    $i++
    $mailNick = $address -replace '@.*', ''

    try {
        $group = Get-MgGroup -Filter "mail eq '$address'" -Property $selectProps -ErrorAction Stop |
                 Select-Object -First 1
        if (-not $group) {
            $group = Get-MgGroup -Filter "mailNickname eq '$mailNick'" -Property $selectProps -ErrorAction Stop |
                     Select-Object -First 1
        }
        if (-not $group) {
            Write-Warning "[$i/$total] Not found: $address"
            $failed++; continue
        }

        $isTeam    = $group.AdditionalProperties['resourceProvisioningOptions'] -contains 'Team'
        $isUnified = $group.GroupTypes -contains 'Unified'
        $label = if ($isTeam -and $isUnified) { 'Team/M365 Group' }
                 elseif ($isTeam)              { 'Team' }
                 elseif ($isUnified)           { 'M365 Group' }
                 else                          { 'Group' }

        $owners = @(Get-MgGroupOwner -GroupId $group.Id -Property Id | Select-Object -ExpandProperty Id)
        if ($owners -contains $targetUser.Id) {
            Write-Host "[$i/$total] Already owner [$label]: $($group.DisplayName)" -ForegroundColor Yellow
            $skipped++; continue
        }

        if ($WhatIf) {
            Write-Host "[$i/$total] WhatIf — would add as owner [$label]: $($group.DisplayName)" -ForegroundColor Yellow
            $skipped++; continue
        }

        $ref = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($targetUser.Id)" }
        New-MgGroupOwner  -GroupId $group.Id -BodyParameter $ref -ErrorAction Stop

        $members = @(Get-MgGroupMember -GroupId $group.Id -Property Id | Select-Object -ExpandProperty Id)
        if ($members -notcontains $targetUser.Id) {
            New-MgGroupMember -GroupId $group.Id -BodyParameter $ref -ErrorAction SilentlyContinue
        }

        Write-Host "[$i/$total] Added as owner [$label]: $($group.DisplayName)" -ForegroundColor Green
        $added++
    } catch {
        Write-Warning "[$i/$total] Failed: $address — $($_.Exception.Message.Split([Environment]::NewLine)[0])"
        $failed++
    }
}

Write-Host "`n=== Complete ===" -ForegroundColor Green
Write-Host "Added: $added   Already owner/WhatIf: $skipped   Failed: $failed"
if ($failed -gt 0) { Write-Warning "$failed group(s) could not be updated."; exit 1 }
