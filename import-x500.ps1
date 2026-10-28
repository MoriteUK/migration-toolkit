<#
.SYNOPSIS
    Adds X500 addresses to target AD users based on the Targetaddress column (Column D).

.DESCRIPTION
    For each CSV row, finds the target AD user by matching the Targetaddress column
    against the mail attribute or smtp: proxy addresses, then adds the X500 addresses
    to their proxyAddresses attribute.

    Only X500 addresses not already present are added (case-insensitive).
    Nothing existing is removed or changed. Supports -WhatIf.

    A result CSV is written for every row.

.PARAMETER CsvPath
    Path to the CSV file containing Targetaddress and X500Addresses columns.

.PARAMETER Server
    Optional DC to read from and write to.

.PARAMETER LogPath
    Result CSV path. Default: .\X500Import_<timestamp>.csv

.EXAMPLE
    .\import-x500.ps1 -CsvPath .\SourceX500.csv -WhatIf

.EXAMPLE
    .\import-x500.ps1 -CsvPath .\SourceX500.csv -Server dc01.contoso.local
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [string]$Server,
    [string]$LogPath = ".\X500Import_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

Import-Module ActiveDirectory -ErrorAction Stop

$adParams = @{}
if ($Server) { $adParams.Server = $Server }

$rows = @(Import-Csv -Path $CsvPath)
Write-Host "Loaded $($rows.Count) row(s) from $CsvPath" -ForegroundColor Cyan

$results = New-Object System.Collections.Generic.List[object]
$i = 0

foreach ($row in $rows) {
    $i++
    $displayName = if ($row.DisplayName) { $row.DisplayName } else { "Row $i" }
    Write-Progress -Activity 'Adding X500 addresses' -Status $displayName -PercentComplete (($i / $rows.Count) * 100)

    $result = [ordered]@{
        RowNumber        = $i
        DisplayName      = $displayName
        Targetaddress    = $null
        TargetDN         = $null
        Status           = $null
        AddedCount       = 0
        AddedAddresses   = $null
        Message          = $null
    }

    try {
        # Get the target address from Column D
        $targetAddress = $null
        if ($row.PSObject.Properties['Targetaddress'] -and $row.Targetaddress) {
            $targetAddress = ([string]$row.Targetaddress).Trim()
        }

        $result.Targetaddress = $targetAddress

        if (-not $targetAddress) {
            $result.Status = 'NoTargetAddress'
            $result.Message = 'No Targetaddress value in CSV row'
            continue
        }

        # Get X500 addresses from the row
        $x500List = @(([string]$row.X500Addresses) -split ';' |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -like 'x500:*' })

        if ($x500List.Count -eq 0) {
            $result.Status = 'NoX500'
            $result.Message = 'Row has no X500 addresses'
            continue
        }

        # Find the target AD user
        $users = @(Get-ADUser @adParams -Filter "mail -eq '$targetAddress' -or proxyAddresses -like 'smtp:$targetAddress'" `
            -Properties proxyAddresses, mail, userPrincipalName -ErrorAction Stop)

        if ($users.Count -eq 0) {
            $result.Status = 'NotFound'
            $result.Message = "No AD user matched '$targetAddress'"
            continue
        }

        if ($users.Count -gt 1) {
            $result.Status = 'Ambiguous'
            $result.Message = "$($users.Count) users matched: " + (($users.DistinguishedName) -join ' | ')
            continue
        }

        $user = $users[0]
        $targetDN = [string]$user.DistinguishedName
        $result.TargetDN = $targetDN

        if ([string]::IsNullOrWhiteSpace($targetDN)) {
            $result.Status = 'Error'
            $result.Message = 'Could not retrieve DistinguishedName from user object'
            continue
        }

        # Build list of existing proxy addresses (case-insensitive)
        $existing = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
        foreach ($p in @($user.proxyAddresses)) {
            [void]$existing.Add($p)
        }

        # Find X500 addresses that need to be added
        $toAdd = @($x500List |
            ForEach-Object { 'X500:' + $_.Substring(5) } |
            Where-Object { -not $existing.Contains($_) } |
            Sort-Object -Unique)

        if ($toAdd.Count -eq 0) {
            $result.Status = 'AlreadyPresent'
            $result.Message = 'All X500 addresses already present'
            continue
        }

        # Add the X500 addresses
        if ($PSCmdlet.ShouldProcess($targetDN, "Add $($toAdd.Count) X500 address(es)")) {
            Set-ADObject @adParams -Identity $targetDN `
                -Add @{ proxyAddresses = [string[]]$toAdd } -ErrorAction Stop
            $result.Status = 'Added'
            $result.Message = "Added $($toAdd.Count) X500 address(es)"
        } else {
            $result.Status = 'WhatIf'
            $result.Message = "Would add $($toAdd.Count) X500 address(es)"
        }

        $result.AddedCount = $toAdd.Count
        $result.AddedAddresses = $toAdd -join ';'
    }
    catch {
        $result.Status = 'Error'
        $result.Message = $_.Exception.Message
    }
    finally {
        $results.Add([pscustomobject]$result)
    }
}

Write-Progress -Activity 'Adding X500 addresses' -Completed

$results | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8

Write-Host ''
$results | Group-Object Status | Sort-Object Name |
    ForEach-Object { Write-Host ('  {0,-15} {1}' -f $_.Name, $_.Count) }
Write-Host ''
Write-Host "Results written to $((Resolve-Path $LogPath).Path)" -ForegroundColor Green
