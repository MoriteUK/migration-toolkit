<#
.SYNOPSIS
    Adds the X500 addresses exported by Export-SourceX500.ps1 to the matching objects in the
    TARGET tenant's on-premises Active Directory (proxyAddresses), so they sync up via Entra Connect.

.DESCRIPTION
    For each CSV row the target AD object (user, group or contact) is found by, in order:
      1. TargetIdentity column, if populated - matched against UPN, mail, sAMAccountName,
         an smtp: proxy address or a distinguishedName.
      2. -MatchBy (PrimarySmtpAddress or UserPrincipalName), with the source domain rewritten
         to the target domain when -SourceDomain / -TargetDomain are given.

    Only X500 addresses not already present in proxyAddresses are added (case-insensitive).
    Nothing existing is removed or changed. Supports -WhatIf.

    A result CSV is written for every row (Added / AlreadyPresent / NotFound / Ambiguous / Error).

.PARAMETER CsvPath
    CSV produced by Export-SourceX500.ps1.

.PARAMETER MatchBy
    Source column used to find the target object when TargetIdentity is blank.
    PrimarySmtpAddress (default) or UserPrincipalName.

.PARAMETER SourceDomain
    Domain to replace in the match value, e.g. fabrikam.com

.PARAMETER TargetDomain
    Replacement domain, e.g. contoso.com

.PARAMETER Server
    Optional DC to read from and write to.

.PARAMETER LogPath
    Result CSV path. Default: .\X500Import_<timestamp>.csv

.EXAMPLE
    .\Import-X500ToOnPremAD.ps1 -CsvPath .\SourceX500.csv -WhatIf
.EXAMPLE
    .\Import-X500ToOnPremAD.ps1 -CsvPath .\SourceX500.csv -SourceDomain fabrikam.com -TargetDomain contoso.com
.EXAMPLE
    .\Import-X500ToOnPremAD.ps1 -CsvPath .\SourceX500.csv -MatchBy UserPrincipalName -Server dc01.contoso.local
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [ValidateSet('PrimarySmtpAddress', 'UserPrincipalName')]
    [string]$MatchBy = 'PrimarySmtpAddress',

    [string]$SourceDomain,
    [string]$TargetDomain,
    [string]$Server,
    [string]$LogPath = ".\X500Import_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

Import-Module ActiveDirectory -ErrorAction Stop

if ([bool]$SourceDomain -ne [bool]$TargetDomain) {
    throw 'Specify both -SourceDomain and -TargetDomain, or neither.'
}

$adParams = @{}
if ($Server) { $adParams.Server = $Server }

# Escape a value for use inside an LDAP filter (RFC 4515)
function ConvertTo-LdapValue {
    param([string]$Value)
    $Value.Replace('\', '\5c').Replace('*', '\2a').Replace('(', '\28').Replace(')', '\29').Replace("`0", '\00')
}

function Convert-Domain {
    param([string]$Value)
    if (-not $SourceDomain -or -not $Value) { return $Value }
    $Value -replace "@$([regex]::Escape($SourceDomain))$", "@$TargetDomain"
}

$objectClassFilter = '(|(objectClass=user)(objectClass=group)(objectClass=contact))'

function Find-TargetObject {
    param([string]$Identity, [string]$Mode)

    $v = ConvertTo-LdapValue $Identity
    $clause = switch ($Mode) {
        'TargetIdentity'     { "(|(userPrincipalName=$v)(mail=$v)(sAMAccountName=$v)(proxyAddresses=smtp:$v)(distinguishedName=$v))" }
        'PrimarySmtpAddress' { "(|(mail=$v)(proxyAddresses=smtp:$v))" }
        'UserPrincipalName'  { "(userPrincipalName=$v)" }
    }
    @(Get-ADObject @adParams -LDAPFilter "(&$objectClassFilter$clause)" `
            -Properties proxyAddresses, mail, userPrincipalName, objectClass)
}

$rows = @(Import-Csv -Path $CsvPath)
Write-Host "Loaded $($rows.Count) row(s) from $CsvPath" -ForegroundColor Cyan

$results = New-Object System.Collections.Generic.List[object]
$i = 0

foreach ($row in $rows) {
    $i++
    Write-Progress -Activity 'Adding X500 addresses' -Status $row.DisplayName -PercentComplete (($i / $rows.Count) * 100)

    $result = [ordered]@{
        DisplayName      = $row.DisplayName
        SourcePrimary    = $row.PrimarySmtpAddress
        LookupValue      = $null
        TargetDN         = $null
        Status           = $null
        AddedCount       = 0
        AddedAddresses   = $null
        Message          = $null
    }

    try {
        # Work out what to look up
        if ($row.TargetIdentity) {
            $lookup = $row.TargetIdentity.Trim()
            $mode = 'TargetIdentity'
        } else {
            $lookup = Convert-Domain ([string]$row.$MatchBy).Trim()
            $mode = $MatchBy
        }
        $result.LookupValue = $lookup

        if (-not $lookup) {
            $result.Status = 'NotFound'
            $result.Message = "No $mode value in CSV"
            continue
        }

        $x500List = @(([string]$row.X500Addresses) -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ -like 'x500:*' })
        if ($x500List.Count -eq 0) {
            $result.Status = 'NoX500'
            $result.Message = 'Row has no X500 addresses'
            continue
        }

        $found = Find-TargetObject -Identity $lookup -Mode $mode
        if ($found.Count -eq 0) {
            $result.Status = 'NotFound'
            $result.Message = "No AD object matched $mode '$lookup'"
            continue
        }
        if ($found.Count -gt 1) {
            $result.Status = 'Ambiguous'
            $result.Message = "$($found.Count) objects matched: " + (($found.DistinguishedName) -join ' | ')
            continue
        }

        $target = $found[0]
        $result.TargetDN = $target.DistinguishedName

        $existing = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
        foreach ($p in @($target.proxyAddresses)) { [void]$existing.Add($p) }

        $toAdd = @($x500List | ForEach-Object { 'X500:' + $_.Substring(5) } |
                Where-Object { -not $existing.Contains($_) } | Sort-Object -Unique)

        if ($toAdd.Count -eq 0) {
            $result.Status = 'AlreadyPresent'
            continue
        }

        if ($PSCmdlet.ShouldProcess($target.DistinguishedName, "Add $($toAdd.Count) X500 address(es)")) {
            Set-ADObject @adParams -Identity $target.DistinguishedName -Add @{ proxyAddresses = [string[]]$toAdd } -ErrorAction Stop
            $result.Status = 'Added'
        } else {
            $result.Status = 'WhatIf'
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
Write-Host "Results written to $((Resolve-Path $LogPath).Path)" -ForegroundColor Green
