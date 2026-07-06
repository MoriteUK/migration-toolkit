#Requires -Version 7.0
<#
.SYNOPSIS
    Renames email domain across all M365 objects in a Discovery folder and
    removes the old domain registration from the tenant.
.PARAMETER DiscoveryFolder
    Path to the Discovery folder produced by search-domain.ps1.
.PARAMETER OldDomain
    Domain to rename FROM (e.g. contoso.com).
.PARAMETER NewDomain
    Domain to rename TO (e.g. ourvolaris.onmicrosoft.com).
.PARAMETER Sections
    Comma-separated list of CSV filenames to process, or 'all' (default).
.PARAMETER WhatIf
    Preview only — no changes will be made.
#>
param(
    [Parameter(Mandatory=$true)]  [string]$DiscoveryFolder,
    [Parameter(Mandatory=$true)]  [string]$OldDomain,
    [Parameter(Mandatory=$true)]  [string]$NewDomain,
    [Parameter(Mandatory=$false)] [string]$Sections = 'all',
    [Parameter(Mandatory=$false)] [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

# Resolve discovery folder
$discFolder = $DiscoveryFolder.Trim().Trim('"')
$candidate  = Join-Path $discFolder 'Discovery'
if ((Split-Path $discFolder -Leaf) -ne 'Discovery' -and (Test-Path $candidate)) {
    $discFolder = $candidate
}
if (-not (Test-Path $discFolder)) { Write-Error "Discovery folder not found: $discFolder"; exit 1 }

$oldEsc = [regex]::Escape($OldDomain)

Write-Host "=== Rename Domain Objects$(if ($WhatIf) { ' [WhatIf]' }) ===" -ForegroundColor Cyan
Write-Host "Discovery folder : $discFolder"
Write-Host "Rename           : @$OldDomain  →  @$NewDomain"
Write-Host ''

# Load ExchangeOnlineManagement
$mod = Get-Module -ListAvailable -Name 'ExchangeOnlineManagement' -ErrorAction SilentlyContinue
if (-not $mod) {
    Write-Error "ExchangeOnlineManagement is not installed.`nRun: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    exit 1
}
Import-Module 'ExchangeOnlineManagement' -ErrorAction Stop

$script:exoConnected = $false

function Connect-EXOIfNeeded {
    if ($script:exoConnected) { return $true }
    Write-Host 'Connecting to Exchange Online — sign in when the browser opens...' -ForegroundColor Cyan
    $cmds = @('Get-Mailbox','Set-Mailbox','Get-DistributionGroup','Set-DistributionGroup',
              'Get-UnifiedGroup','Set-UnifiedGroup','Get-MailContact','Set-MailContact',
              'Get-Recipient','Remove-AcceptedDomain')
    try {
        Connect-ExchangeOnline -ShowBanner:$false -CommandName $cmds -ErrorAction Stop
        $script:exoConnected = $true
        Write-Host 'Exchange Online connected.' -ForegroundColor Green
        return $true
    } catch {
        Write-Host "ERROR: EXO connect failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" -ForegroundColor Red
        return $false
    }
}

# Builds a renamed EmailAddresses array:
#   Primary SMTP (SMTP:) → rename domain
#   Secondary smtp: with old domain → drop
#   Everything else → keep unchanged
function Get-RenamedAddresses {
    param([object[]]$Addresses)
    $newList = [System.Collections.Generic.List[string]]::new()
    $log     = [System.Collections.Generic.List[string]]::new()
    foreach ($a in ($Addresses | ForEach-Object { "$_" })) {
        $prefix = ($a -split ':')[0]
        if ($a -imatch "@$oldEsc$") {
            if ($prefix -ceq 'SMTP') {
                $renamed = $a -ireplace "@$oldEsc$", "@$NewDomain"
                $newList.Add($renamed)
                $log.Add("  PRIMARY : $a  →  $renamed")
            } else {
                $log.Add("  DROP    : $a")
            }
        } else {
            $newList.Add($a)
        }
    }
    return [pscustomobject]@{ Addresses = $newList.ToArray(); Log = $log }
}

$filter = if ($Sections -eq 'all') { $null } else { @($Sections -split ',') | ForEach-Object { $_.Trim() } }
function Should-Process { param([string]$n) -not $filter -or $n -in $filter }

# ── Mailboxes & Shared Mailboxes ──────────────────────────────────────────────
foreach ($csvName in @('02_Mailboxes.csv', '05_SharedMailboxes.csv')) {
    if (-not (Should-Process $csvName)) { continue }
    $csvPath = Join-Path $discFolder $csvName
    if (-not (Test-Path $csvPath)) { Write-Host "Skipped (not found): $csvName"; continue }
    $rows  = @(Import-Csv -Path $csvPath -Encoding UTF8)
    $label = if ($csvName -eq '02_Mailboxes.csv') { 'Mailboxes' } else { 'Shared Mailboxes' }
    Write-Host "--- $label: $($rows.Count) item(s) ---"
    if ($rows.Count -eq 0) { Write-Host ''; continue }
    if (-not $WhatIf -and -not (Connect-EXOIfNeeded)) { Write-Host ''; continue }
    $ok = 0; $fail = 0; $nochange = 0
    foreach ($r in $rows) {
        $id = if ($r.PSObject.Properties['UserPrincipalName'] -and $r.UserPrincipalName) { $r.UserPrincipalName } else { $r.PrimarySmtpAddress }
        if (-not $id) { continue }
        if ($WhatIf) { Write-Host "  WhatIf : $id"; $ok++; continue }
        try {
            $obj    = Get-Mailbox -Identity $id -ErrorAction Stop
            $result = Get-RenamedAddresses -Addresses $obj.EmailAddresses
            if ($result.Log.Count -eq 0) { Write-Host "  No change : $id"; $nochange++; continue }
            foreach ($l in $result.Log) { Write-Host $l }
            Set-Mailbox -Identity $id -EmailAddresses $result.Addresses -ErrorAction Stop
            Write-Host "  Renamed : $id" -ForegroundColor Green; $ok++
        } catch { Write-Host "  FAILED  : $id — $($_.Exception.Message.Split([Environment]::NewLine)[0])" -ForegroundColor Red; $fail++ }
    }
    Write-Host "  $label: $ok renamed  |  $fail failed  |  $nochange unchanged"
    Write-Host ''
}

# ── Distribution Groups ───────────────────────────────────────────────────────
if (Should-Process '03_DistributionGroups.csv') {
    $csvPath = Join-Path $discFolder '03_DistributionGroups.csv'
    if (-not (Test-Path $csvPath)) { Write-Host "Skipped (not found): 03_DistributionGroups.csv" }
    else {
        $rows = @(Import-Csv -Path $csvPath -Encoding UTF8)
        Write-Host "--- Distribution Groups: $($rows.Count) item(s) ---"
        if ($rows.Count -gt 0 -and ($WhatIf -or (Connect-EXOIfNeeded))) {
            $ok = 0; $fail = 0; $nochange = 0
            foreach ($r in $rows) {
                $addr = $r.PrimarySmtpAddress; if (-not $addr) { continue }
                if ($WhatIf) { Write-Host "  WhatIf : $addr"; $ok++; continue }
                try {
                    $obj    = Get-DistributionGroup -Identity $addr -ErrorAction Stop
                    $result = Get-RenamedAddresses -Addresses $obj.EmailAddresses
                    if ($result.Log.Count -eq 0) { Write-Host "  No change : $addr"; $nochange++; continue }
                    foreach ($l in $result.Log) { Write-Host $l }
                    Set-DistributionGroup -Identity $addr -EmailAddresses $result.Addresses -ErrorAction Stop
                    Write-Host "  Renamed : $addr" -ForegroundColor Green; $ok++
                } catch { Write-Host "  FAILED  : $addr — $($_.Exception.Message.Split([Environment]::NewLine)[0])" -ForegroundColor Red; $fail++ }
            }
            Write-Host "  Distribution Groups: $ok renamed  |  $fail failed  |  $nochange unchanged"
        }
        Write-Host ''
    }
}

# ── Mail Contacts ─────────────────────────────────────────────────────────────
if (Should-Process '04_MailContacts.csv') {
    $csvPath = Join-Path $discFolder '04_MailContacts.csv'
    if (-not (Test-Path $csvPath)) { Write-Host "Skipped (not found): 04_MailContacts.csv" }
    else {
        $rows = @(Import-Csv -Path $csvPath -Encoding UTF8)
        Write-Host "--- Mail Contacts: $($rows.Count) item(s) ---"
        if ($rows.Count -gt 0 -and ($WhatIf -or (Connect-EXOIfNeeded))) {
            $ok = 0; $fail = 0; $nochange = 0
            foreach ($r in $rows) {
                $addr = $r.ExternalEmailAddress; $dn = $r.DisplayName; if (-not $addr) { continue }
                if ($addr -imatch "@$oldEsc$") {
                    $newAddr = $addr -ireplace "@$oldEsc$", "@$NewDomain"
                    if ($WhatIf) { Write-Host "  WhatIf : $addr  →  $newAddr  [$dn]"; $ok++; continue }
                    try {
                        Set-MailContact -Identity $addr -ExternalEmailAddress $newAddr -ErrorAction Stop
                        Write-Host "  Renamed : $addr  →  $newAddr  [$dn]" -ForegroundColor Green; $ok++
                    } catch { Write-Host "  FAILED  : $addr — $($_.Exception.Message.Split([Environment]::NewLine)[0])" -ForegroundColor Red; $fail++ }
                } else { Write-Host "  No change : $addr  [$dn]"; $nochange++ }
            }
            Write-Host "  Mail Contacts: $ok renamed  |  $fail failed  |  $nochange unchanged"
        }
        Write-Host ''
    }
}

# ── M365 Groups + Teams ───────────────────────────────────────────────────────
if (Should-Process '06_M365Groups.csv') {
    $csvPath = Join-Path $discFolder '06_M365Groups.csv'
    if (-not (Test-Path $csvPath)) { Write-Host "Skipped (not found): 06_M365Groups.csv" }
    else {
        $rows = @(Import-Csv -Path $csvPath -Encoding UTF8)
        Write-Host "--- M365 Groups + Teams: $($rows.Count) item(s) ---"
        if ($rows.Count -gt 0 -and ($WhatIf -or (Connect-EXOIfNeeded))) {
            $ok = 0; $fail = 0; $nochange = 0
            foreach ($r in $rows) {
                $addr = $r.PrimarySmtpAddress; $dn = $r.DisplayName; if (-not $addr) { continue }
                if ($WhatIf) { Write-Host "  WhatIf : $addr  [$dn]"; $ok++; continue }
                try {
                    $obj    = Get-UnifiedGroup -Identity $addr -ErrorAction Stop
                    $result = Get-RenamedAddresses -Addresses $obj.EmailAddresses
                    if ($result.Log.Count -eq 0) { Write-Host "  No change : $addr  [$dn]"; $nochange++; continue }
                    foreach ($l in $result.Log) { Write-Host $l }
                    Set-UnifiedGroup -Identity $addr -EmailAddresses $result.Addresses -ErrorAction Stop
                    Write-Host "  Renamed : $addr  [$dn]" -ForegroundColor Green; $ok++
                } catch { Write-Host "  FAILED  : $addr — $($_.Exception.Message.Split([Environment]::NewLine)[0])" -ForegroundColor Red; $fail++ }
            }
            Write-Host "  M365 Groups + Teams: $ok renamed  |  $fail failed  |  $nochange unchanged"
        }
        Write-Host ''
    }
}

# ── Proxy Addresses — strip old-domain secondary smtp: entries ────────────────
if (Should-Process '12_ProxyAddresses.csv') {
    $csvPath = Join-Path $discFolder '12_ProxyAddresses.csv'
    if (-not (Test-Path $csvPath)) { Write-Host "Skipped (not found): 12_ProxyAddresses.csv" }
    else {
        $rows = @(Import-Csv -Path $csvPath -Encoding UTF8)
        Write-Host "--- Proxy Addresses: $($rows.Count) item(s) ---"
        if ($rows.Count -gt 0 -and ($WhatIf -or (Connect-EXOIfNeeded))) {
            $primaries    = @($rows | Where-Object { $_.IsPrimary -eq 'True' })
            $nonPrimaries = @($rows | Where-Object { $_.IsPrimary -ne 'True' })
            if ($primaries.Count -gt 0) {
                Write-Host "  Skipping $($primaries.Count) primary SMTP address(es) — renamed per-object above"
            }
            $byRecipient = $nonPrimaries | Group-Object -Property PrimarySmtpAddress
            $ok = 0; $fail = 0
            foreach ($grp in $byRecipient) {
                $primaryAddr     = $grp.Name
                $addressesToDrop = @($grp.Group | ForEach-Object { "$($_.AddressType):$($_.ProxyAddress)" })
                if ($WhatIf) {
                    foreach ($a in $addressesToDrop) { Write-Host "  WhatIf : remove proxy $a from $primaryAddr" }
                    $ok += $addressesToDrop.Count
                } else {
                    try {
                        $recip = Get-Recipient -Identity $primaryAddr -ErrorAction Stop
                        $currentProxies = @($recip.EmailAddresses | ForEach-Object { $_.ToString() })
                        $newProxies = @($currentProxies | Where-Object { $a = $_; -not ($addressesToDrop | Where-Object { $_ -ieq $a }) })
                        $removed = $currentProxies.Count - $newProxies.Count
                        if ($removed -eq 0) { Write-Host "  No matching proxies on $primaryAddr" }
                        else {
                            switch ($recip.RecipientTypeDetails) {
                                { $_ -in 'UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox' } {
                                    Set-Mailbox -Identity $primaryAddr -EmailAddresses $newProxies -ErrorAction Stop
                                }
                                'MailUniversalDistributionGroup' {
                                    Set-DistributionGroup -Identity $primaryAddr -EmailAddresses $newProxies -ErrorAction Stop
                                }
                                'GroupMailbox' {
                                    Set-UnifiedGroup -Identity $primaryAddr -EmailAddresses $newProxies -ErrorAction Stop
                                }
                                default { Write-Host "  SKIPPED $primaryAddr — type: $($recip.RecipientTypeDetails)" }
                            }
                            foreach ($dropped in $addressesToDrop) { Write-Host "  Removed proxy: $dropped  from: $primaryAddr" }
                            $ok += $removed
                        }
                    } catch { Write-Host "  FAILED : $primaryAddr — $($_.Exception.Message.Split([Environment]::NewLine)[0])" -ForegroundColor Red; $fail++ }
                }
            }
            Write-Host "  Proxy Addresses: $ok removed  |  $fail failed"
        }
        Write-Host ''
    }
}

# ── Accepted Domains — remove domain registration from tenant ─────────────────
if (Should-Process '01_AcceptedDomains.csv') {
    $csvPath = Join-Path $discFolder '01_AcceptedDomains.csv'
    if (-not (Test-Path $csvPath)) { Write-Host "Skipped (not found): 01_AcceptedDomains.csv" }
    else {
        $rows = @(Import-Csv -Path $csvPath -Encoding UTF8)
        Write-Host "--- Remove domain from tenant (Accepted Domains): $($rows.Count) item(s) ---"
        if ($rows.Count -gt 0 -and ($WhatIf -or (Connect-EXOIfNeeded))) {
            $removable = @($rows | Where-Object { $_.IsDefault -ne 'True' })
            $skipCount = $rows.Count - $removable.Count
            if ($skipCount -gt 0) { Write-Host "  Skipping $skipCount default/initial domain(s)" }
            $ok = 0; $fail = 0
            foreach ($r in $removable) {
                $name = $r.DomainName
                if ($WhatIf) { Write-Host "  WhatIf : remove domain $name from tenant"; $ok++; continue }
                try {
                    Remove-AcceptedDomain -Identity $name -Confirm:$false -ErrorAction Stop
                    Write-Host "  Removed domain from tenant: $name" -ForegroundColor Green; $ok++
                } catch { Write-Host "  FAILED : $name — $($_.Exception.Message.Split([Environment]::NewLine)[0])" -ForegroundColor Red; $fail++ }
            }
            Write-Host "  Accepted Domains: $ok removed  |  $fail failed"
        }
        Write-Host ''
    }
}

if ($script:exoConnected) { try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {} }
Write-Host '=== Done ===' -ForegroundColor Green
