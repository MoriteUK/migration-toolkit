#Requires -Version 7.0
<#
.SYNOPSIS
    Renames email domain across all M365 objects in a Discovery folder and
    removes the old domain registration from the tenant.
.PARAMETER DiscoveryFolder
    Path to the Discovery folder produced by search-domain.ps1.
    Optional — if omitted, rename operations are skipped and only OldDomain is removed from the tenant.
.PARAMETER OldDomain
    Domain to rename FROM (e.g. contoso.com).
.PARAMETER NewDomain
    Domain to rename TO (e.g. ourvolaris.onmicrosoft.com). Required unless DiscoveryFolder is omitted.
.PARAMETER Sections
    Comma-separated list of CSV filenames to process, or 'all' (default).
.PARAMETER WhatIf
    Preview only — no changes will be made.
.PARAMETER UserPrincipalName
    Optional. Admin account UPN to use for Exchange Online connection (e.g. admin@tenant.onmicrosoft.com).
    If specified, forces a fresh login with this account.
#>
param(
    [Parameter(Mandatory=$false)] [string]$DiscoveryFolder = '',
    [Parameter(Mandatory=$true)]  [string]$OldDomain,
    [Parameter(Mandatory=$false)] [string]$NewDomain = '',
    [Parameter(Mandatory=$false)] [string]$Sections = 'all',
    [Parameter(Mandatory=$false)] [switch]$WhatIf,
    [Parameter(Mandatory=$false)] [string]$UserPrincipalName = ''
)

$ErrorActionPreference = 'Stop'

# Resolve discovery folder — none means "domain removal only"
$discFolder = $DiscoveryFolder.Trim().Trim('"')
$domainOnly = -not $discFolder
if (-not $domainOnly) {
    $candidate  = Join-Path $discFolder 'Discovery'
    if ((Split-Path $discFolder -Leaf) -ne 'Discovery' -and (Test-Path $candidate)) {
        $discFolder = $candidate
    }
    if (-not (Test-Path $discFolder)) { Write-Error "Discovery folder not found: $discFolder"; exit 1 }
    if (-not $NewDomain) { Write-Error "NewDomain is required when renaming objects."; exit 1 }
}

$oldEsc = [regex]::Escape($OldDomain)

Write-Host "=== Rename Domain Objects$(if ($WhatIf) { ' [WhatIf]' }) ===" -ForegroundColor Cyan
if ($domainOnly) {
    Write-Host "Mode             : domain removal only (no Discovery folder — rename operations skipped)"
    Write-Host "Domain           : $OldDomain"
    $Sections = '01_AcceptedDomains.csv'
} else {
    Write-Host "Discovery folder : $discFolder"
    Write-Host "Rename           : @$OldDomain  →  @$NewDomain"
}
Write-Host ''

# Load ExchangeOnlineManagement (not needed for domain removal only — that uses Graph)
if (-not $domainOnly) {
    $mod = Get-Module -ListAvailable -Name 'ExchangeOnlineManagement' -ErrorAction SilentlyContinue
    if (-not $mod) {
        Write-Error "ExchangeOnlineManagement is not installed.`nRun: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
        exit 1
    }
    Import-Module 'ExchangeOnlineManagement' -ErrorAction Stop
}

$script:exoConnected = $false

function Connect-EXOIfNeeded {
    if ($script:exoConnected) { return $true }

    # Disconnect any existing session to force fresh login
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}

    $msg = if ($UserPrincipalName) { "Connecting to Exchange Online as $UserPrincipalName — sign in when prompted..." } else { 'Connecting to Exchange Online — sign in when the browser opens...' }
    Write-Host $msg -ForegroundColor Cyan

    $cmds = @('Get-Mailbox','Set-Mailbox','Get-DistributionGroup','Set-DistributionGroup',
              'Get-UnifiedGroup','Set-UnifiedGroup','Get-MailContact','Set-MailContact',
              'Get-Recipient')
    try {
        $connectParams = @{
            ShowBanner  = $false
            CommandName = $cmds
            ErrorAction = 'Stop'
        }
        if ($UserPrincipalName) {
            $connectParams.UserPrincipalName = $UserPrincipalName
        } else {
            $connectParams.DisableWAM = $true
        }
        Connect-ExchangeOnline @connectParams
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
foreach ($csvName in @('02_ADUsers.csv', '05_SharedMailboxes.csv')) {
    if (-not (Should-Process $csvName)) { continue }
    $csvPath = Join-Path $discFolder $csvName
    if (-not (Test-Path $csvPath)) { Write-Host "Skipped (not found): $csvName"; continue }
    $rows  = @(Import-Csv -Path $csvPath -Encoding UTF8)
    $label = if ($csvName -eq '02_ADUsers.csv') { 'Mailboxes' } else { 'Shared Mailboxes' }
    Write-Host "--- ${label}: $($rows.Count) item(s) ---"
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
    Write-Host "  ${label}: $ok renamed  |  $fail failed  |  $nochange unchanged"
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
if (Should-Process '13_ProxyAddresses.csv') {
    $csvPath = Join-Path $discFolder '13_ProxyAddresses.csv'
    if (-not (Test-Path $csvPath)) { Write-Host "Skipped (not found): 13_ProxyAddresses.csv" }
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
    $csvPath = if ($domainOnly) { $null } else { Join-Path $discFolder '01_AcceptedDomains.csv' }
    if (-not $domainOnly -and -not (Test-Path $csvPath)) { Write-Host "Skipped (not found): 01_AcceptedDomains.csv" }
    else {
        $rows = if ($domainOnly) { @([pscustomobject]@{ DomainName = $OldDomain; IsDefault = 'False' }) }
                else { @(Import-Csv -Path $csvPath -Encoding UTF8) }
        Write-Host "--- Remove domain from tenant (Accepted Domains): $($rows.Count) item(s) ---"
        # EXO accepted domains are synced from Entra ID — there is no Remove-AcceptedDomain in
        # Exchange Online. The domain must be deleted from Entra via Microsoft Graph instead.
        $removable = @($rows | Where-Object { $_.IsDefault -ne 'True' -and $_.DomainName -notmatch '\.onmicrosoft\.com$' })
        $skipCount = $rows.Count - $removable.Count
        if ($skipCount -gt 0) { Write-Host "  Skipping $skipCount default/initial domain(s)" }
        $graphReady = $WhatIf
        if ($removable.Count -gt 0 -and -not $WhatIf) {
            try {
                Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
                Write-Host 'Connecting to Microsoft Graph — sign in when prompted (needs Global Admin or Domain Name Administrator)...' -ForegroundColor Cyan
                Connect-MgGraph -Scopes 'Domain.ReadWrite.All','Directory.Read.All' -ContextScope Process -NoWelcome -ErrorAction Stop
                $graphReady = $true
            } catch {
                Write-Host "ERROR: Graph connect failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" -ForegroundColor Red
                Write-Host "  Remove the domain manually via M365 Admin Center: https://admin.microsoft.com/#/Domains" -ForegroundColor Yellow
            }
        }
        if ($graphReady) {
            $ok = 0; $fail = 0; $blocked = 0
            foreach ($r in $removable) {
                $name = $r.DomainName
                if ($WhatIf) { Write-Host "  WhatIf : remove domain $name from tenant"; $ok++; continue }
                try {
                    # Entra refuses the delete while any object still references the domain — list them first
                    $refs = @((Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/domains/$name/domainNameReferences?`$select=id,displayName" -ErrorAction Stop).value)
                    if ($refs.Count -gt 0) {
                        Write-Host "  BLOCKED : $name — still referenced by $($refs.Count) object(s):" -ForegroundColor Yellow
                        foreach ($ref in ($refs | Select-Object -First 25)) {
                            $type = ($ref.'@odata.type' -replace '^#microsoft\.graph\.', '')
                            Write-Host "    [$type] $($ref.displayName)"
                        }
                        if ($refs.Count -gt 25) { Write-Host "    ...and $($refs.Count - 25) more" }
                        Write-Host "    (UPNs, proxy addresses and app identifiers must all move off the domain first; EXO changes can take a while to sync to Entra)"
                        $blocked++; continue
                    }
                    Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/v1.0/domains/$name" -ErrorAction Stop | Out-Null
                    Write-Host "  Removed domain from tenant: $name" -ForegroundColor Green; $ok++
                } catch { Write-Host "  FAILED : $name — $($_.Exception.Message.Split([Environment]::NewLine)[0])" -ForegroundColor Red; $fail++ }
            }
            Write-Host "  Accepted Domains: $ok removed  |  $blocked blocked  |  $fail failed"
            if (-not $WhatIf) { try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {} }
        }
        Write-Host ''
    }
}

if ($script:exoConnected) { try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {} }
Write-Host '=== Done ===' -ForegroundColor Green
