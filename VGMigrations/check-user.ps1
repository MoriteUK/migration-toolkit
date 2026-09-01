<#
.SYNOPSIS
    Finds every place an email address / UPN is in use across an Microsoft 365 tenant.

.DESCRIPTION
    AvePoint (and most migration tools) fail with "address already in use" when the
    target address exists ANYWHERE in the tenant as a PrimarySmtpAddress, an alias
    (proxyAddress), an X500 legacyExchangeDN, or a UserPrincipalName — including on
    objects that are easy to miss:
      - Active mailboxes (user, shared, room, equipment)
      - Mail contacts / mail users
      - Distribution groups and mail-enabled security groups
      - Microsoft 365 Groups / Teams
      - Mail-enabled public folders
      - SOFT-DELETED ("inactive") mailboxes still in the recoverable items window
      - Azure AD / Entra ID users whose UPN or proxyAddresses attribute matches,
        even if they have no mailbox at all

    This script checks all of the above and prints a report of every match, with the
    object type and which attribute matched, so you know exactly what to remediate
    before re-running the AvePoint job.

.REQUIREMENTS
    - ExchangeOnlineManagement module   (Install-Module ExchangeOnlineManagement)
    - Microsoft.Graph.Users module      (Install-Module Microsoft.Graph.Users)
    - An account with, at minimum, Exchange View-Only Recipients + Entra ID
      Directory Readers (Global Reader works for everything here).

.USAGE
    .\Find-AddressConflict.ps1 -Address "klea.feko@mbuAEPts.onmicrosoft.com"

    Add -IncludeSoftDeleted:$false to skip the soft-deleted mailbox scan if it's slow
    on a very large tenant.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Address,
    [switch]$IncludeSoftDeleted = $true
)
$ErrorActionPreference = 'Stop'
$results = New-Object System.Collections.Generic.List[object]

function Add-Result {
    param($ObjectType, $Identity, $MatchedOn, $Detail)
    $results.Add([pscustomobject]@{
        ObjectType = $ObjectType
        Identity   = $Identity
        MatchedOn  = $MatchedOn
        Detail     = $Detail
    })
}

Write-Host "Searching tenant for every occurrence of: $Address" -ForegroundColor Cyan

# ---------------------------------------------------------------
# 0. Connect
# ---------------------------------------------------------------
if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
    Connect-ExchangeOnline -ShowBanner:$false
}

try {
    Get-MgContext -ErrorAction Stop | Out-Null
} catch {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    Connect-MgGraph -Scopes "User.Read.All" -NoWelcome
}

$needle = $Address.Trim()
$needlePattern = [regex]::Escape($needle)

# ---------------------------------------------------------------
# 1. All mail-enabled recipients (mailboxes, contacts, mail users,
#    distribution/security/M365 groups, public folders — Get-Recipient
#    covers every recipient type in one pass)
# ---------------------------------------------------------------
Write-Host "Scanning all recipients (mailboxes, groups, contacts, public folders)..." -ForegroundColor Yellow

Get-Recipient -ResultSize Unlimited -Properties EmailAddresses, PrimarySmtpAddress, Alias, WindowsLiveID |
    ForEach-Object {
        $r = $_

        if ($r.PrimarySmtpAddress -and $r.PrimarySmtpAddress.ToString() -ieq $needle) {
            Add-Result $r.RecipientTypeDetails $r.Identity "PrimarySmtpAddress" $r.PrimarySmtpAddress
        }

        foreach ($ea in $r.EmailAddresses) {
            # EmailAddresses entries look like "smtp:alias@domain.com" or "X500:/o=..."
            if ($ea -imatch $needlePattern) {
                Add-Result $r.RecipientTypeDetails $r.Identity "EmailAddresses (proxyAddress)" $ea
            }
        }

        if ($r.WindowsLiveID -and $r.WindowsLiveID.ToString() -ieq $needle) {
            Add-Result $r.RecipientTypeDetails $r.Identity "WindowsLiveID (UPN)" $r.WindowsLiveID
        }
    }

# ---------------------------------------------------------------
# 2. Soft-deleted ("inactive") mailboxes — the #1 sneaky cause of
#    "address already in use" during migrations. These don't show
#    up in Get-Recipient at all.
# ---------------------------------------------------------------
if ($IncludeSoftDeleted) {
    Write-Host "Scanning soft-deleted mailboxes..." -ForegroundColor Yellow

    Get-Mailbox -SoftDeletedMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue |
        ForEach-Object {
            $mbx = Get-Mailbox -SoftDeletedMailbox -Identity $_.Identity -Properties EmailAddresses, PrimarySmtpAddress

            if ($mbx.PrimarySmtpAddress -and $mbx.PrimarySmtpAddress.ToString() -ieq $needle) {
                Add-Result "SoftDeletedMailbox" $mbx.Identity "PrimarySmtpAddress" $mbx.PrimarySmtpAddress
            }

            foreach ($ea in $mbx.EmailAddresses) {
                if ($ea -imatch $needlePattern) {
                    Add-Result "SoftDeletedMailbox" $mbx.Identity "EmailAddresses (proxyAddress)" $ea
                }
            }
        }

    # Also check the recoverable-items-only "inactive mailbox" set held under
    # litigation/retention hold, which Get-Mailbox -SoftDeletedMailbox can miss
    # in some tenant configurations.
    Get-Mailbox -InactiveMailboxOnly -ResultSize Unlimited -ErrorAction SilentlyContinue |
        ForEach-Object {
            if ($_.PrimarySmtpAddress -and $_.PrimarySmtpAddress.ToString() -ieq $needle) {
                Add-Result "InactiveMailbox" $_.Identity "PrimarySmtpAddress" $_.PrimarySmtpAddress
            }
            foreach ($ea in $_.EmailAddresses) {
                if ($ea -imatch $needlePattern) {
                    Add-Result "InactiveMailbox" $_.Identity "EmailAddresses (proxyAddress)" $ea
                }
            }
        }
}

# ---------------------------------------------------------------
# 3. Entra ID (Azure AD) users — catches accounts with the UPN or
#    proxyAddresses set but NO mailbox provisioned at all (common
#    right after a rename/merge, or with an unlicensed guest/user).
# ---------------------------------------------------------------
Write-Host "Scanning Entra ID users (UPN, proxyAddresses, mail)..." -ForegroundColor Yellow

$domainPart = ($needle -split '@')[-1]
Get-MgUser -All -Property "Id,UserPrincipalName,Mail,ProxyAddresses,OnPremisesSyncEnabled,AccountEnabled" `
    -Filter "endswith(userPrincipalName,'$domainPart') or endswith(mail,'$domainPart')" -ErrorAction SilentlyContinue |
    ForEach-Object {
        $u = $_

        if ($u.UserPrincipalName -ieq $needle) {
            Add-Result "EntraIDUser" $u.Id "UserPrincipalName" $u.UserPrincipalName
        }
        if ($u.Mail -ieq $needle) {
            Add-Result "EntraIDUser" $u.Id "Mail" $u.Mail
        }
        foreach ($pa in $u.ProxyAddresses) {
            if ($pa -imatch $needlePattern) {
                Add-Result "EntraIDUser" $u.Id "proxyAddresses" $pa
            }
        }
    }

# ---------------------------------------------------------------
# Report
# ---------------------------------------------------------------
Write-Host ""
if ($results.Count -eq 0) {
    Write-Host "No matches found for '$needle'. It should be free to use — if AvePoint still" -ForegroundColor Green
    Write-Host "rejects it, check for a stale entry in the AvePoint job cache/mapping itself," -ForegroundColor Green
    Write-Host "or an Exchange Online directory sync delay (can take up to a few hours)." -ForegroundColor Green
} else {
    Write-Host "Found $($results.Count) match(es) for '$needle':" -ForegroundColor Red
    $results | Sort-Object ObjectType, Identity | Format-Table -AutoSize -Wrap

    $out = ".\AddressConflict_$($needle -replace '[^\w]','_').csv"
    $results | Export-Csv -Path $out -NoTypeInformation
    Write-Host "Full results exported to $out" -ForegroundColor Cyan
}