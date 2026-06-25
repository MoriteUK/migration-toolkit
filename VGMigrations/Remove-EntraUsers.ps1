#Requires -Version 7.0
<#
.SYNOPSIS
    Remove-EntraUsers.ps1 — Last-resort removal of Entra ID (Azure AD) cloud users
    tied to a domain. Exports full user details to a backup CSV before deleting so
    accounts can be recreated after the domain is fully removed.

.PARAMETER Domain
    The domain to target (e.g. olddomain.com).
    Finds all cloud-only users whose UPN ends with @domain.

.PARAMETER OutputFolder
    Folder to write the backup CSV and log to. Defaults to the script's logs\ folder.

.PARAMETER WhatIf
    Preview which users would be removed without making any changes.
    The backup CSV is always written regardless of WhatIf.

.NOTES
    On-premises synced users (OnPremisesSyncEnabled = true) are always skipped —
    remove them by updating the on-prem AD and syncing.

    Deleted users land in the Entra ID recycle bin for 30 days. The backup CSV
    also lets you recreate them manually or via a future recovery script.
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$Domain,
    [string]$OutputFolder = '',
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

$outDir = if ($OutputFolder) { $OutputFolder.Trim().Trim('"') } else { Join-Path $PSScriptRoot 'logs' }
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$csvPath   = Join-Path $outDir "entra-users-backup-$Domain-$timestamp.csv"
$logFile   = Join-Path $outDir "remove-entra-users-$timestamp.log"

function Log { param([string]$m) $ts = Get-Date -Format 'HH:mm:ss'; "$ts $m" | Tee-Object -FilePath $logFile -Append | Write-Host }

Log "=== Remove-EntraUsers$(if ($WhatIf) { ' [WhatIf]' }) ==="
Log "Domain     : @$Domain"
Log "Backup CSV : $csvPath"
Log "Log file   : $logFile"
Log ''

# ── Connect ────────────────────────────────────────────────────────────────────
Log 'Connecting to Microsoft Graph — sign in when prompted...'
Connect-MgGraph -Scopes 'User.ReadWrite.All','Group.Read.All','Directory.Read.All' `
    -UseDeviceCode -NoWelcome -ErrorAction Stop
Log 'Connected to Microsoft Graph.'
Log ''

# ── Find users ─────────────────────────────────────────────────────────────────
Log "Searching for cloud users with UPN ending @$Domain ..."
$props = @('Id','DisplayName','GivenName','Surname','UserPrincipalName','Mail',
           'AccountEnabled','Department','JobTitle','OfficeLocation',
           'MobilePhone','BusinessPhones','AssignedLicenses','ProxyAddresses',
           'OnPremisesSyncEnabled','CreatedDateTime','UsageLocation')
$users = @(Get-MgUser -Filter "endsWith(userPrincipalName,'@$Domain')" -All `
    -Property $props -ErrorAction Stop)

Log "Found $($users.Count) user(s) with UPN @$Domain"

$cloudUsers = @($users | Where-Object { $_.OnPremisesSyncEnabled -ne $true })
$syncedUsers = @($users | Where-Object { $_.OnPremisesSyncEnabled -eq $true })

if ($syncedUsers.Count -gt 0) {
    Log "  $($syncedUsers.Count) are on-premises synced — these will be SKIPPED (update in AD instead)"
}
Log "  $($cloudUsers.Count) cloud-only user(s) will be processed"
Log ''

if ($cloudUsers.Count -eq 0) {
    Log 'Nothing to remove.'
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    exit 0
}

# ── Load licence SKU names ─────────────────────────────────────────────────────
$skuMap = @{}
try {
    Get-MgSubscribedSku -ErrorAction Stop | ForEach-Object {
        $skuMap[[string]$_.SkuId] = $_.SkuPartNumber
    }
} catch { Log "WARNING: Could not load licence SKU names: $_" }

# ── Export backup CSV ──────────────────────────────────────────────────────────
Log 'Exporting user data to backup CSV...'

$exportRows = foreach ($u in $cloudUsers) {
    $licences = ($u.AssignedLicenses | ForEach-Object {
        if ($skuMap.ContainsKey([string]$_.SkuId)) { $skuMap[[string]$_.SkuId] } else { [string]$_.SkuId }
    }) -join ';'
    $proxies = ($u.ProxyAddresses) -join ';'
    $phones  = ($u.BusinessPhones) -join ';'

    $groups = ''
    try {
        $memberOf = @(Get-MgUserMemberOf -UserId $u.Id -ErrorAction Stop)
        $groups   = ($memberOf |
            Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' } |
            ForEach-Object { $_.AdditionalProperties['displayName'] }) -join ';'
    } catch { }

    [pscustomobject]@{
        ObjectId              = $u.Id
        DisplayName           = $u.DisplayName
        GivenName             = $u.GivenName
        Surname               = $u.Surname
        UserPrincipalName     = $u.UserPrincipalName
        Mail                  = $u.Mail
        AccountEnabled        = $u.AccountEnabled
        Department            = $u.Department
        JobTitle              = $u.JobTitle
        OfficeLocation        = $u.OfficeLocation
        MobilePhone           = $u.MobilePhone
        BusinessPhones        = $phones
        UsageLocation         = $u.UsageLocation
        CreatedDateTime       = $u.CreatedDateTime
        Licences              = $licences
        ProxyAddresses        = $proxies
        Groups                = $groups
    }
}

$exportRows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Log "Backup written: $($exportRows.Count) user(s) → $csvPath"
Log ''

# ── Remove users ───────────────────────────────────────────────────────────────
Log "--- $(if ($WhatIf) { 'WhatIf — no changes will be made' } else { 'Deleting users' }) ---"

$ok = 0; $fail = 0

foreach ($u in $cloudUsers) {
    if ($WhatIf) {
        Log "  WhatIf : would delete $($u.UserPrincipalName)  [$($u.DisplayName)]"
        $ok++
    } else {
        try {
            Remove-MgUser -UserId $u.Id -ErrorAction Stop
            Log "  Deleted : $($u.UserPrincipalName)  [$($u.DisplayName)]"
            $ok++
        } catch {
            Log "  FAILED  : $($u.UserPrincipalName) — $($_.Exception.Message.Split([Environment]::NewLine)[0])"
            $fail++
        }
    }
}

Log ''
Log "=== Complete: $(if ($WhatIf) { 'preview only' } else { "$ok deleted  |  $fail failed" }) ==="
Log "Backup CSV : $csvPath"
Log ''
Log 'Note: deleted users remain in the Entra ID recycle bin for 30 days.'
Log 'Note: use the backup CSV to recreate accounts after the domain is removed.'

Disconnect-MgGraph -ErrorAction SilentlyContinue
