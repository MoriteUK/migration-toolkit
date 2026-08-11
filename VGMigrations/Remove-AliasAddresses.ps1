#Requires -Version 7.0
<#
.SYNOPSIS
    Remove-AliasAddresses.ps1 — Strips alias (secondary smtp:) and IM (SIP/EUM) addresses from
    recipients listed in discovery CSVs, checking on-prem AD and Exchange Online.

.DESCRIPTION
    On-prem AD is tried first for hybrid/synced objects (everything except M365 Groups, which
    are cloud-native). Exchange Online is still checked too regardless of the on-prem result —
    some addresses (notably Teams/Skype-for-Business-Online auto-generated SIP addresses) are
    provisioned entirely cloud-side and never appear in on-prem proxyAddresses at all, even
    though the object itself is DirSync'd. When Exchange Online refuses an edit because the
    object is synchronized from on-premises, that's reported as "PROTECTED" rather than a
    generic failure — it means the address only exists in the cloud on an object EXO won't let
    you edit directly.

.PARAMETER DiscoveryFolder
    Path to the folder containing discovery CSV files (02_Mailboxes, 03_DistributionGroups, etc.)

.PARAMETER Domain
    Optional: restrict removal to addresses containing this domain (e.g. "olddomain.com").
    If omitted, all matching address types are removed regardless of domain.

.PARAMETER SkipAliases
    Skip removal of secondary SMTP alias addresses (smtp: prefix, lowercase).

.PARAMETER SkipSIP
    Skip removal of IM/SIP and EUM addresses (SIP: and EUM: prefixes).

.PARAMETER WhatIf
    Preview which addresses would be removed without making changes.

.NOTES
    Requires : ExchangeOnlineManagement
    CSVs used: 02_Mailboxes, 03_DistributionGroups, 04_MailContacts,
               05_SharedMailboxes, 06_M365Groups
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$DiscoveryFolder,

    [Parameter(Mandatory=$false)]
    [string]$Domain,

    [switch]$SkipAliases,

    [switch]$SkipSIP,

    [switch]$WhatIf
)

$script:RootDir = $PSScriptRoot

$_logDir = Join-Path $script:RootDir 'logs'
if (-not (Test-Path $_logDir)) { New-Item -ItemType Directory -Path $_logDir -Force | Out-Null }
$script:LogFile = Join-Path $_logDir "remove-aliases-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

function _RawLog {
    param([string]$Msg)
    "$(Get-Date -Format 'HH:mm:ss.fff') $Msg" | Add-Content -Path $script:LogFile -Encoding UTF8 -ErrorAction SilentlyContinue
}

_RawLog "=== Remove-AliasAddresses.ps1 started  PID=$PID  PSVersion=$($PSVersionTable.PSVersion) ==="

# OnPrem=$false for M365 Groups — cloud-native, no on-premises AD equivalent, always via EXO.
# Everything else may be a hybrid, AD-synced object — tried on-prem first (see
# Update-AliasesOnPrem), because Exchange Online outright refuses to modify EmailAddresses on a
# DirSync'd object ("out of the current user's write scope... should be performed on the object
# in your on-premises organization") regardless of which specific address is being changed.
$script:SectionDefs = @(
    [pscustomobject]@{ CsvName='02_Mailboxes.csv';          Label='Mailboxes';           KeyField='UserPrincipalName';    GetCmd='Get-Mailbox';           SetCmd='Set-Mailbox';           OnPrem=$true  }
    [pscustomobject]@{ CsvName='03_DistributionGroups.csv'; Label='Distribution Groups'; KeyField='PrimarySmtpAddress';   GetCmd='Get-DistributionGroup'; SetCmd='Set-DistributionGroup'; OnPrem=$true  }
    [pscustomobject]@{ CsvName='04_MailContacts.csv';       Label='Mail Contacts';       KeyField='ExternalEmailAddress'; GetCmd='Get-MailContact';       SetCmd='Set-MailContact';       OnPrem=$true  }
    [pscustomobject]@{ CsvName='05_SharedMailboxes.csv';    Label='Shared Mailboxes';    KeyField='PrimarySmtpAddress';   GetCmd='Get-Mailbox';           SetCmd='Set-Mailbox';           OnPrem=$true  }
    [pscustomobject]@{ CsvName='06_M365Groups.csv';         Label='M365 Groups';         KeyField='PrimarySmtpAddress';   GetCmd='Get-UnifiedGroup';      SetCmd='Set-UnifiedGroup';      OnPrem=$false }
)

$script:AdAvailable = $false
try { Import-Module ActiveDirectory -ErrorAction Stop; $script:AdAvailable = $true }
catch { }

# Looks the identity up in on-prem AD (by mail or UserPrincipalName) and, if found, applies the
# same Split-Addresses removal directly to its proxyAddresses attribute. Returns $null if the
# object isn't found on-prem (or AD isn't available) so the caller falls back to Exchange Online.
function Update-AliasesOnPrem {
    param([string]$Identity, [string]$FilterDomain, [bool]$DoAliases, [bool]$DoSIP, [switch]$WhatIfMode)
    if (-not $script:AdAvailable) { return $null }
    try {
        $obj = Get-ADObject -Filter "mail -eq '$Identity' -or userPrincipalName -eq '$Identity'" -Properties proxyAddresses -ErrorAction Stop | Select-Object -First 1
        if (-not $obj) { return $null }

        $current = @($obj.proxyAddresses)
        $split   = Split-Addresses -Addresses $current -FilterDomain $FilterDomain -DoAliases $DoAliases -DoSIP $DoSIP
        if ($split.Remove.Count -eq 0) { return [pscustomobject]@{ Found = $true; Removed = @() } }

        if (-not $WhatIfMode) {
            Set-ADObject -Identity $obj.DistinguishedName -Replace @{ proxyAddresses = $split.Keep } -ErrorAction Stop
        }
        return [pscustomobject]@{ Found = $true; Removed = $split.Remove }
    } catch {
        return $null
    }
}

# Returns which addresses to keep and which to remove.
# Keeps SMTP: (primary) and any type not selected for removal.
# If FilterDomain is set, only removes addresses containing that domain.
function Split-Addresses {
    param(
        [string[]]$Addresses,
        [string]$FilterDomain,
        [bool]$DoAliases,
        [bool]$DoSIP
    )

    $keep   = [System.Collections.Generic.List[string]]::new()
    $remove = [System.Collections.Generic.List[string]]::new()

    foreach ($addr in $Addresses) {
        $prefix = ($addr -split ':')[0]
        $isAlias = $DoAliases -and ($prefix -ceq 'smtp')
        $isIM    = $DoSIP     -and ($prefix -iin @('sip','eum'))

        if (($isAlias -or $isIM) -and (-not $FilterDomain -or $addr -like "*@$FilterDomain")) {
            $remove.Add($addr)
        } else {
            $keep.Add($addr)
        }
    }

    return [pscustomobject]@{ Keep = @($keep); Remove = @($remove) }
}

if (-not $DiscoveryFolder) {
    Write-Host 'Usage: Remove-AliasAddresses.ps1 -DiscoveryFolder <path> [-Domain <domain>] [-WhatIf]'
    exit 1
}

# ── Headless mode ──────────────────────────────────────────────────────────────
$discFolder = $DiscoveryFolder.Trim().Trim('"')
$candidate  = Join-Path $discFolder 'Discovery'
if ((Split-Path $discFolder -Leaf) -ne 'Discovery' -and (Test-Path $candidate)) {
    $discFolder = $candidate
}

if (-not (Test-Path $discFolder)) {
    Write-Host "ERROR: Discovery folder not found: $discFolder"
    exit 1
}

if ($SkipAliases -and $SkipSIP) {
    Write-Host 'ERROR: Nothing to remove — both types are skipped.'
    exit 1
}

$typeList = @()
if (-not $SkipAliases) { $typeList += 'smtp: aliases' }
if (-not $SkipSIP)     { $typeList += 'SIP/EUM addresses' }
$scopeMsg = ($typeList -join ' and ') + $(if ($Domain) { " matching @$Domain" } else { '' })

Write-Host "=== Remove Alias Addresses$(if ($WhatIf) { ' [WhatIf]' }) ==="
Write-Host "Discovery folder : $discFolder"
Write-Host "Removing         : $scopeMsg"
Write-Host $(if ($script:AdAvailable) { 'ActiveDirectory module loaded — on-prem objects checked first.' } else { 'ActiveDirectory module not available — all objects handled via Exchange Online only.' })
_RawLog "DiscoveryFolder=$discFolder  Domain=$Domain  SkipAliases=$SkipAliases  SkipSIP=$SkipSIP  WhatIf=$WhatIf  AdAvailable=$($script:AdAvailable)"

$mod = Get-Module -ListAvailable -Name 'ExchangeOnlineManagement' -ErrorAction SilentlyContinue
if (-not $mod) {
    Write-Host 'ERROR: ExchangeOnlineManagement module is not installed.'
    Write-Host 'Install it with: Install-Module ExchangeOnlineManagement -Scope CurrentUser'
    exit 1
}
Import-Module 'ExchangeOnlineManagement' -ErrorAction Stop
Write-Host 'ExchangeOnlineManagement module loaded.'

$exoConnected = $false

foreach ($sec in $script:SectionDefs) {
    $csvPath = Join-Path $discFolder $sec.CsvName
    if (-not (Test-Path $csvPath)) {
        Write-Host "Skipped (not found): $($sec.CsvName)"
        continue
    }

    $rows = @(Import-Csv -Path $csvPath -Encoding UTF8)
    Write-Host "--- $($sec.Label): $($rows.Count) item(s) ---"
    if ($rows.Count -eq 0) { continue }

    if (-not $exoConnected) {
        Write-Host 'Connecting to Exchange Online — sign in when the browser opens...'
        try {
            $cmds = @('Get-Mailbox','Set-Mailbox','Get-DistributionGroup','Set-DistributionGroup',
                      'Get-MailContact','Set-MailContact','Get-UnifiedGroup','Set-UnifiedGroup')
            Connect-ExchangeOnline -ShowBanner:$false -CommandName $cmds -DisableWAM -ErrorAction Stop
            $exoConnected = $true
            Write-Host 'Exchange Online connected.'
        } catch {
            Write-Host "ERROR: Failed to connect to Exchange Online: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
            exit 1
        }
    }

    $ok = 0; $fail = 0; $nochange = 0; $protected = 0

    foreach ($r in $rows) {
        $identity = $r.($sec.KeyField)
        if (-not $identity) { continue }

        # Try on-prem AD first. Note this alone isn't sufficient for hybrid objects: some
        # addresses (notably Teams/Skype-for-Business-Online auto-generated SIP addresses) are
        # provisioned entirely cloud-side and never appear in on-prem proxyAddresses at all, even
        # though the object itself is DirSync'd — so Exchange Online is still checked below
        # regardless of what happened on-prem.
        $onPremResult = $null
        if ($sec.OnPrem) {
            $onPremResult = Update-AliasesOnPrem -Identity $identity -FilterDomain $Domain -DoAliases (-not $SkipAliases) -DoSIP (-not $SkipSIP) -WhatIfMode:$WhatIf
            if ($onPremResult -and $onPremResult.Removed.Count -gt 0) {
                $removeList = $onPremResult.Removed -join ', '
                if ($WhatIf) {
                    Write-Host "  WhatIf    : $identity  would remove (on-prem AD): $removeList"
                    _RawLog "WHATIF-ONPREM $identity  remove=[$removeList]"
                } else {
                    Write-Host "  Removed   : $identity  (on-prem AD)  [$removeList]"
                    _RawLog "OK-ONPREM $identity  removed=[$removeList]"
                }
            }
        }

        try {
            $obj     = & $sec.GetCmd -Identity $identity -ErrorAction Stop
            $current = @($obj.EmailAddresses | ForEach-Object { "$_" })
            $split   = Split-Addresses -Addresses $current -FilterDomain $Domain -DoAliases (-not $SkipAliases) -DoSIP (-not $SkipSIP)

            # Exclude anything already removed on-prem — Azure AD Connect sync isn't instant, so
            # Exchange Online's copy still shows these for a while after a successful on-prem
            # removal. Without this, every on-prem success is immediately followed by a confusing
            # "PROTECTED" re-attempt on the exact same addresses that were just fixed.
            if ($onPremResult -and $onPremResult.Removed.Count -gt 0) {
                $onPremRemovedLower = @($onPremResult.Removed | ForEach-Object { $_.ToLowerInvariant() })
                $split.Remove = @($split.Remove | Where-Object { $onPremRemovedLower -notcontains $_.ToLowerInvariant() })
            }

            if ($split.Remove.Count -eq 0) {
                if ($onPremResult -and $onPremResult.Removed.Count -gt 0) { $ok++ } else { Write-Host "  No change : $identity"; $nochange++ }
                continue
            }

            $removeList = $split.Remove -join ', '

            if ($WhatIf) {
                Write-Host "  WhatIf    : $identity  would remove (Exchange Online): $removeList"
                _RawLog "WHATIF $identity  remove=[$removeList]"
                $ok++
            } else {
                try {
                    & $sec.SetCmd -Identity $identity -EmailAddresses $split.Keep -ErrorAction Stop
                    Write-Host "  Removed   : $identity  (Exchange Online)  [$removeList]"
                    _RawLog "OK $identity  removed=[$removeList]"
                    $ok++
                } catch {
                    $msg = $_.Exception.Message.Split([Environment]::NewLine)[0]
                    if ($msg -match "write scope|synchronized from your on-premises") {
                        Write-Host "  PROTECTED : $identity — [$removeList] exist only in Exchange Online on a DirSync'd object EXO won't let this edit through; likely needs Teams/Skype for Business Online admin tools, or will clear once directory sync ends for this domain"
                        _RawLog "PROTECTED $identity  remove=[$removeList]  $msg"
                        $protected++
                    } else {
                        Write-Host "  FAILED    : $identity — $msg"
                        _RawLog "FAIL $identity  $msg"
                        $fail++
                    }
                }
            }
        } catch {
            $msg = $_.Exception.Message.Split([Environment]::NewLine)[0]
            if ($onPremResult -and $onPremResult.Removed.Count -gt 0) {
                Write-Host "  (Removed on-prem AD; could not verify Exchange Online for $identity — $msg)"
                $ok++
            } else {
                Write-Host "  FAILED    : $identity — $msg"
                _RawLog "FAIL $identity  $msg"
                $fail++
            }
        }
    }

    Write-Host "  $($sec.Label): $ok processed  |  $nochange unchanged  |  $protected hybrid-protected  |  $fail failed"
}

if ($exoConnected) {
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
}
Write-Host '=== Done ==='
