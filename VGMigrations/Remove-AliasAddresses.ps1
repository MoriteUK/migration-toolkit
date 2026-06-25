#Requires -Version 7.0
<#
.SYNOPSIS
    Remove-AliasAddresses.ps1 — Strips alias (secondary smtp:) and IM (SIP/EUM) addresses
    from Exchange Online recipients listed in discovery CSVs.

.PARAMETER DiscoveryFolder
    Path to the folder containing discovery CSV files (02_Mailboxes, 03_DistributionGroups, etc.)

.PARAMETER Domain
    Optional: restrict removal to addresses containing this domain (e.g. "olddomain.com").
    If omitted, all secondary smtp: and SIP/EUM addresses are removed.

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

$script:SectionDefs = @(
    [pscustomobject]@{ CsvName='02_Mailboxes.csv';          Label='Mailboxes';           KeyField='UserPrincipalName';    GetCmd='Get-Mailbox';           SetCmd='Set-Mailbox'           }
    [pscustomobject]@{ CsvName='03_DistributionGroups.csv'; Label='Distribution Groups'; KeyField='PrimarySmtpAddress';   GetCmd='Get-DistributionGroup'; SetCmd='Set-DistributionGroup' }
    [pscustomobject]@{ CsvName='04_MailContacts.csv';       Label='Mail Contacts';       KeyField='ExternalEmailAddress'; GetCmd='Get-MailContact';       SetCmd='Set-MailContact'       }
    [pscustomobject]@{ CsvName='05_SharedMailboxes.csv';    Label='Shared Mailboxes';    KeyField='PrimarySmtpAddress';   GetCmd='Get-Mailbox';           SetCmd='Set-Mailbox'           }
    [pscustomobject]@{ CsvName='06_M365Groups.csv';         Label='M365 Groups';         KeyField='PrimarySmtpAddress';   GetCmd='Get-UnifiedGroup';      SetCmd='Set-UnifiedGroup'      }
)

# Returns which addresses to keep and which to remove.
# Removes: smtp: (alias), sip:, eum: — keeps SMTP: (primary) and all others.
# If FilterDomain is set, only removes addresses containing that domain.
function Split-Addresses {
    param([string[]]$Addresses, [string]$FilterDomain)

    $keep   = [System.Collections.Generic.List[string]]::new()
    $remove = [System.Collections.Generic.List[string]]::new()

    foreach ($addr in $Addresses) {
        $prefix = ($addr -split ':')[0]
        # smtp: (lowercase) = alias; sip:/eum: = IM — case-sensitive prefix check
        $isRemovable = ($prefix -ceq 'smtp') -or ($prefix -iin @('sip','eum'))

        if ($isRemovable -and (-not $FilterDomain -or $addr -like "*@$FilterDomain")) {
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

$scopeMsg = if ($Domain) { "addresses matching @$Domain" } else { 'all alias (smtp:) and IM (sip:/eum:) addresses' }
Write-Host "=== Remove Alias Addresses$(if ($WhatIf) { ' [WhatIf]' }) ==="
Write-Host "Discovery folder : $discFolder"
Write-Host "Scope            : removing $scopeMsg"
_RawLog "DiscoveryFolder=$discFolder  Domain=$Domain  WhatIf=$WhatIf"

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
            Connect-ExchangeOnline -ShowBanner:$false -CommandName $cmds -ErrorAction Stop
            $exoConnected = $true
            Write-Host 'Exchange Online connected.'
        } catch {
            Write-Host "ERROR: Failed to connect to Exchange Online: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
            exit 1
        }
    }

    $ok = 0; $fail = 0; $nochange = 0

    foreach ($r in $rows) {
        $identity = $r.($sec.KeyField)
        if (-not $identity) { continue }

        try {
            $obj     = & $sec.GetCmd -Identity $identity -ErrorAction Stop
            $current = @($obj.EmailAddresses | ForEach-Object { "$_" })
            $split   = Split-Addresses -Addresses $current -FilterDomain $Domain

            if ($split.Remove.Count -eq 0) {
                Write-Host "  No change : $identity"
                $nochange++
                continue
            }

            $removeList = $split.Remove -join ', '

            if ($WhatIf) {
                Write-Host "  WhatIf    : $identity  would remove: $removeList"
                _RawLog "WHATIF $identity  remove=[$removeList]"
                $ok++
            } else {
                & $sec.SetCmd -Identity $identity -EmailAddresses $split.Keep -ErrorAction Stop
                Write-Host "  Removed   : $identity  [$removeList]"
                _RawLog "OK $identity  removed=[$removeList]"
                $ok++
            }
        } catch {
            $msg = $_.Exception.Message.Split([Environment]::NewLine)[0]
            Write-Host "  FAILED    : $identity — $msg"
            _RawLog "FAIL $identity  $msg"
            $fail++
        }
    }

    Write-Host "  $($sec.Label): $ok processed  |  $nochange unchanged  |  $fail failed"
}

if ($exoConnected) {
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
}
Write-Host '=== Done ==='
