#Requires -Version 7.0
<#
.SYNOPSIS
    search-domain.ps1 — Discovers all references to a domain across an M365 tenant and on-prem AD,
    producing one CSV per object type plus a merged Excel workbook for migration planning.

.DESCRIPTION
    Scans a tenant in 22 sections covering identity, Exchange, SharePoint, Teams, Power Platform,
    devices, Entra policies, and license assignments. Hybrid mode also includes on-prem AD users,
    groups, and contacts.

    Exchange Online:
      - Accepted domains
      - AD Users + mailbox stats joined by UPN (Hybrid only) — full attribute dump including
        all 15 extension attributes, mailbox sizes, archive, holds, forwarding, quotas
      - Distribution and mail-enabled security groups (+ members with -IncludeMembers)
      - Mail contacts
      - Shared, Resource, Equipment, Scheduling, and Team mailboxes (with full stats)
      - M365 Groups (+ members/owners with -IncludeMembers)
      - Microsoft Teams (derived from M365 Groups)
      - Proxy addresses to remove
      - Exchange Transport Rules

    Microsoft Graph:
      - App registrations + Enterprise applications (service principals)
      - SharePoint sites — multi-strategy: getAllSites → group-association → SPO Management
        Shell (child process) → search fallback. Parallel storage enrichment.
      - OneDrives — per-user enumeration (AD ∪ Graph) with sizes, item counts, and status
      - Entra registered devices
      - Planner plans (per group; placeholder rows when permissions block details)
      - Conditional Access policies + named locations
      - Authentication policies (federation, HRD, token issuance, cross-tenant access)
      - User license assignments (per-user list + aggregate count per SKU)

    Power Platform (child process to avoid SDK assembly conflicts):
      - Power Apps + Power Automate Flows

    Hybrid (-Hybrid):
      - On-prem AD Users / Groups / Contacts via LDAP filter

    Output:
      - <domain>\Discovery\NN_<area>.csv per section
      - <domain>\1. <prefix> Discovery Objects.xlsx merged workbook
      - <domain>\_Search-M365Domain_*.log verbose log file
      - Previous runs auto-archived to <domain>\Archive\<prefix>_<timestamp>\

.PARAMETER Domain
    Domain name to search for (e.g. contoso.com). Prompted if omitted.

.PARAMETER IncludeMembers
    Includes Distribution Group members + M365 Group members/owners.

.PARAMETER Hybrid
    Enables on-prem AD scanning. Requires the ActiveDirectory PowerShell module.
    When set, section 2 (Users + mailboxes) is driven by Get-ADUser and joins AD attributes
    (description, OU, all 15 ext attrs, etc.) to mailbox stats.
    When omitted, section 2 falls back to a cloud-only mode: enumerates users via Graph by UPN
    suffix and pulls mailbox stats only — AD-specific columns in the CSV will be blank.
    Sections 19–21 (on-prem AD users/groups/contacts) only run with -Hybrid.

.PARAMETER BusinessUnitId
    Filters AD-sourced sections by ExtensionAttribute7. Cloud-only sections (DLs, M365 Groups)
    are not filtered because CustomAttribute7 is rarely populated on those objects.

.PARAMETER SkipPowerPlatform
    Bypasses the upfront Power Platform sign-in and section 16. Use for unattended/overnight
    runs (no interactive PP sign-in available) or in batch orchestration where one PP scan
    is shared across all domains. CSV is written with a placeholder row so downstream tools
    don't fail.

.NOTES
    Version    : 2.11.4
    Last edit  : 2026-05-11
    Author     : Andrew White / Claude (Anthropic)
    Repository : internal — Volaris M365 migration tooling

    Version history:
      2.11.4 The "argument 'SkuPartNumber'" error reappeared after v2.11.2's Sort-Object
             fix, but from somewhere new — the section 22 catch block was collapsing the
             failure to a single line of text with no location info. Two changes:
               1. The catch block now logs InvocationInfo.ScriptLineNumber / OffsetInLine /
                  Line + the full exception + ScriptStackTrace so any future section 22
                  failure surfaces with usable diagnostics.
               2. Pre-emptively converted all three `Get-LicenseFriendlyName -SkuPartNumber X`
                  call sites to the positional form `Get-LicenseFriendlyName X`. The named
                  form is correct PowerShell, but in contexts where the parser misreads
                  surrounding tokens, the bare `-SkuPartNumber` can be detached from the
                  function call and reported as a positional arg to a different command.
                  Positional calling sidesteps the whole class of issue.
      2.11.3 Suppressed the MSAL "Error Acquiring Token: System.NullReferenceException"
             stack trace that was leaking onto the screen when Connect-ExchangeOnline's
             WAM broker fails. Root cause: MSAL writes broker failure diagnostics
             directly to .NET's Console.Out — bypassing every PowerShell stream — so
             *>$null does not catch it. Fix: temporarily redirect [Console]::Out and
             [Console]::Error to a StringWriter for the duration of the broker attempt,
             restore them in finally. Swallowed output is optionally dumped to the log
             file at DEBUG level for diagnostics. Device-code fallback flow is
             unchanged (its prompts still appear on screen as intended).
      2.11.2 Fixed the ACTUAL root cause of the section 22 "A positional parameter
             cannot be found that accepts argument 'SkuPartNumber'" error. The
             culprit was `Sort-Object AssignedCount -Descending, SkuPartNumber` —
             PowerShell treats the comma as the array operator, leaving the bareword
             SkuPartNumber as an unbound positional arg. v2.11.1's hashtable-literal
             refactor was a separate cosmetic improvement but did not address this
             bug. Replaced with the hashtable-property sort syntax which lets each
             property declare its own direction:
                 Sort-Object -Property @{Expression='AssignedCount'; Descending=$true},
                                       @{Expression='SkuPartNumber'; Descending=$false}
      2.11.1 Refactored the 22a_LicenseCounts.csv builder to assign the friendly-name
             lookup to a variable before the hashtable literal, instead of using an
             inline `if/else` as a hashtable value. (Cosmetic; not the cause of the
             section 22 failure — see 2.11.2.)
      2.11.0 Added section 22 (User Licenses). Enumerates assigned licenses for every user
             matched in section 2 via Graph /users/{upn}/licenseDetails, in parallel
             (throttle = 10). Resolves SKU IDs to skuPartNumber + friendly product names
             via /subscribedSkus and a built-in friendly-name map. Emits two CSVs:
             22_UserLicenses.csv (one row per user with semicolon-joined license list)
             and 22a_LicenseCounts.csv (aggregate count per SKU, sorted by count desc).
             Falls back to a Graph endsWith(UPN) enumeration when section 2 produced no
             rows. Section count bumped from 21 to 22 across all Start-Section labels.
      2.10.1 Fixed StrictMode crash in archive block when archived directory contained zero
             or one file: (Get-ChildItem ...).Count now wrapped in @() for array coercion.
      2.10.0 Renamed CSV files for sections 11–21 so each filename matches its section number.
             Previously sections 11–21 were off-by-one (e.g. section 12 wrote 11_Devices.csv).
             Old archives keep their original names; remediate-domain.ps1 v2.5.0 reads the
             new names. Renamed: 10_Teams→11_Teams, 11_Devices→12_Devices, 12_ProxyAddresses→
             13_ProxyAddresses, 13_TransportRules→14_TransportRules, 14_PlannerPlans→
             15_PlannerPlans, 15_PowerPlatform→16_PowerPlatform, 16_EntraCA_Policies→
             17_EntraCA_Policies, 17_EntraAuthPolicies→18_EntraAuthPolicies,
             18_OnPrem_ADUsers→19_OnPrem_ADUsers, 19_OnPrem_ADGroups→20_OnPrem_ADGroups,
             20_OnPrem_ADContacts→21_OnPrem_ADContacts.
      2.9.3  EXO WAM connect attempt now redirects all six output streams to $null so the
             noisy MSAL stack trace doesn't appear on screen. The full exception is still
             captured by the catch block and written to the log file. User now sees only a
             clean two-line WARN "falling back to device-code" message followed by the code.
      2.9.2  Removed Disconnect-ExchangeOnline at end of run. The cmdlet calls MSAL's
             ClearAllTokensAsync() on a background thread which crashes the PowerShell
             process on hosts where the WAM RuntimeBroker is broken (same root cause as
             the connect failure handled in 2.8.1, but uncatchable because it's on a
             worker thread). Sessions are now left open and reaped on process exit.
             Bonus: keeping connections alive between batch domains avoids re-auth.
      2.9.1  Section 9 filter loop reads cached site properties via PSObject.Properties[name]
             instead of direct property access. Avoids StrictMode "property cannot be found"
             errors when Get-SPOSite returned objects with empty values that ConvertTo-Json
             omitted from the cached JSON.
      2.9.0  Section 9 filter rewritten as a simple URL/Title -like match using direct property
             access (no Get-ObjProp helper, no Graph enrichment). Drops 30k-site scan from
             ~4 hours to ~30 seconds. Storage values come straight from the SPO cache instead
             of per-site Graph calls. Trade-off: matches purely on URL and Title — sites whose
             only domain reference is in description or owner are no longer caught. For a
             complete tenant migration inventory, this matches the bias of the get-spsites.ps1
             companion that informed this redesign.
      2.8.2  Removed per-site drive-owner Graph pre-fetch in section 9. On tenants with 30k+
             sites it could run for 4+ hours making one Graph call per site lacking ownerEmail.
             The SPO Management Shell cache already populates ownerEmail from $_.Owner so this
             pre-fetch was largely redundant. Section 9 now filters and enriches purely from
             cached data, with a small accuracy trade-off for sites whose match would have
             required only the drive-owner relationship.
      2.8.1  Connect-ExchangeOnline now falls back to -Device (device-code) auth when the
             Microsoft.Identity.Client RuntimeBroker raises NullReferenceException, which
             happens on some Windows hosts where WAM isn't fully available.
      2.8.0  Added -SkipPowerPlatform switch for unattended/batch runs. Companion script
             run-multiple-domains.ps1 added for sequential overnight scanning of multiple
             domains; reuses the SPO sites cache to avoid re-auth.
      2.7.0  Section 2 now runs in cloud-only mode when -Hybrid is omitted: enumerates Graph
             users by UPN suffix and pulls full mailbox stats (size, item count, archive),
             writing 02_ADUsers.csv with AD-only columns blank. Same CSV schema in both modes
             so downstream Excel merge and remediation are unaffected.
      2.6.0  Removed -OptionalAuth switch. Upfront sign-in to Graph + EXO + SPO + Power Platform
             is now mandatory and unprompted. SPO sign-in is automatically skipped when a fresh
             tenant-sites cache (≤7 days) exists. Same account is expected for all four sign-ins.
      2.5.0  Section 9 now reads tenant-sites_<tag>.json cache produced by the companion
             export-tenant-sites.ps1 if present (≤7 days old). When the cache is fresh,
             all live tenant enumeration strategies (getAllSites / SPO Management Shell /
             search) are skipped, cutting section 9 from ~2h to ~30s on large tenants.
             The OptionalAuth prompt for SPO sign-in is also skipped when a fresh cache exists.
      2.4.0  Parallelised section 9 (Strategy 1.5, drive-owner pre-fetch, storage enrichment)
             cutting SharePoint runtime from ~2h to ~10m on tenants with thousands of sites.
             SPO child launches in pwsh.exe (PowerShell 7). Detailed diagnostic block
             in SPO child on failure.
      2.3.0  Two-stream logging — verbose log file, minimal console. Archive folder names
             now prefixed with the domain (e.g. expretio_20260507_062442).
      2.2.0  Replaced PnP.PowerShell with SPO Management Shell (no app registration needed).
      2.1.0  Added -OptionalAuth for upfront sign-ins; Power Platform child process; Excel merge.
      2.0.0  Restructure: 02_ADUsers.csv now merges AD attributes + mailbox stats. 21 sections.
      1.x    Initial 23-section discovery script.

.EXAMPLE
    .\search-domain.ps1 -Domain expretio.com -IncludeMembers -Hybrid -BusinessUnitId 30364

    Hybrid scan with all members. All four sign-ins (Graph, EXO, SPO, Power Platform) prompt
    sequentially at startup before any discovery work begins. Sign in with the same account
    each time.

.EXAMPLE
    .\search-domain.ps1 -Domain contoso.com

    Cloud-only minimal scan (no AD attributes, no member enumeration).
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)][string]$Domain,
    [switch]$IncludeMembers,
    [switch]$Hybrid,
    [string]$BusinessUnitId,
    [switch]$SkipPowerPlatform,      # When set, the upfront Power Platform scan is bypassed (useful for unattended/overnight runs and batch orchestration)
    [string]$OutputPath              # Base directory for output; overrides (Get-Location) when passed from the Electron UI
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$Script:ScriptStart = Get-Date


# ─────────────────────────────────────────────────────────────
# HYBRID / BUSINESS-UNIT HELPERS
# ─────────────────────────────────────────────────────────────
$Script:ADCacheUPN   = @{}
$Script:ADCacheMail  = @{}
$Script:ADCacheSam   = @{}
$Script:HybridReady  = $false

function Get-OnPremData {
    param([string]$Identifier)
    if (-not $Script:HybridReady -or [string]::IsNullOrWhiteSpace($Identifier)) { return $null }
    $key = $Identifier.ToLower()
    if ($Script:ADCacheUPN.ContainsKey($key))  { return $Script:ADCacheUPN[$key]  }
    if ($Script:ADCacheMail.ContainsKey($key)) { return $Script:ADCacheMail[$key] }
    if ($Script:ADCacheSam.ContainsKey($key))  { return $Script:ADCacheSam[$key]  }
    return $null
}

function Get-CustomAttr7 {
    # Read CustomAttribute7 from an EXO recipient object; falls back to AD cache.
    param($Obj, [string]$Identifier)
    try {
        if ($Obj -and $Obj.PSObject.Properties.Match('CustomAttribute7').Count -gt 0 -and $Obj.CustomAttribute7) {
            return [string]$Obj.CustomAttribute7
        }
    } catch {}
    $onPrem = Get-OnPremData -Identifier $Identifier
    if ($onPrem) { return $onPrem.ExtensionAttribute7 }
    return $null
}

function Test-MatchesBusinessUnit {
    # When -BusinessUnitId is supplied, returns $true only if the row's ExtensionAttribute7 matches it.
    # Otherwise returns $true (no filtering).
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($BusinessUnitId)) { return $true }
    return ($Value -eq $BusinessUnitId)
}

# ─────────────────────────────────────────────────────────────
# PS7 GUARD
# ─────────────────────────────────────────────────────────────
if ($PSVersionTable.PSEdition -eq 'Desktop') {
    Write-Host ''
    Write-Host '╔══════════════════════════════════════════════════════════════╗' -ForegroundColor Red
    Write-Host '║  ERROR: Requires PowerShell 7+ (pwsh.exe).                  ║' -ForegroundColor Red
    Write-Host '║  You are running Windows PowerShell 5.1 (Desktop edition).  ║' -ForegroundColor Red
    Write-Host '╚══════════════════════════════════════════════════════════════╝' -ForegroundColor Red
    exit 1
}

# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────
$Script:LogPath = $null
$Script:SectionTime = $null

function Write-Log {
    <#
    .SYNOPSIS  Dual-stream logger.
    .DESCRIPTION
      Always writes to the log file. Console behaviour by level:
        - INFO  / DEBUG : file only (silent on screen unless -Force is used)
        - WARN          : screen + file (yellow)
        - ERROR         : screen + file (red)
        - SUCCESS       : screen + file (green)
      Use Write-Status for clean stage-progress output to the screen.
    #>
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','DEBUG')][string]$Level = 'INFO',
        [switch]$Force      # Force INFO/DEBUG to also appear on screen
    )
    $ts    = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
    $entry = "[$ts] [$($Level.PadRight(7))] $Message"

    # Always to file
    if ($Script:LogPath) { try { Add-Content -Path $Script:LogPath -Value $entry -Encoding UTF8 } catch {} }

    # Console behaviour by level
    switch ($Level) {
        'WARN'    { Write-Host $entry -ForegroundColor Yellow }
        'ERROR'   { Write-Host $entry -ForegroundColor Red    }
        'SUCCESS' { Write-Host $entry -ForegroundColor Green  }
        'INFO'    { if ($Force) { Write-Host $entry -ForegroundColor Cyan } }
        'DEBUG'   { if ($Force) { Write-Host $entry -ForegroundColor DarkGray } }
    }
}

function Write-Status {
    <#
    .SYNOPSIS  Clean, minimal screen output for stage progress.
    .DESCRIPTION
      Writes a one-line status to the screen and a copy to the file (as INFO).
      No timestamps on screen for legibility; full timestamps in the file.
    #>
    param(
        [Parameter(Mandatory)][string]$Message,
        [ConsoleColor]$Color = 'Cyan'
    )
    Write-Host "  $Message" -ForegroundColor $Color
    if ($Script:LogPath) {
        $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
        try { Add-Content -Path $Script:LogPath -Value "[$ts] [STATUS ] $Message" -Encoding UTF8 } catch {}
    }
}

function Write-Detail {
    <#
    .SYNOPSIS  File-only verbose detail (DEBUG level).
    .DESCRIPTION
      Use for fine-grained trace info that should appear in the log file but
      never on the screen. Examples: Graph URIs, retry attempts, per-record
      decision points, cache state changes.
    #>
    param([Parameter(Mandatory)][string]$Message)
    Write-Log -Message $Message -Level DEBUG
}

function Start-Section {
    param([Parameter(Mandatory)][string]$Label)
    $Script:SectionTime = Get-Date
    Write-Status $Label -Color Cyan
    Write-Log "===== START: $Label =====" -Level INFO
}

function End-Section {
    $elapsed = (Get-Date) - $Script:SectionTime
    $msg = "===== END  : completed in {0:mm\:ss\.fff} =====" -f $elapsed
    Write-Log $msg -Level INFO
    # Surface a clean per-section completion line on screen if it took notable time
    if ($elapsed.TotalSeconds -ge 5) {
        Write-Status ("    completed in {0:mm\:ss}" -f $elapsed) -Color DarkGray
    }
}

function Export-SafeCsv {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter()][AllowEmptyCollection()][AllowNull()][object[]]$Data,
        [Parameter(Mandatory)][string]$Label
    )
    $Data = @($Data)
    if ($Data.Count -eq 0) {
        # Write a placeholder row so every section produces a CSV — useful for
        # downstream tools (Excel merge, AvePoint imports) that expect every file.
        $placeholder = [PSCustomObject]@{
            Status    = "No $Label found"
            Domain    = $Domain
            Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        }
        try {
            $placeholder | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Force
            Write-Log "No results for '$Label' — wrote placeholder row to $(Split-Path $Path -Leaf)" -Level WARN
        } catch {
            Write-Log "Failed to write placeholder for '$Label': $_" -Level ERROR
        }
        return
    }
    try {
        $Data | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Force
        Write-Log "Exported $($Data.Count) record(s) → $(Split-Path $Path -Leaf)" -Level SUCCESS
    }
    catch { Write-Log "Failed to export '$Label': $_" -Level ERROR }
}

function Get-CsvRowCountFast {
    param([Parameter(Mandatory)][string]$Path)
    $count = 0
    $sr = [System.IO.StreamReader]::new($Path)
    try { while ($null -ne $sr.ReadLine()) { $count++ } }
    finally { $sr.Dispose() }
    return [System.Math]::Max(0, $count - 1)
}

function Ensure-Module {
    param(
        [Parameter(Mandatory)][string]$Name,
        [switch]$Optional
    )
    try {
        if (-not (Get-Module -ListAvailable -Name $Name)) {
            if ($Optional) { Write-Log "Optional module missing: $Name" -Level WARN; return $false }
            Write-Log "Required module missing: $Name. Install: Install-Module $Name -Scope CurrentUser -Force" -Level ERROR
            exit 1
        }
        Import-Module $Name -ErrorAction Stop
        Write-Log "Loaded module: $Name" -Level INFO
        return $true
    } catch {
        if ($Optional) { Write-Log "Optional module failed to load: $Name — $($_.Exception.Message)" -Level WARN; return $false }
        Write-Log "Failed to load module: $Name — $($_.Exception.Message)" -Level ERROR
        exit 1
    }
}

function Get-ObjProp {
    param(
        [Parameter(Mandatory)][object]$Obj,
        [Parameter(Mandatory)][string]$Name
    )
    if ($null -eq $Obj) { return $null }

    if ($Obj -is [System.Collections.IDictionary]) {
        if ($Obj.Contains($Name)) { return $Obj[$Name] }
        return $null
    }

    $p = $Obj.PSObject.Properties[$Name]
    if ($p) { return $p.Value }
    return $null
}

function UrlEncode {
    param([Parameter(Mandatory)][string]$Text)
    [System.Uri]::EscapeDataString($Text)
}

function Normalize-GraphUri {
    param([Parameter(Mandatory)][string]$Uri)

    # If logs show &amp; from UI copy/paste, normalize anyway
    $u = $Uri -replace '&amp;', '&'

    # Guard: if anything accidentally uses $search, force to search
    $u = $u -replace '\?\$search=', '?search='
    $u = $u -replace '&\$search=', '&search='
    return $u
}

function Test-Cmdlet {
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    try {
        Get-Command $Name -ErrorAction Stop | Out-Null
        return $true
    } catch {
        return $false
    }
}

# ─────────────────────────────────────────────────────────────
# LICENSE SKU FRIENDLY NAME MAP
# ─────────────────────────────────────────────────────────────
# Maps Microsoft skuPartNumber values to user-friendly product names.
# Source: Microsoft "Product names and service plan identifiers for licensing".
# Not exhaustive — unknown SKUs fall back to the raw skuPartNumber.
$Script:LicenseFriendlyMap = @{
    'SPE_E5'                              = 'Microsoft 365 E5'
    'SPE_E3'                              = 'Microsoft 365 E3'
    'SPE_F1'                              = 'Microsoft 365 F1'
    'SPB'                                 = 'Microsoft 365 Business Premium'
    'O365_BUSINESS_PREMIUM'               = 'Microsoft 365 Business Standard'
    'O365_BUSINESS_ESSENTIALS'            = 'Microsoft 365 Business Basic'
    'O365_BUSINESS'                       = 'Microsoft 365 Apps for Business'
    'OFFICESUBSCRIPTION'                  = 'Microsoft 365 Apps for Enterprise'
    'M365_F1'                             = 'Microsoft 365 F1'
    'M365_F3'                             = 'Microsoft 365 F3 (Frontline)'
    'ENTERPRISEPACK'                      = 'Office 365 E3'
    'ENTERPRISEPREMIUM'                   = 'Office 365 E5'
    'ENTERPRISEPACK_GOV'                  = 'Office 365 E3 (Government)'
    'ENTERPRISEPREMIUM_GOV'               = 'Office 365 E5 (Government)'
    'STANDARDPACK'                        = 'Office 365 E1'
    'DESKLESSPACK'                        = 'Office 365 F3'
    'EXCHANGESTANDARD'                    = 'Exchange Online (Plan 1)'
    'EXCHANGEENTERPRISE'                  = 'Exchange Online (Plan 2)'
    'EXCHANGEDESKLESS'                    = 'Exchange Online Kiosk'
    'EXCHANGEARCHIVE_ADDON'               = 'Exchange Online Archiving for Exchange Online'
    'EXCHANGEARCHIVE'                     = 'Exchange Online Archiving for Exchange Server'
    'EXCHANGEESSENTIALS'                  = 'Exchange Online Essentials'
    'SHAREPOINTSTANDARD'                  = 'SharePoint Online (Plan 1)'
    'SHAREPOINTENTERPRISE'                = 'SharePoint Online (Plan 2)'
    'MCOSTANDARD'                         = 'Skype for Business Online (Plan 2)'
    'MCOPSTN1'                            = 'Microsoft 365 Domestic Calling Plan'
    'MCOPSTN2'                            = 'Microsoft 365 Domestic and International Calling Plan'
    'MCOEV'                               = 'Microsoft Teams Phone Standard'
    'MCOEV_VIRTUALUSER'                   = 'Microsoft Teams Phone Resource Account'
    'MCOMEETADV'                          = 'Microsoft 365 Audio Conferencing'
    'TEAMS_EXPLORATORY'                   = 'Microsoft Teams Exploratory'
    'TEAMS_COMMERCIAL_TRIAL'              = 'Microsoft Teams Commercial Trial'
    'PROJECTPROFESSIONAL'                 = 'Project Plan 3'
    'PROJECTPREMIUM'                      = 'Project Plan 5'
    'PROJECT_P1'                          = 'Project Plan 1'
    'PROJECTESSENTIALS'                   = 'Project Online Essentials'
    'VISIOCLIENT'                         = 'Visio Plan 2'
    'VISIOONLINE_PLAN1'                   = 'Visio Plan 1'
    'POWER_BI_PRO'                        = 'Power BI Pro'
    'POWER_BI_STANDARD'                   = 'Power BI (Free)'
    'PBI_PREMIUM_PER_USER'                = 'Power BI Premium Per User'
    'POWERAPPS_PER_USER'                  = 'Power Apps per User Plan'
    'POWERAPPS_VIRAL'                     = 'Microsoft Power Apps Plan 2 Trial'
    'FLOW_FREE'                           = 'Microsoft Power Automate Free'
    'FLOW_PER_USER'                       = 'Power Automate per User Plan'
    'EMS'                                 = 'Enterprise Mobility + Security E3'
    'EMSPREMIUM'                          = 'Enterprise Mobility + Security E5'
    'AAD_PREMIUM'                         = 'Microsoft Entra ID P1'
    'AAD_PREMIUM_P2'                      = 'Microsoft Entra ID P2'
    'AAD_BASIC'                           = 'Microsoft Entra ID Basic'
    'INTUNE_A'                            = 'Microsoft Intune Plan 1'
    'INTUNE_A_VL'                         = 'Microsoft Intune Plan 1 (VL)'
    'IDENTITY_THREAT_PROTECTION'          = 'Microsoft 365 E5 Security'
    'INFORMATION_PROTECTION_COMPLIANCE'   = 'Microsoft 365 E5 Compliance'
    'WIN_DEF_ATP'                         = 'Microsoft Defender for Endpoint'
    'ATP_ENTERPRISE'                      = 'Microsoft Defender for Office 365 (Plan 1)'
    'THREAT_INTELLIGENCE'                 = 'Microsoft Defender for Office 365 (Plan 2)'
    'DYN365_ENTERPRISE_SALES'             = 'Dynamics 365 Sales Enterprise'
    'DYN365_ENTERPRISE_CUSTOMER_SERVICE'  = 'Dynamics 365 Customer Service Enterprise'
    'DYN365_BUSCENTRAL_ESSENTIAL'         = 'Dynamics 365 Business Central Essentials'
    'STREAM'                              = 'Microsoft Stream Trial'
    'WIN10_PRO_ENT_SUB'                   = 'Windows 10/11 Enterprise E3'
    'WIN10_VDA_E5'                        = 'Windows 10/11 Enterprise E5'
    'SHAREPOINTSTORAGE'                   = 'SharePoint Online Storage'
    'RIGHTSMANAGEMENT'                    = 'Azure Information Protection (Plan 1)'
}

function Get-LicenseFriendlyName {
    param([string]$SkuPartNumber)
    if ([string]::IsNullOrWhiteSpace($SkuPartNumber)) { return $null }
    if ($Script:LicenseFriendlyMap.ContainsKey($SkuPartNumber)) {
        return $Script:LicenseFriendlyMap[$SkuPartNumber]
    }
    return $SkuPartNumber
}



# ─────────────────────────────────────────────────────────────
# SPO EXECUTABLE PICKER
# ─────────────────────────────────────────────────────────────
# All scripts target PowerShell 7 exclusively. The SPO child process also runs
# under pwsh.exe — Microsoft.Online.SharePoint.PowerShell now supports PS7.
function Resolve-SPOExecutable {
    $ps7 = $null
    try { $ps7 = (Get-Command 'pwsh.exe' -ErrorAction Stop).Source } catch {}
    if ($ps7) { return $ps7 }
    # Check the standard PS7 install location as a fallback for machines where
    # pwsh.exe is installed but not yet on PATH in this session.
    $default = Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'
    if (Test-Path $default) { return $default }
    Write-Log "pwsh.exe not found on PATH or at '$default' — SPO enumeration may fail." -Level WARN
    return 'pwsh.exe'
}

# ─────────────────────────────────────────────────────────────
# GRAPH HELPERS
# ─────────────────────────────────────────────────────────────
$GraphAvailable = $false
$GraphAdvancedHeaders = @{ 'ConsistencyLevel' = 'eventual' }

function Invoke-GraphGetAll {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [hashtable]$Headers = $null,
        [int]$MaxRetries = 3,
        [int]$RetryBaseMs = 2000
    )

    if (-not $GraphAvailable) {
        Write-Detail "Invoke-GraphGetAll: skipped (Graph unavailable). URI=$Uri"
        return @()
    }

    $Uri = Normalize-GraphUri -Uri $Uri

    if ($Uri -notmatch '\$top=' -and $Uri -notmatch '/planner/' -and $Uri -notmatch '/planner\?') {
        $sep = if ($Uri -match '\?') { '&' } else { '?' }
        $Uri = "$Uri${sep}`$top=999"
    }

    Write-Detail "Invoke-GraphGetAll: starting URI=$Uri MaxRetries=$MaxRetries"
    $callStart = Get-Date

    $results = [System.Collections.Generic.List[object]]::new()
    $nextUri = $Uri
    $page = 0

    do {
        $page++
        $attempt = 0
        $success = $false

        while (-not $success -and $attempt -le $MaxRetries) {
            $attempt++
            try {
                $pageStart = Get-Date
                if ($page -eq 1 -and $attempt -eq 1) {
                    $logUri = if ($nextUri.Length -gt 120) { $nextUri.Substring(0,117) + '...' } else { $nextUri }
                    Write-Log "Graph GET: $logUri"
                }
                Write-Detail "  page=$page attempt=$attempt URI=$nextUri"

                $resp = if ($Headers) {
                    Invoke-MgGraphRequest -Uri $nextUri -Method GET -Headers $Headers -OutputType PSObject
                } else {
                    Invoke-MgGraphRequest -Uri $nextUri -Method GET -OutputType PSObject
                }

                $vals = Get-ObjProp $resp 'value'
                $pageCount = if ($vals) { @($vals).Count } else { 0 }
                if ($vals) { $results.AddRange(@($vals)) }

                $nl = Get-ObjProp $resp '@odata.nextLink'
                $nextUri = if ($nl) { [string]$nl } else { $null }

                $success = $true
                $pageMs = [int]((Get-Date) - $pageStart).TotalMilliseconds
                Write-Detail "  page=$page returned $pageCount record(s) in ${pageMs}ms; total=$($results.Count); hasNext=$([bool]$nextUri)"

                if ($page -gt 1 -and ($page % 10 -eq 0)) {
                    Write-Log "Graph paging: page $page, $($results.Count) records"
                }
            }
            catch {
                $msg        = $_.ToString()
                $statusCode = $null
                try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}

                # Build a clean reason string: prefer the JSON error message, then status code, then first line
                $reason = $null
                try {
                    $body = $_.ErrorDetails.Message
                    if ($body) {
                        $j = $body | ConvertFrom-Json -ErrorAction Stop
                        if ($j.error.message) { $reason = $j.error.message }
                    }
                } catch {}
                if (-not $reason) {
                    if ($statusCode) { $reason = "HTTP $statusCode" }
                    else             { $reason = ($msg -split "`n")[0] }
                }

                # Capture the full exception chain to the file for diagnosis
                Write-Detail "  page=$page attempt=$attempt FAILED status=$statusCode reason=$reason"
                Write-Detail "  exception: $($_.Exception.GetType().FullName) | $($_.Exception.Message)"
                if ($_.ScriptStackTrace) {
                    Write-Detail "  stack: $($_.ScriptStackTrace -replace "`r?`n",' || ')"
                }

                # Non-retryable: 400 (bad query), 401 (auth), 403 (forbidden), 404 (not found)
                $nonRetryable = ($statusCode -in 400,401,403,404) -or
                                ($msg -match 'HTTP/\d\.\d\s+(400|401|403|404)|BadRequest|Unauthorized|Forbidden|NotFound|Request_UnsupportedQuery|Syntax error')
                if ($nonRetryable) {
                    Write-Log "Graph request failed (non-retryable): $reason" -Level WARN
                    return @($results)
                }

                if ($msg -match '429|TooManyRequests') {
                    $wait = try { [int]$_.Exception.Response.Headers['Retry-After'] } catch { 30 }
                    Write-Log "Graph throttled — waiting ${wait}s (attempt $attempt/$MaxRetries)" -Level WARN
                    Start-Sleep -Seconds $wait
                } else {
                    $waitMs = $RetryBaseMs * [System.Math]::Pow(2, ($attempt - 1))
                    Write-Log "Graph error — retrying in $([int]($waitMs/1000))s (attempt $attempt/$MaxRetries): $reason" -Level WARN
                    Start-Sleep -Milliseconds $waitMs
                }
            }
        }
    } while ($nextUri)

    $totalMs = [int]((Get-Date) - $callStart).TotalMilliseconds
    Write-Detail "Invoke-GraphGetAll: complete. pages=$page records=$($results.Count) elapsedMs=$totalMs"
    return @($results)
}

function Invoke-GraphBatchRequests {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Requests,  # <- NOT typed; avoids argument type mismatch
        [hashtable]$Headers = $null
    )

    if (-not $GraphAvailable) { return @() }

    # Normalize to array no matter what input is (List, single object, array)
    $reqArray = @($Requests)

    $allResponses = [System.Collections.Generic.List[object]]::new()

    for ($i = 0; $i -lt $reqArray.Count; $i += 20) {
        $end = [System.Math]::Min($i + 20, $reqArray.Count)
        $chunk = $reqArray[$i..($end-1)]

        $payload = @{
            requests = @(
                foreach ($r in $chunk) {
                    $rid = [string](Get-ObjProp $r 'Id')
                    $url = [string](Get-ObjProp $r 'Url')
                    @{
                        id     = $rid
                        method = 'GET'
                        url    = $url
                        headers = @{ 'ConsistencyLevel' = 'eventual' }
                    }
                }
            )
        } | ConvertTo-Json -Depth 10

        $resp = if ($Headers) {
            Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/`$batch" -Method POST -Headers $Headers -Body $payload -ContentType "application/json" -OutputType PSObject
        } else {
            Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/`$batch" -Method POST -Body $payload -ContentType "application/json" -OutputType PSObject
        }

        $rs = Get-ObjProp $resp 'responses'
        if ($rs) { $allResponses.AddRange(@($rs)) }
    }

    return @($allResponses)
}

# ─────────────────────────────────────────────────────────────
# EXCEL MERGE HELPERS
# ─────────────────────────────────────────────────────────────
function Get-SafeWorksheetName {
    param([Parameter(Mandatory)][string]$Name, [hashtable]$UsedNames)
    $safe = ($Name -replace '[\\\/\?\*\[\]\:]', ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'Sheet' }
    if ($safe.Length -gt 31) { $safe = $safe.Substring(0,31).Trim() }

    $base = $safe
    $n = 1
    while ($UsedNames.ContainsKey($safe.ToLower())) {
        $suffix = " ($n)"
        $maxBase = 31 - $suffix.Length
        $safeBase = $base.Substring(0, [System.Math]::Min($base.Length, $maxBase)).Trim()
        $safe = $safeBase + $suffix
        $n++
    }
    $UsedNames[$safe.ToLower()] = $true
    return $safe
}

function Import-CsvOrHeaders {
    param([Parameter(Mandatory)][string]$Path)

    $data = @(Import-Csv -Path $Path)
    if ($data.Count -gt 0) { return $data }

    $firstLine = Get-Content -Path $Path -TotalCount 1 -ErrorAction SilentlyContinue
    if (-not $firstLine) { return @() }

    $headers = $firstLine -split ','
    $ht = [ordered]@{}
    foreach ($h in $headers) {
        $col = $h.Trim('"')
        if (-not [string]::IsNullOrWhiteSpace($col)) { $ht[$col] = '' }
    }
    if ($ht.Count -eq 0) { return @() }
    return @([pscustomobject]$ht)
}

function Merge-CsvFolderToExcel {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$FolderPath, [Parameter(Mandatory)][string]$ExcelPath, [switch]$IncludeSummary)

    $csvFiles = Get-ChildItem -Path $FolderPath -Filter '*.csv' -File | Sort-Object Name
    if (-not $IncludeSummary) { $csvFiles = $csvFiles | Where-Object { $_.Name -ne '_Summary.csv' } }
    if (-not $csvFiles -or $csvFiles.Count -eq 0) { Write-Log "No CSV files found to merge." -Level WARN; return }

    $haveImportExcel = $false
    try {
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            Write-Log "ImportExcel not found. Attempting install (CurrentUser)..." -Level WARN
            Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
        Import-Module ImportExcel -ErrorAction Stop
        $haveImportExcel = $true
    } catch {
        Write-Log "ImportExcel unavailable: $($_.Exception.Message). Trying Excel COM fallback..." -Level WARN
    }

    if ($haveImportExcel) {
        if (Test-Path $ExcelPath) { Remove-Item $ExcelPath -Force -ErrorAction SilentlyContinue }
        $used = @{}
        $first = $true
        foreach ($csv in $csvFiles) {
            $rawName = ($csv.BaseName -replace '^\d+\w?_','')
            $sheet = Get-SafeWorksheetName -Name $rawName -UsedNames $used
            $data = Import-CsvOrHeaders -Path $csv.FullName
            if ($data.Count -eq 0) { $data = @([pscustomobject]@{ Info = "No data in $($csv.Name)" }) }

            $params = @{
                Path          = $ExcelPath
                WorksheetName = $sheet
                AutoSize      = $true
                FreezeTopRow  = $true
                BoldTopRow    = $true
                AutoFilter    = $true
            }
            if (-not $first) { $params.Append = $true }
            $data | Export-Excel @params
            Write-Log "Added worksheet '$sheet' from $($csv.Name)" -Level SUCCESS
            $first = $false
        }
        Write-Log "Excel workbook created: $ExcelPath" -Level SUCCESS
        return
    }

    # Excel COM fallback
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()

    $used = @{}
    $sheetIndex = 1
    foreach ($csv in $csvFiles) {
        $rawName = ($csv.BaseName -replace '^\d+\w?_','')
        $sheetName = Get-SafeWorksheetName -Name $rawName -UsedNames $used

        $ws = if ($sheetIndex -le $workbook.Worksheets.Count) { $workbook.Worksheets.Item($sheetIndex) } else { $workbook.Worksheets.Add() }
        $ws.Name = $sheetName

        $qt = $ws.QueryTables.Add("TEXT;$($csv.FullName)", $ws.Range("A1"))
        $qt.TextFileParseType = 1
        $qt.TextFileCommaDelimiter = $true
        $qt.Refresh($false) | Out-Null
        $qt.Delete()

        $ws.Rows.Item(1).Font.Bold = $true
        $ws.Application.ActiveWindow.SplitRow = 1
        $ws.Application.ActiveWindow.FreezePanes = $true
        $ws.UsedRange.EntireColumn.AutoFit() | Out-Null

        Write-Log "Added worksheet '$sheetName' from $($csv.Name) (COM)" -Level SUCCESS
        $sheetIndex++
    }

    if (Test-Path $ExcelPath) { Remove-Item $ExcelPath -Force -ErrorAction SilentlyContinue }
    $workbook.SaveAs($ExcelPath, 51)
    $workbook.Close($true)
    $excel.Quit()

    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Log "Excel workbook created (COM): $ExcelPath" -Level SUCCESS
}

# ─────────────────────────────────────────────────────────────
# MODULE LOAD
# ─────────────────────────────────────────────────────────────
Ensure-Module -Name 'ExchangeOnlineManagement' | Out-Null
$GraphModuleOk = Ensure-Module -Name 'Microsoft.Graph.Authentication' -Optional

# ─────────────────────────────────────────────────────────────
# DOMAIN INPUT
# ─────────────────────────────────────────────────────────────
if ([string]::IsNullOrWhiteSpace($Domain)) {
    $Domain = Read-Host -Prompt 'Enter the domain name to search for (e.g. contoso.com)'
}
$Domain = $Domain.Trim().ToLower().TrimStart('@')
if ([string]::IsNullOrWhiteSpace($Domain)) { throw "No domain supplied." }
$DomainPrefix = ($Domain -split '\.')[0]

# ─────────────────────────────────────────────────────────────
# OUTPUT FOLDER + LOG
# ─────────────────────────────────────────────────────────────
$SafeName     = $Domain -replace '[\\/:*?"<>|]', '_'
if ($OutputPath) {
    if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
    Set-Location $OutputPath
}
$OutputFolder    = Join-Path (Get-Location) $SafeName
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
$DiscoveryFolder = Join-Path $OutputFolder 'Discovery'

# ── ARCHIVE PREVIOUS SCAN OUTPUT (if any) ────────────────────
# Move existing Discovery folder, log files, and the Excel workbook
# into Archive\<timestamp>\ so the current scan starts with a clean slate.
$archiveRoot = Join-Path $OutputFolder 'Archive'

# Detect what counts as previous-scan output
$prevLogs    = @(Get-ChildItem -Path $OutputFolder -Filter '_Search-M365Domain_*.log' -File -ErrorAction SilentlyContinue)
$prevPpLogs  = @(Get-ChildItem -Path $OutputFolder -Filter '_PowerPlatform_*.log'    -File -ErrorAction SilentlyContinue)
$prevSpLogs  = @(Get-ChildItem -Path $OutputFolder -Filter '_SPOTenantSites_*.log'   -File -ErrorAction SilentlyContinue)
$prevExcel   = @(Get-ChildItem -Path $OutputFolder -Filter '1. * Discovery Objects.xlsx' -File -ErrorAction SilentlyContinue)
$prevDiscovery = (Test-Path $DiscoveryFolder) -and (@(Get-ChildItem -Path $DiscoveryFolder -File -ErrorAction SilentlyContinue).Count -gt 0)

if ($prevLogs.Count -gt 0 -or $prevExcel.Count -gt 0 -or $prevDiscovery) {
    # Determine the previous scan's date from the main log filename (preferred)
    # then fall back to the most recent CSV timestamp, then to file LastWriteTime.
    $archiveStamp = $null
    if ($prevLogs.Count -gt 0) {
        $latestLog = $prevLogs | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($latestLog.BaseName -match '_Search-M365Domain_(\d{8}_\d{6})') {
            $archiveStamp = $Matches[1]
        }
    }
    if (-not $archiveStamp -and $prevDiscovery) {
        $latestCsv = @(Get-ChildItem -Path $DiscoveryFolder -Filter '*.csv' -File -ErrorAction SilentlyContinue) |
                     Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($latestCsv) { $archiveStamp = $latestCsv.LastWriteTime.ToString('yyyyMMdd_HHmmss') }
    }
    if (-not $archiveStamp) {
        $archiveStamp = "previous_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    }

    # Prefix the archive folder name with the domain prefix for clarity when reviewing later.
    $archiveFolderName = "${DomainPrefix}_$archiveStamp"
    $archiveDir = Join-Path $archiveRoot $archiveFolderName
    # Handle the unlikely case where the archive folder already exists (e.g. re-run within same second)
    $suffix = 1
    $baseArchiveDir = $archiveDir
    while (Test-Path $archiveDir) {
        $archiveDir = "$baseArchiveDir`_$suffix"
        $suffix++
    }
    if (-not (Test-Path $archiveRoot)) { New-Item -ItemType Directory -Path $archiveRoot -Force | Out-Null }
    New-Item -ItemType Directory -Path $archiveDir -Force | Out-Null

    Write-Host ""
    Write-Host "Previous scan output detected — archiving to Archive\$archiveFolderName" -ForegroundColor Yellow

    # Move logs (main + child-process logs)
    foreach ($f in @($prevLogs + $prevPpLogs + $prevSpLogs)) {
        try { Move-Item -Path $f.FullName -Destination $archiveDir -Force } catch {
            Write-Host "  Failed to move $($f.Name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    # Move Excel workbook
    foreach ($f in $prevExcel) {
        try { Move-Item -Path $f.FullName -Destination $archiveDir -Force } catch {
            Write-Host "  Failed to move $($f.Name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    # Move the entire Discovery folder if it exists with content
    if ($prevDiscovery) {
        try {
            Move-Item -Path $DiscoveryFolder -Destination $archiveDir -Force
        } catch {
            Write-Host "  Failed to archive Discovery folder: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    $movedCount = @(Get-ChildItem -Path $archiveDir -Recurse -File -ErrorAction SilentlyContinue).Count
    Write-Host "  Archived $movedCount file(s) to: $archiveDir" -ForegroundColor Yellow
    Write-Host ""
}

# Now safe to create a fresh Discovery folder
if (-not (Test-Path $DiscoveryFolder)) { New-Item -ItemType Directory -Path $DiscoveryFolder -Force | Out-Null }
$Script:LogPath = Join-Path $OutputFolder "_Search-M365Domain_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

Write-Host ""
Write-Host '================================================' -ForegroundColor Cyan
Write-Status "M365 Domain Reference Scanner"
Write-Status "Domain         : $Domain"
Write-Status "IncludeMembers : $($IncludeMembers.IsPresent)"
Write-Status "Output         : $OutputFolder"
Write-Status "Hybrid mode    : $($Hybrid.IsPresent)"
if ($BusinessUnitId) { Write-Status "Business unit  : ExtensionAttribute7 = $BusinessUnitId" } else { Write-Status "Business unit  : (no filter)" }
Write-Host '================================================' -ForegroundColor Cyan
Write-Host ""

# ── Detailed environment block (file only) ────────────────
Write-Detail "================ SESSION START ================"
Write-Detail "Date           : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')"
Write-Detail "Script         : $PSCommandPath"
Write-Detail "PSVersion      : $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))"
Write-Detail "OS             : $($PSVersionTable.OS)"
Write-Detail "Platform       : $($PSVersionTable.Platform)"
Write-Detail "Hostname       : $env:COMPUTERNAME"
Write-Detail "RunningUser    : $env:USERNAME"
Write-Detail "WorkingDir     : $(Get-Location)"
Write-Detail "Parameters     : Domain=$Domain IncludeMembers=$($IncludeMembers.IsPresent) Hybrid=$($Hybrid.IsPresent) BusinessUnitId=$BusinessUnitId"
Write-Detail "OutputFolder   : $OutputFolder"
Write-Detail "DiscoveryFolder: $DiscoveryFolder"
Write-Detail "LogPath        : $Script:LogPath"
Write-Detail "PSModulePath   : $env:PSModulePath"
$ipAddrs = try { (Get-NetIPAddress -ErrorAction SilentlyContinue | Where-Object { $_.AddressFamily -eq 'IPv4' -and $_.IPAddress -notlike '169.*' -and $_.IPAddress -ne '127.0.0.1' } | Select-Object -ExpandProperty IPAddress) -join ', ' } catch { 'unavailable' }
Write-Detail "Local IPv4     : $ipAddrs"
Write-Detail "================================================"

# ─────────────────────────────────────────────────────────────
# CONNECT — all upfront sign-ins happen in this block
# ─────────────────────────────────────────────────────────────

# 1. Microsoft Graph (interactive sign-in)
if ($GraphModuleOk) {
    Write-Log "[1/4] Connecting to Microsoft Graph (tenant: $Domain)..."
    $graphScopes = @('Directory.Read.All','Group.Read.All','Sites.Read.All','Application.Read.All','User.Read.All','Device.Read.All','DeviceManagementManagedDevices.Read.All','Tasks.Read','Policy.Read.All')
    # Suppress WAM broker noise from [Console]::Out — same pattern as the EXO block below.
    $consoleSwallowG = New-Object System.IO.StringWriter
    $origOutG  = [Console]::Out
    $origErrG  = [Console]::Error
    [Console]::SetOut($consoleSwallowG)
    [Console]::SetError($consoleSwallowG)
    try {
        # -TenantId directs the login prompt to the SOURCE company's tenant.
        # Sign in with an account that has at least Global Reader access there.
        Connect-MgGraph -Scopes $graphScopes -TenantId $Domain -NoWelcome
        [Console]::SetOut($origOutG); [Console]::SetError($origErrG)
        $GraphAvailable = $true
        Write-Log "Connected to Microsoft Graph." -Level SUCCESS
    } catch {
        [Console]::SetOut($origOutG); [Console]::SetError($origErrG)
        $graphMsg = $_.Exception.Message
        $swallowedG = $consoleSwallowG.ToString()
        if ($swallowedG) { Write-Detail "Graph broker console output (swallowed): $($swallowedG -replace [Environment]::NewLine,' || ')" }

        if ($graphMsg -match 'window handle|WAM|broker|RuntimeBroker|InteractiveBrowser' -or $_.Exception -is [System.NullReferenceException]) {
            Write-Log "Graph WAM auth failed — falling back to device-code flow." -Level WARN
            Write-Log "  When prompted, enter the code shown at https://microsoft.com/devicelogin" -Level WARN
            try {
                Connect-MgGraph -Scopes $graphScopes -TenantId $Domain -NoWelcome -UseDeviceAuthentication
                $GraphAvailable = $true
                Write-Log "Connected to Microsoft Graph via device-code." -Level SUCCESS
            } catch {
                $GraphAvailable = $false
                Write-Log "Graph device-code connection failed: $($_.Exception.Message)" -Level ERROR
            }
        } else {
            $GraphAvailable = $false
            Write-Log "Graph connection failed — Graph sections will be skipped: $graphMsg" -Level WARN
        }
    } finally {
        if ([Console]::Out  -ne $origOutG)  { [Console]::SetOut($origOutG)   }
        if ([Console]::Error -ne $origErrG) { [Console]::SetError($origErrG) }
    }

    if ($GraphAvailable) {
        try {
            $ctx = Get-MgContext
            Write-Log "  Signed in as : $($ctx.Account)" -Level SUCCESS
            Write-Log "  Tenant ID    : $($ctx.TenantId)" -Level SUCCESS
            Write-Log "  Ensure this is the SOURCE company's tenant, not your own." -Level SUCCESS
        } catch {}
        try {
            $gc = [Microsoft.Graph.PowerShell.Authentication.GraphSession]::Instance.GraphHttpClient
            $gc.Timeout = [System.TimeSpan]::FromSeconds(900)
            Write-Log "Graph HTTP timeout set to 900 s."
        } catch { }
    }
} else {
    Write-Log "Microsoft.Graph.Authentication not available — Graph sections will be skipped." -Level WARN
}

# 2. Exchange Online (interactive sign-in)
Write-Log "[2/4] Connecting to Exchange Online..."
$exoCmds = @(
    'Get-AcceptedDomain',
    'Get-EXOMailbox','Get-Mailbox','Get-MailboxStatistics',
    'Get-DistributionGroup','Get-DistributionGroupMember',
    'Get-MailContact',
    'Get-Recipient',
    'Get-TransportRule',
    'Get-UnifiedGroup','Get-UnifiedGroupLinks'
)
# Try WAM silently first — on healthy hosts this succeeds in seconds; on broken WAM hosts
# MSAL writes a long NullReferenceException stack trace ("Error Acquiring Token: ...
# Microsoft.Identity.Client.Platforms.Features.RuntimeBroker.RuntimeBroker..ctor") DIRECTLY
# to .NET's Console — bypassing every PowerShell stream. Plain `*>$null` does not catch
# that because it only covers PS streams (success/error/warning/verbose/debug/information).
#
# To suppress the screen flood, we temporarily redirect [Console]::Out and [Console]::Error
# to a StringWriter for the duration of the broker attempt. Anything MSAL writes lands in
# the StringWriter (which we discard, or optionally log at DEBUG level for diagnostics).
# The actual thrown exception is still caught normally via try/catch.
$consoleSwallow = New-Object System.IO.StringWriter
$origStdOut = [Console]::Out
$origStdErr = [Console]::Error
[Console]::SetOut($consoleSwallow)
[Console]::SetError($consoleSwallow)
try {
    # -Organization directs the connection to the source company's Exchange Online tenant.
    Connect-ExchangeOnline -ShowBanner:$false -Organization $Domain -CommandName $exoCmds -ErrorAction Stop *>$null
    [Console]::SetOut($origStdOut)
    [Console]::SetError($origStdErr)
    Write-Log "Connected to Exchange Online (org: $Domain)." -Level SUCCESS
} catch {
    # Restore the real console BEFORE anything else writes (logging, device-code prompt, etc.)
    [Console]::SetOut($origStdOut)
    [Console]::SetError($origStdErr)
    $msg = $_.Exception.Message
    # Optional: dump the swallowed MSAL noise to the log file at DEBUG (file-only) for
    # post-mortem diagnostics. Comment out the next two lines if you don't want it at all.
    $swallowed = $consoleSwallow.ToString()
    if ($swallowed) { Write-Detail "EXO broker console output (swallowed): $($swallowed -replace [Environment]::NewLine,' || ')" }

    Write-Log "Standard Connect-ExchangeOnline failed (broker issue) — falling back to device-code flow." -Level WARN
    Write-Detail "EXO connect exception: $($_.Exception.GetType().FullName) | $msg"
    if ($_.ScriptStackTrace) { Write-Detail "EXO connect stack: $($_.ScriptStackTrace -replace [Environment]::NewLine,' || ')" }

    if ($msg -match 'Object reference not set|RuntimeBroker|broker' -or $_.Exception -is [System.NullReferenceException]) {
        Write-Log "  When prompted, enter the code shown below at https://login.microsoft.com/device" -Level WARN
        try {
            # Device-code output to the screen IS what we want here — don't suppress.
            Connect-ExchangeOnline -ShowBanner:$false -Organization $Domain -CommandName $exoCmds -Device -ErrorAction Stop
                    Write-Log "Connected to Exchange Online via device-code (org: $Domain)." -Level SUCCESS
        } catch {
            Write-Log "Device-code retry also failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" -Level ERROR
            throw
        }
    } else {
        throw
    }
} finally {
    # Belt-and-braces — guarantee the real console is restored even on unexpected control flow.
    if ([Console]::Out -ne $origStdOut) { [Console]::SetOut($origStdOut) }
    if ([Console]::Error -ne $origStdErr) { [Console]::SetError($origStdErr) }
    if ($consoleSwallow) { $consoleSwallow.Dispose() }
}

$HaveExoMailbox = [bool](Test-Cmdlet 'Get-EXOMailbox')
if ($HaveExoMailbox) { Write-Log "Using Get-EXOMailbox (REST mode)" -Level INFO }

# 3. (Optional) SharePoint Online Management Shell — tenant enumeration
$Script:SPOSitesJsonPath  = $null     # set if we run SPO upfront and it succeeds
$Script:RunPowerPlatform  = $true     # default: run inline in section 16
$Script:PowerPlatformCsv  = $null     # set if we run Power Platform upfront

if ($true) {
    Write-Host ""
    Write-Host "  ─── Upfront sign-ins (use the same account for all four) ───" -ForegroundColor Cyan
    Write-Host "  Doing all sign-ins upfront avoids interruptions later in the run." -ForegroundColor DarkGray
    Write-Host ""

    # If a tenant-sites cache file exists and is recent, we don't need SPO sign-in at all.
    $skipSpoPrompt = $false
    $cacheBaseDirCheck = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $cacheCandCheck = @(Get-ChildItem -Path $cacheBaseDirCheck -Filter 'tenant-sites.json' -File -ErrorAction SilentlyContinue)
    if ($cacheCandCheck.Count -gt 0) {
        $newest = $cacheCandCheck | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        $ageDays = ((Get-Date) - $newest.LastWriteTime).TotalDays
        if ($ageDays -le 7) {
            Write-Host "  [3/4] Tenant-sites cache found ($($newest.Name), $([math]::Round($ageDays,1))d old) — skipping SPO sign-in." -ForegroundColor Green
            $skipSpoPrompt = $true
        } else {
            Write-Host "  [3/4] Tenant-sites cache found but $([math]::Round($ageDays,1))d old (>7d) — re-running SPO enumeration." -ForegroundColor Yellow
        }
    } else {
        Write-Host "  [3/4] No tenant-sites cache found — running SPO enumeration." -ForegroundColor Cyan
    }

    if ($skipSpoPrompt) {
        $spAns = 'N'
    } else {
        $spAns = 'Y'
    }
    if ($spAns -match '^[Yy]') {
        if (-not (Get-Module -ListAvailable -Name 'Microsoft.Online.SharePoint.PowerShell' -ErrorAction SilentlyContinue)) {
            Write-Log "Microsoft.Online.SharePoint.PowerShell not installed — skipping. Install: Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser" -Level WARN
        } else {
            # ── Resolve the SPO admin URL ─────────────────────────────────────────
            # Default is the Volaris tenant admin site.
            # Override per-tenant via Settings > Customer > SharePoint Admin URL.
            $adminUrl = 'https://ourvolaris-admin.sharepoint.com'

            $sharedCfgPath = Join-Path $env:LOCALAPPDATA 'FlyMigration\shared-config.json'
            try {
                if (Test-Path $sharedCfgPath) {
                    $sharedCfg = Get-Content $sharedCfgPath -Raw -Encoding UTF8 | ConvertFrom-Json
                    $cfgVal = if ($sharedCfg.SharePointAdminUrl) { $sharedCfg.SharePointAdminUrl.Trim().TrimEnd('/') } else { $null }
                    if ($cfgVal) {
                        $adminUrl = $cfgVal
                        Write-Log "SPO admin URL overridden from shared config: $adminUrl" -Level SUCCESS
                    } else {
                        Write-Log "SPO admin URL: using default ($adminUrl)"
                    }
                }
            } catch {
                Write-Log "Could not read shared config for SPO admin URL ($($_.Exception.Message)) — using default $adminUrl" -Level WARN
            }

            if ($adminUrl) {
                Write-Log "Launching SPO Management Shell tenant enumeration upfront — admin URL: $adminUrl"
                Write-Log "An interactive sign-in window will appear. Sign in with a SharePoint admin account."

                $spoJsonPath = Join-Path $OutputFolder "_SPOTenantSites.json"
                $spoLogPath  = Join-Path $OutputFolder ("_SPOTenantSites_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".log")
                $spoScript   = Join-Path $OutputFolder "_SPOTenantSites_scan.ps1"

                # Child script — uses Microsoft's first-party SPO module (no app reg needed)
                @'
param(
    [Parameter(Mandatory)][string]$AdminUrl,
    [Parameter(Mandatory)][string]$JsonPath,
    [Parameter(Mandatory)][string]$LogPath
)
function Write-ChildLog {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$ts] [$($Level.PadRight(7))] [spo-child] $Message"
    Add-Content -Path $LogPath -Value $entry -Encoding UTF8
    Write-Host $entry
}

# ── Diagnostic environment block ──
Write-ChildLog "Environment: PSEdition=$($PSVersionTable.PSEdition) PSVersion=$($PSVersionTable.PSVersion)"
Write-ChildLog "Process    : $([System.Diagnostics.Process]::GetCurrentProcess().ProcessName) PID=$PID"
Write-ChildLog "Hostname   : $env:COMPUTERNAME  User=$env:USERNAME"
Write-ChildLog "AdminUrl   : $AdminUrl"

try {
    Import-Module 'Microsoft.Online.SharePoint.PowerShell' -DisableNameChecking -ErrorAction Stop
    $spoMod = Get-Module 'Microsoft.Online.SharePoint.PowerShell' | Select-Object -First 1
    if ($spoMod) {
        Write-ChildLog "SPO module loaded: version=$($spoMod.Version) path=$($spoMod.ModuleBase)"
    } else {
        Write-ChildLog "SPO module loaded but not visible via Get-Module" -Level WARN
    }

    Write-ChildLog "Calling Connect-SPOService -Url $AdminUrl ... (interactive sign-in window will appear)"
    $connectStart = Get-Date
    try {
        Connect-SPOService -Url $AdminUrl -ErrorAction Stop
        $connectElapsed = ((Get-Date) - $connectStart).TotalSeconds
        Write-ChildLog ("Connected to {0} in {1:F1}s" -f $AdminUrl, $connectElapsed) -Level SUCCESS
    } catch {
        $connectElapsed = ((Get-Date) - $connectStart).TotalSeconds
        Write-ChildLog ("Connect-SPOService failed after {0:F1}s" -f $connectElapsed) -Level ERROR
        Write-ChildLog "  Exception type    : $($_.Exception.GetType().FullName)" -Level ERROR
        Write-ChildLog "  Exception message : $($_.Exception.Message)" -Level ERROR
        if ($_.Exception.InnerException) {
            Write-ChildLog "  Inner type        : $($_.Exception.InnerException.GetType().FullName)" -Level ERROR
            Write-ChildLog "  Inner message     : $($_.Exception.InnerException.Message)" -Level ERROR
        }
        if ($_.Exception.Message -match 'No valid OAuth 2\.0 authentication session') {
            Write-ChildLog "DIAGNOSTIC: 'No valid OAuth 2.0 session' typically means one of:" -Level ERROR
            Write-ChildLog "  - The sign-in window was closed before completion" -Level ERROR
            Write-ChildLog "  - The signed-in account lacks the SharePoint Administrator role" -Level ERROR
            Write-ChildLog "  - Conditional Access blocked the session (location/device/MFA)" -Level ERROR
            Write-ChildLog "  - The admin URL is wrong (must be https://<tenant>-admin.sharepoint.com)" -Level ERROR
        }
        if ($_.Exception.Message -match 'AADSTS') {
            $stsCode = ([regex]::Match($_.Exception.Message,'AADSTS\d+')).Value
            Write-ChildLog "DIAGNOSTIC: Azure AD token error code: $stsCode (search Microsoft docs for that code)" -Level ERROR
        }
        if ($_.ScriptStackTrace) {
            Write-ChildLog "  Stack trace       : $($_.ScriptStackTrace -replace [Environment]::NewLine,' || ')" -Level ERROR
        }
        throw
    }

    Write-ChildLog "Calling Get-SPOSite -Limit All ..."
    $listStart = Get-Date
    $sites = @(Get-SPOSite -Limit All -ErrorAction Stop)
    $listElapsed = ((Get-Date) - $listStart).TotalSeconds
    Write-ChildLog ("Get-SPOSite returned $($sites.Count) site(s) in {0:F1}s" -f $listElapsed) -Level SUCCESS

    $projected = $sites | ForEach-Object {
        [PSCustomObject]@{
            id              = if ($_.SiteId) { $_.SiteId.ToString() } else { '' }
            displayName     = $_.Title
            webUrl          = $_.Url
            description     = ''
            createdDateTime = ''
            template        = $_.Template
            storageUsageMB  = $_.StorageUsageCurrent
            storageQuotaMB  = $_.StorageQuota
            owner           = $_.Owner
            ownerEmail      = $_.Owner
            sharingCapability = $_.SharingCapability
            lastContentModifiedDate = if ($_.LastContentModifiedDate) { $_.LastContentModifiedDate.ToString('o') } else { '' }
            status          = $_.Status
        }
    }
    $projected | ConvertTo-Json -Depth 6 -Compress | Set-Content -Path $JsonPath -Encoding UTF8
    Write-ChildLog "Wrote $($projected.Count) site(s) to $(Split-Path $JsonPath -Leaf)" -Level SUCCESS

    Disconnect-SPOService
}
catch {
    Write-ChildLog "SPO enumeration failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" -Level ERROR
    exit 1
}
'@ | Set-Content -Path $spoScript -Encoding UTF8

                try {
                    $childArgs = @(
                        '-NoProfile','-NoLogo','-ExecutionPolicy','Bypass',
                        '-File', "`"$spoScript`"",
                        '-AdminUrl', "`"$adminUrl`"",
                        '-JsonPath', "`"$spoJsonPath`"",
                        '-LogPath', "`"$spoLogPath`""
                    )
                    $spoExe = Resolve-SPOExecutable
                    Write-Log "SPO child will be launched in: $spoExe"
                    $proc = Start-Process -FilePath $spoExe -ArgumentList $childArgs -Wait -NoNewWindow -PassThru
                    if ($proc.ExitCode -eq 0 -and (Test-Path $spoJsonPath)) {
                        $Script:SPOSitesJsonPath = $spoJsonPath
                        Write-Log "SPO upfront scan complete. Section 9 will use cached results." -Level SUCCESS

                        # Also write a persistent cache in the script folder so the next discovery
                        # run skips the SPO sign-in entirely (valid for 7 days).
                        $cacheDest     = Join-Path $PSScriptRoot 'tenant-sites.json'
                        $cacheMetaDest = Join-Path $PSScriptRoot 'tenant-sites.meta.json'
                        try {
                            Copy-Item $spoJsonPath $cacheDest -Force
                            $siteCount = @(Get-Content $cacheDest -Raw | ConvertFrom-Json).Count
                            @{ ExportedAt = (Get-Date -Format 'o'); SiteCount = $siteCount; TenantTag = $Domain } |
                                ConvertTo-Json | Set-Content $cacheMetaDest -Encoding UTF8 -Force
                            Write-Log "Persistent SPO cache saved: tenant-sites.json  ($siteCount sites, valid 7 days)" -Level SUCCESS
                        } catch {
                            Write-Log "Could not save persistent SPO cache: $($_.Exception.Message)" -Level WARN
                        }
                    } else {
                        Write-Log "SPO upfront scan failed (exit $($proc.ExitCode)). Section 9 will fall back to Graph search." -Level WARN
                    }
                    if (Test-Path $spoLogPath) {
                        Get-Content $spoLogPath | ForEach-Object { Add-Content -Path $Script:LogPath -Value $_ -Encoding UTF8 }
                    }
                } catch {
                    Write-Log "Failed to launch SPO child process: $($_.Exception.Message.Split("`n")[0])" -Level ERROR
                } finally {
                    if (Test-Path $spoScript) { Remove-Item $spoScript -Force -ErrorAction SilentlyContinue }
                }
            }
        }
    } else {
        Write-Log "User declined SharePoint admin enumeration. Section 9 will use Graph only."
    }

    # 4. Power Platform — child process auth + scan (always upfront unless -SkipPowerPlatform)
    if ($SkipPowerPlatform) {
        Write-Host "  [4/4] Power Platform — skipped (-SkipPowerPlatform set)." -ForegroundColor DarkGray
        $ppAns = 'N'
        $Script:RunPowerPlatform = $false
    } else {
        Write-Host "  [4/4] Running Power Platform admin scan." -ForegroundColor Cyan
        $ppAns = 'Y'
    }
    if ($ppAns -match '^[Yy]') {
        if (-not (Get-Module -ListAvailable -Name 'Microsoft.PowerApps.Administration.PowerShell' -ErrorAction SilentlyContinue)) {
            Write-Log "Microsoft.PowerApps.Administration.PowerShell not installed — skipping." -Level WARN
        } else {
            Write-Log "Launching Power Platform scan upfront."
            Write-Log "An interactive sign-in window will appear. Sign in with a Power Platform admin account."

            $ppCsvPath = Join-Path $DiscoveryFolder '16_PowerPlatform.csv'
            $ppLogPath = Join-Path $OutputFolder ("_PowerPlatform_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".log")
            $ppScript  = Join-Path $OutputFolder "_PowerPlatform_scan.ps1"

            @'
param(
    [Parameter(Mandatory)][string]$Domain,
    [Parameter(Mandatory)][string]$DomainPrefix,
    [Parameter(Mandatory)][string]$CsvPath,
    [Parameter(Mandatory)][string]$LogPath
)

function Write-ChildLog {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$ts] [$($Level.PadRight(7))] [pp-child] $Message"
    Add-Content -Path $LogPath -Value $entry -Encoding UTF8
    Write-Host $entry
}

try {
    Import-Module 'Microsoft.PowerApps.Administration.PowerShell' -ErrorAction Stop
    Add-PowerAppsAccount -ErrorAction Stop
    Write-ChildLog "Authenticated to Power Platform." -Level SUCCESS

    $results = [System.Collections.Generic.List[object]]::new()
    Write-ChildLog "Scanning Power Apps..."
    try {
        $apps = @(Get-AdminPowerApp -ErrorAction Stop | Where-Object {
            ($_.Owner.UserPrincipalName -like "*$Domain*") -or
            ($_.Internal.displayName    -like "*$DomainPrefix*")
        })
        foreach ($a in $apps) {
            $results.Add([PSCustomObject]@{
                ObjectType   = 'PowerApp'
                DisplayName  = $a.Internal.displayName
                ObjectId     = $a.AppName
                Owner        = $a.Owner.UserPrincipalName
                Environment  = $a.EnvironmentName
                CreatedTime  = $a.Internal.createdTime
                LastModified = $a.Internal.lastModifiedTime
                Action       = 'Export and re-import; reassign owner in new tenant'
            }) | Out-Null
        }
        Write-ChildLog "Found $($apps.Count) Power App(s) matching domain."
    } catch { Write-ChildLog "Power Apps scan failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" -Level WARN }

    Write-ChildLog "Scanning Power Automate flows..."
    try {
        $envs = @(Get-AdminPowerAppEnvironment -ErrorAction Stop)
        foreach ($env in $envs) {
            try {
                $flows = @(Get-AdminFlow -EnvironmentName $env.EnvironmentName -ErrorAction Stop | Where-Object {
                    ($_.Internal.properties.creator.userPrincipalName -like "*$Domain*") -or
                    ($_.Internal.properties.displayName               -like "*$DomainPrefix*")
                })
                foreach ($f in $flows) {
                    $results.Add([PSCustomObject]@{
                        ObjectType   = 'PowerAutomateFlow'
                        DisplayName  = $f.Internal.properties.displayName
                        ObjectId     = $f.FlowName
                        Owner        = $f.Internal.properties.creator.userPrincipalName
                        Environment  = $env.DisplayName
                        CreatedTime  = $f.Internal.properties.creationTime
                        LastModified = $f.Internal.properties.lastModifiedTime
                        Action       = 'Export and re-import; re-authenticate all connections'
                    }) | Out-Null
                }
            } catch {
                Write-ChildLog "Flows in environment '$($env.DisplayName)' failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" -Level WARN
            }
        }
    } catch { Write-ChildLog "Power Automate scan failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" -Level WARN }

    if ($results.Count -gt 0) {
        $results | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8 -Force
        Write-ChildLog "Exported $($results.Count) record(s) -> $(Split-Path $CsvPath -Leaf)" -Level SUCCESS
    } else {
        ([PSCustomObject]@{
            Status    = 'No Power Platform found'
            Domain    = $Domain
            Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        }) | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8 -Force
        Write-ChildLog "No Power Platform objects matched — wrote placeholder row." -Level WARN
    }
}
catch {
    Write-ChildLog "Power Platform scan failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" -Level ERROR
    exit 1
}
'@ | Set-Content -Path $ppScript -Encoding UTF8

            try {
                $childArgs = @(
                    '-NoProfile','-NoLogo','-ExecutionPolicy','Bypass',
                    '-File', "`"$ppScript`"",
                    '-Domain', $Domain,
                    '-DomainPrefix', $DomainPrefix,
                    '-CsvPath', "`"$ppCsvPath`"",
                    '-LogPath', "`"$ppLogPath`""
                )
                $proc = Start-Process -FilePath 'pwsh.exe' -ArgumentList $childArgs -Wait -NoNewWindow -PassThru
                if ($proc.ExitCode -eq 0) {
                    $Script:PowerPlatformCsv = $ppCsvPath
                    $Script:RunPowerPlatform = $false
                    Write-Log "Power Platform upfront scan complete. Section 16 will skip its inline scan." -Level SUCCESS
                } else {
                    Write-Log "Power Platform upfront scan failed (exit $($proc.ExitCode)). Section 16 will retry inline." -Level WARN
                }
                if (Test-Path $ppLogPath) {
                    Get-Content $ppLogPath | ForEach-Object { Add-Content -Path $Script:LogPath -Value $_ -Encoding UTF8 }
                }
            } catch {
                Write-Log "Failed to launch Power Platform child: $($_.Exception.Message.Split("`n")[0])" -Level ERROR
            } finally {
                if (Test-Path $ppScript) { Remove-Item $ppScript -Force -ErrorAction SilentlyContinue }
            }
        }
    } else {
        Write-Log "User declined Power Platform scan. Section 16 will be skipped."
        $Script:RunPowerPlatform = $false
    }

    Write-Host ""
    Write-Host "  ─── All sign-ins complete. Proceeding with discovery. ───" -ForegroundColor Green
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────
# 1. ACCEPTED DOMAINS
# ─────────────────────────────────────────────────────────────
Start-Section "[1/22] Accepted Domains"
try {
    $data = @(Get-AcceptedDomain | Where-Object { $_.DomainName -like "*$Domain*" } |
        Select-Object DomainName, DomainType, Default, AuthenticationType)
    Export-SafeCsv -Path (Join-Path $DiscoveryFolder '01_AcceptedDomains.csv') -Data $data -Label 'Accepted Domains'
} catch { Write-Log "Error: $($_.Exception.Message.Split("`n")[0])" -Level ERROR }
End-Section

# ─────────────────────────────────────────────────────────────
# 2. AD USERS (full attribute dump + mailbox stats)
# ─────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────
# HYBRID PRE-SCAN: build on-prem AD attribute cache
# ─────────────────────────────────────────────────────────────
if ($Hybrid) {
    Start-Section "Hybrid: caching on-prem AD attributes"
    try {
        if (-not (Get-Module -ListAvailable -Name 'ActiveDirectory')) {
            Write-Log "ActiveDirectory module not installed — RSAT-AD-PowerShell required for Hybrid mode." -Level ERROR
            Write-Log "Install: Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'" -Level WARN
        } else {
            Import-Module ActiveDirectory -ErrorAction Stop -WarningAction SilentlyContinue
            Write-Log "ActiveDirectory module loaded."

            $ldap = "(|(userPrincipalName=*$Domain*)(mail=*$Domain*)(proxyAddresses=*$Domain*))"
            $userProps = @('UserPrincipalName','mail','proxyAddresses','extensionAttribute7','sAMAccountName','DistinguishedName','Enabled','DisplayName','Description')
            $adUsers = @(Get-ADUser -LDAPFilter $ldap -Properties $userProps -ErrorAction SilentlyContinue)

            foreach ($u in $adUsers) {
                $obj = [PSCustomObject]@{
                    Type                = 'User'
                    DisplayName         = $u.DisplayName
                    UPN                 = $u.UserPrincipalName
                    Mail                = $u.mail
                    sAMAccountName      = $u.sAMAccountName
                    ExtensionAttribute7 = $u.extensionAttribute7
                    DN                  = $u.DistinguishedName
                    Enabled             = $u.Enabled
                    Description         = $u.Description
                    ProxyAddresses      = (@($u.proxyAddresses) -join '; ')
                }
                if ($u.UserPrincipalName) { $Script:ADCacheUPN[$u.UserPrincipalName.ToLower()]  = $obj }
                if ($u.mail)              { $Script:ADCacheMail[$u.mail.ToLower()]              = $obj }
                if ($u.sAMAccountName)    { $Script:ADCacheSam[$u.sAMAccountName.ToLower()]     = $obj }
            }
            Write-Log "Cached $($adUsers.Count) on-prem user(s) referencing the domain."

            $groupProps = @('mail','proxyAddresses','extensionAttribute7','DistinguishedName','GroupCategory','GroupScope','DisplayName')
            $adGroups = @(Get-ADGroup -LDAPFilter "(|(mail=*$Domain*)(proxyAddresses=*$Domain*))" -Properties $groupProps -ErrorAction SilentlyContinue)
            foreach ($g in $adGroups) {
                if ($g.mail) {
                    $obj = [PSCustomObject]@{
                        Type                = 'Group'
                        DisplayName         = $g.DisplayName
                        Mail                = $g.mail
                        ExtensionAttribute7 = $g.extensionAttribute7
                        DN                  = $g.DistinguishedName
                        GroupCategory       = $g.GroupCategory
                        GroupScope          = $g.GroupScope
                    }
                    $Script:ADCacheMail[$g.mail.ToLower()] = $obj
                }
            }
            Write-Log "Cached $($adGroups.Count) on-prem group(s) referencing the domain."

            $Script:HybridReady = $true
        }
    } catch {
        $errMsg = $_.Exception.Message.Split("`n")[0]
        if ($errMsg -match 'Active Directory Web Services|ADWS|default server') {
            Write-Log "Hybrid AD skipped — this machine cannot reach the source domain's AD (no LAN/VPN access)." -Level WARN
            Write-Log "  Uncheck 'Hybrid' in the Discovery form unless you are on the same network as the source domain." -Level WARN
            Write-Log "  Section 2 will continue with cloud-only user enumeration via Graph." -Level WARN
        } else {
            Write-Log "Hybrid AD pre-scan skipped: $errMsg" -Level WARN
        }
        $Script:HybridReady = $false
    }
    End-Section
}

Start-Section "[2/22] AD Users (full attribute dump + mailbox stats)"
$matchedUsers = [System.Collections.Generic.List[object]]::new()
try {
    if ($Hybrid -and -not $Script:HybridReady) {
        Write-Log "Hybrid AD unavailable — falling back to cloud-only user enumeration via Graph." -Level WARN
    }
    if (-not $Hybrid -or -not $Script:HybridReady) {
        # ─── Cloud-only branch (runs when -Hybrid not set, or when AD was unreachable) ──
        # Enumerate users in Graph by UPN suffix; AD-only columns stay null in the CSV.
        if (-not $GraphAvailable) {
            Write-Log "Graph unavailable and Hybrid mode off — section 2 cannot run." -Level WARN
            Export-SafeCsv -Path (Join-Path $DiscoveryFolder '02_ADUsers.csv') -Data @() -Label 'AD Users'
        } else {
            # Get ALL member users in the connected (source) tenant.
            # Filtering by UPN suffix misses users whose UPN ends with .onmicrosoft.com
            # rather than the custom domain — a very common configuration.
            Write-Log "Cloud-only mode: enumerating all member users in the $Domain tenant..."
            $graphUserUri = "https://graph.microsoft.com/v1.0/users?`$filter=userType eq 'Member'&`$select=id,userPrincipalName,displayName,givenName,surname,jobTitle,department,companyName,officeLocation,mail,mobilePhone,businessPhones,streetAddress,city,state,postalCode,country,employeeId,accountEnabled,createdDateTime,onPremisesExtensionAttributes,onPremisesSyncEnabled,onPremisesDistinguishedName,onPremisesSamAccountName,proxyAddresses"
            $graphUsers = @(Invoke-GraphGetAll -Uri $graphUserUri)
            Write-Log "Graph returned $($graphUsers.Count) member user(s)." -Level $(if ($graphUsers.Count -eq 0) { 'WARN' } else { 'SUCCESS' })
            if ($graphUsers.Count -eq 0) {
                Write-Log "  No users found. Common causes:" -Level WARN
                Write-Log "    1. Wrong tenant — verify the signed-in account/tenant matches the source company (see above)." -Level WARN
                Write-Log "    2. Permissions — the signed-in account needs User.Read.All on the source tenant." -Level WARN
                try { Write-Log "    Connected as: $((Get-MgContext).Account)  Tenant: $((Get-MgContext).TenantId)" -Level WARN } catch {}
            }

            # Apply BU filter if requested (most cloud users won't have ext7 unless synced)
            $eligible = [System.Collections.Generic.List[object]]::new()
            foreach ($g in $graphUsers) {
                $ext = $null
                $opa = Get-ObjProp $g 'onPremisesExtensionAttributes'
                if ($opa) { $ext = Get-ObjProp $opa 'extensionAttribute7' }
                if (-not (Test-MatchesBusinessUnit $ext)) { continue }
                $eligible.Add($g) | Out-Null
            }
            Write-Log "Eligible after BU filter: $($eligible.Count) user(s)."

            # Parallel mailbox lookup (same pattern as hybrid branch)
            if ($eligible.Count -gt 0) {
                Write-Log "Querying $($eligible.Count) mailboxes in parallel (throttle = 8)..."
                $parallelStart = Get-Date

                $mbxLookups = $eligible | ForEach-Object -ThrottleLimit 8 -Parallel {
                    $upn = $_.userPrincipalName
                    if ([string]::IsNullOrWhiteSpace($upn)) {
                        return [PSCustomObject]@{ UPN = $null; Found = $false; Error = 'No UPN'; Mb = $null; Stats = $null; ArchStats = $null }
                    }
                    try {
                        # Parallel runspaces don't inherit imported modules from the parent session
                        Import-Module ExchangeOnlineManagement -ErrorAction Stop | Out-Null
                        $exoConnections = Get-ConnectionInformation -ErrorAction SilentlyContinue
                        if (-not $exoConnections -or @($exoConnections).Count -eq 0) {
                            Connect-ExchangeOnline -ShowBanner:$false -CommandName 'Get-Mailbox','Get-MailboxStatistics' -ErrorAction Stop | Out-Null
                        }
                    } catch {
                        return [PSCustomObject]@{ UPN = $upn; Found = $false; Error = "EXO connect failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])"; Mb = $null; Stats = $null; ArchStats = $null }
                    }
                    $mb = $null; $stats = $null; $archStats = $null; $err = $null
                    try {
                        $mb = Get-Mailbox -Identity $upn -ErrorAction Stop -WarningAction SilentlyContinue
                        try {
                            $stats = Get-MailboxStatistics -Identity $upn -ErrorAction Stop -WarningAction SilentlyContinue
                        } catch { $err = "Stats: $($_.Exception.Message.Split([Environment]::NewLine)[0])" }
                        $archStatus = $null
                        if ($mb -and $mb.PSObject.Properties['ArchiveStatus']) { $archStatus = $mb.PSObject.Properties['ArchiveStatus'].Value }
                        if ($archStatus -eq 'Active') {
                            try {
                                $archStats = Get-MailboxStatistics -Identity $upn -Archive -ErrorAction Stop -WarningAction SilentlyContinue
                            } catch { $err = ($err, "Archive: $($_.Exception.Message.Split([Environment]::NewLine)[0])" -ne $null) -join '; ' }
                        }
                        [PSCustomObject]@{ UPN = $upn; Found = $true; Error = $err; Mb = $mb; Stats = $stats; ArchStats = $archStats }
                    } catch {
                        [PSCustomObject]@{ UPN = $upn; Found = $false; Error = $_.Exception.Message.Split([Environment]::NewLine)[0]; Mb = $null; Stats = $null; ArchStats = $null }
                    }
                }
                $parallelElapsed = (Get-Date) - $parallelStart
                Write-Log ("Parallel mailbox lookup completed in {0:mm\:ss}." -f $parallelElapsed)

                $mbxByUpn = @{}
                foreach ($r in $mbxLookups) {
                    if ($r -and $r.UPN) { $mbxByUpn[$r.UPN.ToLower()] = $r }
                }
            } else {
                $mbxByUpn = @{}
            }

            # Build rows — Graph fields populated, AD-only fields null
            $mbxFound = 0; $mbxMissing = 0
            foreach ($g in $eligible) {
                $upn = (Get-ObjProp $g 'userPrincipalName')
                $mb = $null; $mbStats = $null; $archStats = $null; $lookupError = $null
                if ($upn) {
                    $key = $upn.ToLower()
                    if ($mbxByUpn.ContainsKey($key)) {
                        $r = $mbxByUpn[$key]
                        if ($r.Found) {
                            $mb = $r.Mb; $mbStats = $r.Stats; $archStats = $r.ArchStats
                            $mbxFound++
                            if ($r.Error) { Write-Log "Mailbox lookup partial for ${upn}: $($r.Error)" -Level WARN }
                        } else {
                            $lookupError = $r.Error
                            $mbxMissing++
                        }
                    }
                }

                # Pull from onPremisesExtensionAttributes if user is synced
                $opa = Get-ObjProp $g 'onPremisesExtensionAttributes'

                $row = [ordered]@{
                    DisplayName              = (Get-ObjProp $g 'displayName')
                    GivenName                = (Get-ObjProp $g 'givenName')
                    Surname                  = (Get-ObjProp $g 'surname')
                    Initials                 = $null
                    Description              = $null
                    Title                    = (Get-ObjProp $g 'jobTitle')
                    Department               = (Get-ObjProp $g 'department')
                    Company                  = (Get-ObjProp $g 'companyName')
                    Office                   = (Get-ObjProp $g 'officeLocation')
                    OfficePhone              = if ((Get-ObjProp $g 'businessPhones')) { (@(Get-ObjProp $g 'businessPhones') -join '; ') } else { $null }
                    MobilePhone              = (Get-ObjProp $g 'mobilePhone')
                    HomePhone                = $null
                    Fax                      = $null
                    StreetAddress            = (Get-ObjProp $g 'streetAddress')
                    City                     = (Get-ObjProp $g 'city')
                    State                    = (Get-ObjProp $g 'state')
                    PostalCode               = (Get-ObjProp $g 'postalCode')
                    Country                  = (Get-ObjProp $g 'country')
                    POBox                    = $null
                    EmailAddress             = (Get-ObjProp $g 'mail')
                    UserPrincipalName        = $upn
                    sAMAccountName           = (Get-ObjProp $g 'onPremisesSamAccountName')
                    EmployeeID               = (Get-ObjProp $g 'employeeId')
                    EmployeeNumber           = $null
                    ManagerName              = $null
                    DirectReportsCount       = $null

                    AccountEnabled           = (Get-ObjProp $g 'accountEnabled')
                    PasswordLastSet          = $null
                    PasswordNeverExpires     = $null
                    AccountExpirationDate    = $null
                    LastLogonDate            = $null
                    UserAccountControl       = $null

                    WhenCreated              = (Get-ObjProp $g 'createdDateTime')
                    WhenChanged              = $null

                    DistinguishedName        = (Get-ObjProp $g 'onPremisesDistinguishedName')
                    OU                       = $null
                    CanonicalName            = $null
                    ObjectGUID               = (Get-ObjProp $g 'id')
                    SID                      = $null

                    ProxyAddresses           = if ((Get-ObjProp $g 'proxyAddresses')) { (@(Get-ObjProp $g 'proxyAddresses') -join '; ') } else { $null }
                    TargetAddress            = $null
                    HideFromGAL              = $null
                    MsExchRecipientTypeDetails = $null
                    MsExchRemoteRecipientType  = $null

                    ExtensionAttribute1      = if ($opa) { Get-ObjProp $opa 'extensionAttribute1' }  else { $null }
                    ExtensionAttribute2      = if ($opa) { Get-ObjProp $opa 'extensionAttribute2' }  else { $null }
                    ExtensionAttribute3      = if ($opa) { Get-ObjProp $opa 'extensionAttribute3' }  else { $null }
                    ExtensionAttribute4      = if ($opa) { Get-ObjProp $opa 'extensionAttribute4' }  else { $null }
                    ExtensionAttribute5      = if ($opa) { Get-ObjProp $opa 'extensionAttribute5' }  else { $null }
                    ExtensionAttribute6      = if ($opa) { Get-ObjProp $opa 'extensionAttribute6' }  else { $null }
                    ExtensionAttribute7      = if ($opa) { Get-ObjProp $opa 'extensionAttribute7' }  else { $null }
                    ExtensionAttribute8      = if ($opa) { Get-ObjProp $opa 'extensionAttribute8' }  else { $null }
                    ExtensionAttribute9      = if ($opa) { Get-ObjProp $opa 'extensionAttribute9' }  else { $null }
                    ExtensionAttribute10     = if ($opa) { Get-ObjProp $opa 'extensionAttribute10' } else { $null }
                    ExtensionAttribute11     = if ($opa) { Get-ObjProp $opa 'extensionAttribute11' } else { $null }
                    ExtensionAttribute12     = if ($opa) { Get-ObjProp $opa 'extensionAttribute12' } else { $null }
                    ExtensionAttribute13     = if ($opa) { Get-ObjProp $opa 'extensionAttribute13' } else { $null }
                    ExtensionAttribute14     = if ($opa) { Get-ObjProp $opa 'extensionAttribute14' } else { $null }
                    ExtensionAttribute15     = if ($opa) { Get-ObjProp $opa 'extensionAttribute15' } else { $null }

                    GroupMembershipCount     = $null
                    ServicePrincipalNames    = $null

                    MailboxFound             = ($null -ne $mb)
                    RecipientType            = if ($mb) { Get-ObjProp $mb 'RecipientTypeDetails' } else { $null }
                    PrimarySmtpAddress       = if ($mb) { Get-ObjProp $mb 'PrimarySmtpAddress' } else { $null }
                    Alias                    = if ($mb) { Get-ObjProp $mb 'Alias' } else { $null }
                    Database                 = if ($mb) { Get-ObjProp $mb 'Database' } else { $null }
                    ServerName               = if ($mb) { Get-ObjProp $mb 'ServerName' } else { $null }
                    ArchiveStatus            = if ($mb) { Get-ObjProp $mb 'ArchiveStatus' } else { $null }
                    ArchiveName              = if ($mb) { (@(Get-ObjProp $mb 'ArchiveName') -join '; ') } else { $null }
                    ArchiveDatabase          = if ($mb) { Get-ObjProp $mb 'ArchiveDatabase' } else { $null }
                    LitigationHoldEnabled    = if ($mb) { Get-ObjProp $mb 'LitigationHoldEnabled' } else { $null }
                    RetentionHoldEnabled     = if ($mb) { Get-ObjProp $mb 'RetentionHoldEnabled' } else { $null }
                    RetentionPolicy          = if ($mb) { Get-ObjProp $mb 'RetentionPolicy' } else { $null }
                    ForwardingAddress        = if ($mb) { Get-ObjProp $mb 'ForwardingAddress' } else { $null }
                    ForwardingSmtpAddress    = if ($mb) { Get-ObjProp $mb 'ForwardingSmtpAddress' } else { $null }
                    DeliverToMailboxAndForward = if ($mb) { Get-ObjProp $mb 'DeliverToMailboxAndForward' } else { $null }
                    HiddenFromAddressLists   = if ($mb) { Get-ObjProp $mb 'HiddenFromAddressListsEnabled' } else { $null }
                    ProhibitSendQuota        = if ($mb) { [string](Get-ObjProp $mb 'ProhibitSendQuota') } else { $null }
                    ProhibitSendReceiveQuota = if ($mb) { [string](Get-ObjProp $mb 'ProhibitSendReceiveQuota') } else { $null }
                    IssueWarningQuota        = if ($mb) { [string](Get-ObjProp $mb 'IssueWarningQuota') } else { $null }
                    WhenMailboxCreated       = if ($mb) { Get-ObjProp $mb 'WhenMailboxCreated' } else { $null }

                    MailboxSize              = if ($mbStats) { [string]$mbStats.TotalItemSize } else { $null }
                    MailboxItemCount         = if ($mbStats) { $mbStats.ItemCount } else { $null }
                    DeletedItemSize          = if ($mbStats) { [string]$mbStats.TotalDeletedItemSize } else { $null }
                    LastLogonTime            = if ($mbStats) { $mbStats.LastLogonTime } else { $null }
                    LastUserActionTime       = if ($mbStats) { $mbStats.LastUserActionTime } else { $null }
                    ArchiveTotalItemSize     = if ($archStats) { [string]$archStats.TotalItemSize } else { $null }
                    ArchiveItemCount         = if ($archStats) { $archStats.ItemCount } else { $null }

                    LookupError              = $lookupError
                }

                $matchedUsers.Add([PSCustomObject]$row) | Out-Null
            }

            Write-Log "Cloud-only Users — total: $($matchedUsers.Count) | with mailbox: $mbxFound | no mailbox: $mbxMissing"
            Export-SafeCsv -Path (Join-Path $DiscoveryFolder '02_ADUsers.csv') -Data @($matchedUsers) -Label 'AD Users (cloud-only)'
        }
    } else {
        # Re-pull AD users matching the same domain criteria as the pre-scan, but with the FULL attribute set
        $fullProps = @(
            'DisplayName','GivenName','Surname','Initials','Description','Title','Department',
            'Company','Office','OfficePhone','MobilePhone','HomePhone','Fax',
            'StreetAddress','City','State','PostalCode','Country','POBox',
            'EmailAddress','UserPrincipalName','sAMAccountName','EmployeeID','EmployeeNumber',
            'Manager','DirectReports',
            'Enabled','LockedOut','PasswordExpired','PasswordLastSet','PasswordNeverExpires',
            'AccountExpirationDate','LastLogonDate','whenCreated','whenChanged',
            'DistinguishedName','CanonicalName','ObjectGUID','SID',
            'mail','mailNickname','proxyAddresses','targetAddress',
            'msExchHideFromAddressLists','msExchRecipientTypeDetails','msExchRemoteRecipientType',
            'extensionAttribute1','extensionAttribute2','extensionAttribute3','extensionAttribute4',
            'extensionAttribute5','extensionAttribute6','extensionAttribute7','extensionAttribute8',
            'extensionAttribute9','extensionAttribute10','extensionAttribute11','extensionAttribute12',
            'extensionAttribute13','extensionAttribute14','extensionAttribute15',
            'memberOf','servicePrincipalName',
            'userAccountControl','employeeType'
        )

        $ldap = "(|(userPrincipalName=*$Domain*)(mail=*$Domain*)(proxyAddresses=*$Domain*))"
        $adUsers = @(Get-ADUser -LDAPFilter $ldap -Properties $fullProps -ErrorAction SilentlyContinue)
        Write-Log "Get-ADUser returned $($adUsers.Count) candidate user(s)."

        # ── Stage 1: filter by BU (sequential, fast — no remote calls) ──
        $eligible = [System.Collections.Generic.List[object]]::new()
        foreach ($u in $adUsers) {
            if (-not (Test-MatchesBusinessUnit $u.extensionAttribute7)) { continue }
            $eligible.Add($u) | Out-Null
        }
        Write-Log "Filtered $($eligible.Count) of $($adUsers.Count) user(s) by business-unit criteria."

        # ── Stage 2: parallel mailbox lookups (8 concurrent runspaces) ──
        # Each runspace gets the per-user mailbox + statistics + archive stats and returns
        # a hashtable keyed by UPN. Each runspace must reconnect to EXO using cached
        # token; Connect-ExchangeOnline is idempotent.
        if ($eligible.Count -gt 0) {
            Write-Log "Querying $($eligible.Count) mailboxes in parallel (throttle = 8)..."
            $parallelStart = Get-Date

            $mbxLookups = $eligible | ForEach-Object -ThrottleLimit 8 -Parallel {
                $upn = $_.UserPrincipalName
                if ([string]::IsNullOrWhiteSpace($upn)) {
                    return [PSCustomObject]@{ UPN = $null; Found = $false; Error = 'No UPN'; Mb = $null; Stats = $null; ArchStats = $null }
                }

                # Each runspace needs its own EXO connection. ShowBanner suppressed.
                # The first one will trigger interactive auth if no cached token; subsequent ones
                # share the cached token from the first.
                try {
                    # Parallel runspaces don't inherit imported modules from the parent session
                    Import-Module ExchangeOnlineManagement -ErrorAction Stop | Out-Null
                    $exoConnections = Get-ConnectionInformation -ErrorAction SilentlyContinue
                    if (-not $exoConnections -or @($exoConnections).Count -eq 0) {
                        Connect-ExchangeOnline -ShowBanner:$false -CommandName 'Get-Mailbox','Get-MailboxStatistics' -ErrorAction Stop | Out-Null
                    }
                } catch {
                    return [PSCustomObject]@{ UPN = $upn; Found = $false; Error = "EXO connect failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])"; Mb = $null; Stats = $null; ArchStats = $null }
                }

                $mb = $null; $stats = $null; $archStats = $null; $err = $null
                try {
                    $mb = Get-Mailbox -Identity $upn -ErrorAction Stop -WarningAction SilentlyContinue
                    try {
                        $stats = Get-MailboxStatistics -Identity $upn -ErrorAction Stop -WarningAction SilentlyContinue
                    } catch { $err = "Stats: $($_.Exception.Message.Split([Environment]::NewLine)[0])" }

                    # Inline what Get-ObjProp does for ArchiveStatus check (helper not available here)
                    $archStatus = $null
                    if ($mb -and $mb.PSObject.Properties['ArchiveStatus']) { $archStatus = $mb.PSObject.Properties['ArchiveStatus'].Value }
                    if ($archStatus -eq 'Active') {
                        try {
                            $archStats = Get-MailboxStatistics -Identity $upn -Archive -ErrorAction Stop -WarningAction SilentlyContinue
                        } catch { $err = ($err, "Archive: $($_.Exception.Message.Split([Environment]::NewLine)[0])" -ne $null) -join '; ' }
                    }

                    [PSCustomObject]@{ UPN = $upn; Found = $true; Error = $err; Mb = $mb; Stats = $stats; ArchStats = $archStats }
                } catch {
                    [PSCustomObject]@{ UPN = $upn; Found = $false; Error = $_.Exception.Message.Split([Environment]::NewLine)[0]; Mb = $null; Stats = $null; ArchStats = $null }
                }
            }

            $parallelElapsed = (Get-Date) - $parallelStart
            Write-Log ("Parallel mailbox lookup completed in {0:mm\:ss}." -f $parallelElapsed)

            # Build a UPN-keyed hash for the row-build pass
            $mbxByUpn = @{}
            foreach ($r in $mbxLookups) {
                if ($r -and $r.UPN) { $mbxByUpn[$r.UPN.ToLower()] = $r }
            }
        } else {
            $mbxByUpn = @{}
        }

        # ── Stage 3: build rows (sequential — fast, all in memory) ──
        $i = 0
        $total = $eligible.Count
        $mbxFound = 0; $mbxMissing = 0

        foreach ($u in $eligible) {
            $i++
            $upn = $u.UserPrincipalName
            Write-Progress -Activity "AD Users + Mailbox stats" -Status "$i / $total : $upn" `
                -PercentComplete (($i / [System.Math]::Max($total,1)) * 100)

            # Resolve manager DN to display name
            $managerName = $null
            if ($u.Manager) {
                try { $managerName = (Get-ADUser -Identity $u.Manager -Properties DisplayName -ErrorAction Stop).DisplayName } catch {}
            }

            # Look up parallel result
            $mb = $null; $mbStats = $null; $archStats = $null; $lookupError = $null
            if (-not [string]::IsNullOrWhiteSpace($upn)) {
                $key = $upn.ToLower()
                if ($mbxByUpn.ContainsKey($key)) {
                    $r = $mbxByUpn[$key]
                    if ($r.Found) {
                        $mb = $r.Mb; $mbStats = $r.Stats; $archStats = $r.ArchStats
                        if ($r.Error) { Write-Log "Mailbox lookup partial for ${upn}: $($r.Error)" -Level WARN }
                        $mbxFound++
                    } else {
                        $lookupError = $r.Error
                        $mbxMissing++
                    }
                }
            }

            $row = [ordered]@{
                # Identity
                DisplayName              = $u.DisplayName
                GivenName                = $u.GivenName
                Surname                  = $u.Surname
                Initials                 = $u.Initials
                UserPrincipalName        = $upn
                sAMAccountName           = $u.sAMAccountName
                EmailAddress             = $u.EmailAddress
                Mail                     = $u.mail
                MailNickname             = $u.mailNickname
                EmployeeID               = $u.EmployeeID
                EmployeeNumber           = $u.EmployeeNumber
                EmployeeType             = $u.employeeType

                # Employment
                Title                    = $u.Title
                Department               = $u.Department
                Company                  = $u.Company
                ManagerDN                = $u.Manager
                Manager                  = $managerName
                DirectReportCount        = if ($u.DirectReports) { @($u.DirectReports).Count } else { 0 }

                # Contact
                Office                   = $u.Office
                OfficePhone              = $u.OfficePhone
                MobilePhone              = $u.MobilePhone
                HomePhone                = $u.HomePhone
                Fax                      = $u.Fax
                StreetAddress            = $u.StreetAddress
                City                     = $u.City
                State                    = $u.State
                PostalCode               = $u.PostalCode
                Country                  = $u.Country
                POBox                    = $u.POBox
                Description              = $u.Description

                # Account state
                Enabled                  = $u.Enabled
                LockedOut                = $u.LockedOut
                PasswordExpired          = $u.PasswordExpired
                PasswordLastSet          = $u.PasswordLastSet
                PasswordNeverExpires     = $u.PasswordNeverExpires
                AccountExpirationDate    = $u.AccountExpirationDate
                LastLogonDate            = $u.LastLogonDate
                UserAccountControl       = $u.userAccountControl

                # Lifecycle
                WhenCreated              = $u.whenCreated
                WhenChanged              = $u.whenChanged

                # Object identifiers
                DistinguishedName        = $u.DistinguishedName
                OU                       = ($u.DistinguishedName -replace '^CN=[^,]+,','')
                CanonicalName            = $u.CanonicalName
                ObjectGUID               = $u.ObjectGUID
                SID                      = $u.SID

                # AD-side Exchange fields
                ProxyAddresses           = (@($u.proxyAddresses) -join '; ')
                TargetAddress            = $u.targetAddress
                HideFromGAL              = [bool]$u.msExchHideFromAddressLists
                MsExchRecipientTypeDetails = $u.msExchRecipientTypeDetails
                MsExchRemoteRecipientType  = $u.msExchRemoteRecipientType

                # Extension attributes (1–15)
                ExtensionAttribute1      = $u.extensionAttribute1
                ExtensionAttribute2      = $u.extensionAttribute2
                ExtensionAttribute3      = $u.extensionAttribute3
                ExtensionAttribute4      = $u.extensionAttribute4
                ExtensionAttribute5      = $u.extensionAttribute5
                ExtensionAttribute6      = $u.extensionAttribute6
                ExtensionAttribute7      = $u.extensionAttribute7
                ExtensionAttribute8      = $u.extensionAttribute8
                ExtensionAttribute9      = $u.extensionAttribute9
                ExtensionAttribute10     = $u.extensionAttribute10
                ExtensionAttribute11     = $u.extensionAttribute11
                ExtensionAttribute12     = $u.extensionAttribute12
                ExtensionAttribute13     = $u.extensionAttribute13
                ExtensionAttribute14     = $u.extensionAttribute14
                ExtensionAttribute15     = $u.extensionAttribute15

                # Memberships
                GroupMembershipCount     = if ($u.memberOf) { @($u.memberOf).Count } else { 0 }
                ServicePrincipalNames    = (@($u.servicePrincipalName) -join '; ')

                # Mailbox identity (joined by UPN)
                MailboxFound             = ($null -ne $mb)
                RecipientType            = if ($mb) { Get-ObjProp $mb 'RecipientTypeDetails' } else { $null }
                PrimarySmtpAddress       = if ($mb) { Get-ObjProp $mb 'PrimarySmtpAddress' } else { $null }
                Alias                    = if ($mb) { Get-ObjProp $mb 'Alias' } else { $null }
                Database                 = if ($mb) { Get-ObjProp $mb 'Database' } else { $null }
                ServerName               = if ($mb) { Get-ObjProp $mb 'ServerName' } else { $null }
                ArchiveStatus            = if ($mb) { Get-ObjProp $mb 'ArchiveStatus' } else { $null }
                ArchiveName              = if ($mb) { (@(Get-ObjProp $mb 'ArchiveName') -join '; ') } else { $null }
                ArchiveDatabase          = if ($mb) { Get-ObjProp $mb 'ArchiveDatabase' } else { $null }
                LitigationHoldEnabled    = if ($mb) { Get-ObjProp $mb 'LitigationHoldEnabled' } else { $null }
                RetentionHoldEnabled     = if ($mb) { Get-ObjProp $mb 'RetentionHoldEnabled' } else { $null }
                RetentionPolicy          = if ($mb) { Get-ObjProp $mb 'RetentionPolicy' } else { $null }
                ForwardingAddress        = if ($mb) { Get-ObjProp $mb 'ForwardingAddress' } else { $null }
                ForwardingSmtpAddress    = if ($mb) { Get-ObjProp $mb 'ForwardingSmtpAddress' } else { $null }
                DeliverToMailboxAndForward = if ($mb) { Get-ObjProp $mb 'DeliverToMailboxAndForward' } else { $null }
                HiddenFromAddressLists   = if ($mb) { Get-ObjProp $mb 'HiddenFromAddressListsEnabled' } else { $null }
                ProhibitSendQuota        = if ($mb) { [string](Get-ObjProp $mb 'ProhibitSendQuota') } else { $null }
                ProhibitSendReceiveQuota = if ($mb) { [string](Get-ObjProp $mb 'ProhibitSendReceiveQuota') } else { $null }
                IssueWarningQuota        = if ($mb) { [string](Get-ObjProp $mb 'IssueWarningQuota') } else { $null }
                WhenMailboxCreated       = if ($mb) { Get-ObjProp $mb 'WhenMailboxCreated' } else { $null }

                # Mailbox statistics
                TotalItemSize            = if ($mbStats) { [string]$mbStats.TotalItemSize } else { $null }
                TotalItemSizeBytes       = if ($mbStats -and $mbStats.TotalItemSize) { try { $mbStats.TotalItemSize.Value.ToBytes() } catch { $null } } else { $null }
                ItemCount                = if ($mbStats) { $mbStats.ItemCount } else { $null }
                DeletedItemSize          = if ($mbStats) { [string]$mbStats.TotalDeletedItemSize } else { $null }
                LastLogonTime            = if ($mbStats) { $mbStats.LastLogonTime } else { $null }
                LastLogoffTime           = if ($mbStats) { $mbStats.LastLogoffTime } else { $null }

                # Archive statistics
                ArchiveTotalItemSize     = if ($archStats) { [string]$archStats.TotalItemSize } else { $null }
                ArchiveItemCount         = if ($archStats) { $archStats.ItemCount } else { $null }

                LookupError              = $lookupError
            }

            $matchedUsers.Add([PSCustomObject]$row) | Out-Null
        }

        Write-Progress -Activity "AD Users + Mailbox stats" -Completed
        Write-Log "AD Users — total: $($matchedUsers.Count) | with mailbox: $mbxFound | no mailbox: $mbxMissing"
        Export-SafeCsv -Path (Join-Path $DiscoveryFolder '02_ADUsers.csv') -Data @($matchedUsers) -Label 'AD Users'
    }
} catch {
    Write-Log "Error in section 2: $($_.Exception.Message.Split("`n")[0])" -Level ERROR
    if ($_.ScriptStackTrace) { Write-Log "  Stack: $($_.ScriptStackTrace -replace [Environment]::NewLine,' || ')" -Level ERROR }
} finally {
    $s2Path = Join-Path $DiscoveryFolder '02_ADUsers.csv'
    if (-not (Test-Path $s2Path)) {
        Write-Log "Section 2 CSV absent — writing collected rows ($($matchedUsers.Count))" -Level WARN
        Export-SafeCsv -Path $s2Path -Data @($matchedUsers) -Label 'AD Users'
    }
}
End-Section

# ─────────────────────────────────────────────────────────────
# 3. DISTRIBUTION GROUPS + MEMBERS
# Filters by domain only — BU filter is intentionally NOT applied because
# CustomAttribute7 is rarely populated on group objects.
# ─────────────────────────────────────────────────────────────
Start-Section "[3/22] Distribution / Mail-Enabled Security Groups"
$matchedGroups = [System.Collections.Generic.List[object]]::new()
try {
    # Server-side filter: EXO accepts EmailAddresses substring + DisplayName prefix reliably.
    $dlFilter = "EmailAddresses -like '*$Domain*' -or DisplayName -like '*$DomainPrefix*'"
    $groups = @(Get-DistributionGroup -Filter $dlFilter -ResultSize Unlimited -WarningAction SilentlyContinue)
    Write-Log "Get-DistributionGroup (filtered) returned $($groups.Count) candidate group(s)."

    # Fallback: if the filtered call yielded nothing, enumerate all DLs and filter locally.
    if ($groups.Count -eq 0) {
        Write-Log "Filter returned 0 — falling back to full DL enumeration with local matching." -Level WARN
        $allGroups = @(Get-DistributionGroup -ResultSize Unlimited -WarningAction SilentlyContinue)
        $groups = @($allGroups | Where-Object {
            ($_.PrimarySmtpAddress -like "*$Domain*") -or
            ($_.DisplayName        -like "*$DomainPrefix*") -or
            ($_.Alias              -like "*$DomainPrefix*") -or
            (@($_.EmailAddresses | Where-Object { $_ -like "*$Domain*" }).Count -gt 0)
        })
        Write-Log "Local enumeration matched $($groups.Count) of $($allGroups.Count) total group(s)."
    }

    foreach ($grp in $groups) {
        # Coerce to strings — in hybrid tenants Get-DistributionGroup returns
        # SmtpAddress / ProxyAddress objects, not plain strings.
        $primarySmtp = [string]$grp.PrimarySmtpAddress
        $displayName = [string]$grp.DisplayName
        $alias       = [string]$grp.Alias
        $emailAddrs  = @($grp.EmailAddresses | ForEach-Object { [string]$_ })

        $hits = [System.Collections.Generic.List[string]]::new()
        if ($primarySmtp -like "*$Domain*")           { $hits.Add('PrimarySmtp') }
        if ($displayName -like "*$DomainPrefix*")     { $hits.Add('DisplayName(prefix)') }
        if ($displayName -like "*$Domain*")           { $hits.Add('DisplayName(domain)') }
        if ($alias       -like "*$DomainPrefix*")     { $hits.Add('Alias') }
        $px = @($emailAddrs | Where-Object { $_ -like "*$Domain*" -or $_ -like "*$DomainPrefix*" })
        if ($px) { $hits.Add("ProxyAddress(matched on $($px.Count))") }

        if ($hits.Count -gt 0) {
            $matchedGroups.Add([PSCustomObject]@{
                DisplayName        = $displayName
                PrimarySmtpAddress = $primarySmtp
                Alias              = $alias
                GroupType          = $grp.RecipientTypeDetails
                ManagedBy          = ($grp.ManagedBy -join '; ')
                MatchedOn          = $hits -join ', '
            }) | Out-Null
        }
    }

    # Diagnostic: if EXO returned candidates but our local filter rejected them all
    if ($groups.Count -gt 0 -and $matchedGroups.Count -eq 0) {
        Write-Log "EXO returned $($groups.Count) candidate(s) but 0 passed local filtering." -Level WARN
        $sample = $groups | Select-Object -First 3
        foreach ($g in $sample) {
            Write-Log "  Sample DisplayName='$($g.DisplayName)' | Primary='$($g.PrimarySmtpAddress)' | Alias='$($g.Alias)'" -Level WARN
            $allAddr = @($g.EmailAddresses | ForEach-Object { [string]$_ })
            for ($a = 0; $a -lt $allAddr.Count; $a++) {
                Write-Log "      [$a] $($allAddr[$a])" -Level WARN
            }
        }
        Write-Log "  Searching for Domain='$Domain' / DomainPrefix='$DomainPrefix' — check if the candidates actually contain these strings." -Level WARN
    }

    Write-Log "Distribution Groups — $($matchedGroups.Count) match(es) after local filtering."
    Export-SafeCsv -Path (Join-Path $DiscoveryFolder '03_DistributionGroups.csv') -Data @($matchedGroups) -Label 'Distribution Groups'

    if ($IncludeMembers -and $matchedGroups.Count -gt 0) {
        Write-Log "Fetching members for $($matchedGroups.Count) group(s)..." -Level INFO
        $dlMembers = [System.Collections.Generic.List[object]]::new()
        $i = 0

        foreach ($g in $matchedGroups) {
            $i++
            Write-Progress -Activity "DL Members" -Status "$i / $($matchedGroups.Count): $($g.PrimarySmtpAddress)" -PercentComplete (($i/$matchedGroups.Count)*100)
            try {
                $members = @(Get-DistributionGroupMember -Identity $g.PrimarySmtpAddress -ResultSize Unlimited -WarningAction SilentlyContinue)
                if (@($members).Count -eq 0) {
                    $dlMembers.Add([PSCustomObject]@{
                        GroupDisplayName    = $g.DisplayName
                        GroupEmail          = $g.PrimarySmtpAddress
                        MemberDisplayName   = '(no members)'
                        MemberEmail         = ''
                        MemberType          = ''
                        MemberRecipientType = ''
                    }) | Out-Null
                } else {
                    foreach ($m in $members) {
                        $dlMembers.Add([PSCustomObject]@{
                            GroupDisplayName    = $g.DisplayName
                            GroupEmail          = $g.PrimarySmtpAddress
                            MemberDisplayName   = $m.DisplayName
                            MemberEmail         = $m.PrimarySmtpAddress
                            MemberType          = $m.RecipientType
                            MemberRecipientType = $m.RecipientTypeDetails
                        }) | Out-Null
                    }
                }
            } catch {
                Write-Log "DL member lookup failed for $($g.PrimarySmtpAddress): $($_.Exception.Message.Split("`n")[0])" -Level WARN
            }
        }
        Write-Progress -Activity "DL Members" -Completed
        Export-SafeCsv -Path (Join-Path $DiscoveryFolder '03b_DistributionGroup_Members.csv') -Data @($dlMembers) -Label 'DL Members'
    } elseif (-not $IncludeMembers) {
        Write-Log "Member lookup skipped — use -IncludeMembers to include." -Level WARN
    }
} catch { Write-Log "Error: $($_.Exception.Message.Split("`n")[0])" -Level ERROR }
End-Section

# ─────────────────────────────────────────────────────────────
# 4. MAIL CONTACTS
# ─────────────────────────────────────────────────────────────
Start-Section "[4/22] Mail Contacts"
try {
    $ctFilter = "ExternalEmailAddress -like '*$Domain*' -or EmailAddresses -like '*$Domain*'"
    $data = @(Get-MailContact -Filter $ctFilter -ResultSize Unlimited -WarningAction SilentlyContinue |
        ForEach-Object {
            $ext7 = Get-CustomAttr7 -Obj $_ -Identifier $_.PrimarySmtpAddress
            if (-not (Test-MatchesBusinessUnit $ext7)) { return }
            $_ | Add-Member -NotePropertyName ExtensionAttribute7 -NotePropertyValue $ext7 -Force -PassThru
        } |
        Select-Object DisplayName, ExternalEmailAddress, PrimarySmtpAddress, HiddenFromAddressListsEnabled)
    Export-SafeCsv -Path (Join-Path $DiscoveryFolder '04_MailContacts.csv') -Data $data -Label 'Mail Contacts'
} catch { Write-Log "Error: $($_.Exception.Message.Split("`n")[0])" -Level ERROR }
End-Section

# ─────────────────────────────────────────────────────────────
# 5. SHARED, RESOURCE AND TEAM MAILBOXES
# ─────────────────────────────────────────────────────────────
Start-Section "[5/22] Shared, Resource and Team Mailboxes"
try {
    $shFilter = "PrimarySmtpAddress -like '*$Domain*' -or UserPrincipalName -like '*$Domain*' -or EmailAddresses -like '*$Domain*'"
    $matchedShared = [System.Collections.Generic.List[object]]::new()

    # All non-user mailbox types: shared, resource (room/equipment), scheduling, and team mailboxes.
    $nonUserTypes = @('SharedMailbox','RoomMailbox','EquipmentMailbox','SchedulingMailbox','TeamMailbox')
    $shared = [System.Collections.Generic.List[object]]::new()

    foreach ($t in $nonUserTypes) {
        try {
            # Wrap inside @() and assign to an explicitly-typed array variable.
            # In strict mode, $batch.Count fails if EXO returned a single null; @() guarantees array.
            $rawBatch = if ($HaveExoMailbox) {
                Get-EXOMailbox -RecipientTypeDetails $t -Filter $shFilter -ResultSize Unlimited -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
            } else {
                Get-Mailbox -RecipientTypeDetails $t -Filter $shFilter -ResultSize Unlimited -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
            }
            $batch = @()
            if ($null -ne $rawBatch) { $batch = @($rawBatch) }

            if ($batch.Count -gt 0) {
                Write-Log "  $t : $($batch.Count) match(es)"
                $shared.AddRange($batch)
            }
        } catch {
            Write-Log "  $t lookup failed: $($_.Exception.Message.Split("`n")[0])" -Level WARN
        }
    }

    foreach ($mb in $shared) {
        $shSize = $null; $shBytes = $null; $shItems = $null
        try {
            $st = Get-MailboxStatistics -Identity $mb.UserPrincipalName -ErrorAction Stop -WarningAction SilentlyContinue
            if ($st) {
                $shSize  = [string]$st.TotalItemSize
                $shItems = $st.ItemCount
                if ($st.TotalItemSize) { try { $shBytes = $st.TotalItemSize.Value.ToBytes() } catch {} }
            }
        } catch {
            Write-Log "Mailbox stats failed for $($mb.UserPrincipalName): $($_.Exception.Message.Split("`n")[0])" -Level WARN
        }

        $sharedOnPrem = Get-OnPremData -Identifier $mb.UserPrincipalName

        # Pull full mailbox details (was previously only in section 23)
        $shArchStats = $null
        try {
            if ((Get-ObjProp $mb 'ArchiveStatus') -eq 'Active') {
                $shArchStats = Get-MailboxStatistics -Identity $mb.UserPrincipalName -Archive -ErrorAction Stop -WarningAction SilentlyContinue
            }
        } catch {
            Write-Log "Archive stats failed for $($mb.UserPrincipalName): $($_.Exception.Message.Split("`n")[0])" -Level WARN
        }

        $matchedShared.Add([PSCustomObject]@{
            DisplayName              = Get-ObjProp $mb 'DisplayName'
            MailboxType              = Get-ObjProp $mb 'RecipientTypeDetails'
            PrimarySmtpAddress       = Get-ObjProp $mb 'PrimarySmtpAddress'
            UserPrincipalName        = Get-ObjProp $mb 'UserPrincipalName'
            Alias                    = Get-ObjProp $mb 'Alias'
            ArchiveStatus            = Get-ObjProp $mb 'ArchiveStatus'
            LitigationHoldEnabled    = Get-ObjProp $mb 'LitigationHoldEnabled'
            RetentionPolicy          = Get-ObjProp $mb 'RetentionPolicy'
            ForwardingAddress        = Get-ObjProp $mb 'ForwardingAddress'
            ForwardingSmtpAddress    = Get-ObjProp $mb 'ForwardingSmtpAddress'
            HiddenFromAddressLists   = Get-ObjProp $mb 'HiddenFromAddressListsEnabled'
            ProhibitSendQuota        = [string](Get-ObjProp $mb 'ProhibitSendQuota')
            ProhibitSendReceiveQuota = [string](Get-ObjProp $mb 'ProhibitSendReceiveQuota')
            IssueWarningQuota        = [string](Get-ObjProp $mb 'IssueWarningQuota')
            WhenMailboxCreated       = Get-ObjProp $mb 'WhenMailboxCreated'
            TotalItemSize            = $shSize
            TotalItemSizeBytes       = $shBytes
            ItemCount                = $shItems
            ArchiveTotalItemSize     = if ($shArchStats) { [string]$shArchStats.TotalItemSize } else { $null }
            ArchiveItemCount         = if ($shArchStats) { $shArchStats.ItemCount } else { $null }
            Description              = if ($sharedOnPrem) { $sharedOnPrem.Description } else { $null }
        }) | Out-Null
    }
    Export-SafeCsv -Path (Join-Path $DiscoveryFolder '05_SharedMailboxes.csv') -Data @($matchedShared) -Label 'Shared / Resource / Team Mailboxes'
} catch { Write-Log "Error: $($_.Exception.Message.Split("`n")[0])" -Level ERROR }
End-Section

# ─────────────────────────────────────────────────────────────
# 6. M365 GROUPS + MEMBERS
# ─────────────────────────────────────────────────────────────
Start-Section "[6/22] M365 Groups"
$matchedM365 = [System.Collections.Generic.List[object]]::new()
try {
    $ugFilter = "EmailAddresses -like '*$Domain*' -or PrimarySmtpAddress -like '*$Domain*' -or DisplayName -like '*$DomainPrefix*' -or Alias -like '*$DomainPrefix*'"
    $groups = @(Get-UnifiedGroup -Filter $ugFilter -ResultSize Unlimited -WarningAction SilentlyContinue)
    Write-Log "Get-UnifiedGroup returned $($groups.Count) candidate group(s) before local filtering."

    foreach ($grp in $groups) {
        $hits = [System.Collections.Generic.List[string]]::new()
        if ($grp.PrimarySmtpAddress -like "*$Domain*")       { $hits.Add('PrimarySmtp') }
        if ($grp.DisplayName        -like "*$DomainPrefix*") { $hits.Add('DisplayName') }
        if ($grp.Alias              -like "*$DomainPrefix*") { $hits.Add('Alias') }
        if ($grp.SharePointSiteUrl  -like "*$DomainPrefix*") { $hits.Add('SharePointUrl') }

        $px = @($grp.EmailAddresses | Where-Object { $_ -like "*$Domain*" })
        if ($px) { $hits.Add("ProxyAddress($($px -join ';'))") }

        if ($hits.Count -gt 0) {
            $matchedM365.Add([PSCustomObject]@{
                DisplayName        = $grp.DisplayName
                Alias              = $grp.Alias
                PrimarySmtpAddress = $grp.PrimarySmtpAddress
                SharePointSiteUrl  = $grp.SharePointSiteUrl
                AccessType         = $grp.AccessType
                IsTeam             = ($grp.ResourceProvisioningOptions -contains 'Team')
                MatchedOn          = $hits -join ', '
            }) | Out-Null
        }
    }

    Export-SafeCsv -Path (Join-Path $DiscoveryFolder '06_M365Groups.csv') -Data @($matchedM365) -Label 'M365 Groups'

    if ($IncludeMembers -and $matchedM365.Count -gt 0) {
        Write-Log "Fetching members for $($matchedM365.Count) M365 group(s)..." -Level INFO
        $m365Members = [System.Collections.Generic.List[object]]::new()
        $i = 0

        foreach ($g in $matchedM365) {
            $i++
            Write-Progress -Activity "M365 Group Members" -Status "$i / $($matchedM365.Count): $($g.PrimarySmtpAddress)" -PercentComplete (($i/$matchedM365.Count)*100)
            try {
                $members = @(Get-UnifiedGroupLinks -Identity $g.PrimarySmtpAddress -LinkType Members -ResultSize Unlimited -WarningAction SilentlyContinue)
                $owners  = @(Get-UnifiedGroupLinks -Identity $g.PrimarySmtpAddress -LinkType Owners  -ResultSize Unlimited -WarningAction SilentlyContinue)

                $ownerEmails = @($owners | ForEach-Object { $_.PrimarySmtpAddress })
                $all = @(@($members) + @($owners) | Sort-Object PrimarySmtpAddress -Unique)

                if (@($all).Count -eq 0) {
                    $m365Members.Add([PSCustomObject]@{
                        GroupDisplayName  = $g.DisplayName
                        GroupEmail        = $g.PrimarySmtpAddress
                        IsTeam            = $g.IsTeam
                        MemberDisplayName = '(no members)'
                        MemberEmail       = ''
                        MemberUPN         = ''
                        MemberRole        = ''
                        MemberType        = ''
                    }) | Out-Null
                } else {
                    foreach ($m in $all) {
                        $m365Members.Add([PSCustomObject]@{
                            GroupDisplayName  = $g.DisplayName
                            GroupEmail        = $g.PrimarySmtpAddress
                            IsTeam            = $g.IsTeam
                            MemberDisplayName = $m.DisplayName
                            MemberEmail       = $m.PrimarySmtpAddress
                            MemberUPN         = $m.WindowsLiveID
                            MemberRole        = if ($ownerEmails -contains $m.PrimarySmtpAddress) { 'Owner' } else { 'Member' }
                            MemberType        = $m.RecipientTypeDetails
                        }) | Out-Null
                    }
                }
            } catch {
                Write-Log "M365 member lookup failed for $($g.PrimarySmtpAddress): $($_.Exception.Message.Split("`n")[0])" -Level WARN
            }
        }
        Write-Progress -Activity "M365 Group Members" -Completed
        Export-SafeCsv -Path (Join-Path $DiscoveryFolder '06b_M365Group_Members.csv') -Data @($m365Members) -Label 'M365 Group Members'
    } elseif (-not $IncludeMembers) {
        Write-Log "Member lookup skipped — use -IncludeMembers to include." -Level WARN
    }
} catch { Write-Log "Error: $($_.Exception.Message.Split("`n")[0])" -Level ERROR }
End-Section

# ─────────────────────────────────────────────────────────────
# 7. APPS (Graph)
# ─────────────────────────────────────────────────────────────
Start-Section "[7/22] App Registrations (Graph)"
try {
    if (-not $GraphAvailable) { Write-Log "Graph unavailable — skipping App Registrations." -Level WARN }
    else {
        $apps = Invoke-GraphGetAll -Uri "https://graph.microsoft.com/v1.0/applications?`$select=id,displayName,appId,identifierUris,web,publicClient,spa"
        $matchedApps = [System.Collections.Generic.List[object]]::new()

        foreach ($app in $apps) {
            $hits = [System.Collections.Generic.List[string]]::new()
            $dn = Get-ObjProp $app 'displayName'
            $appId = Get-ObjProp $app 'appId'
            $objId = Get-ObjProp $app 'id'
            $idUris = @(Get-ObjProp $app 'identifierUris')

            if ($dn -and $dn -like "*$DomainPrefix*") { $hits.Add('DisplayName') }
            $iHits = @($idUris | Where-Object { $_ -like "*$Domain*" })
            if ($iHits) { $hits.Add("IdentifierUri($($iHits -join ';'))") }

            $ru = @()
            $web = Get-ObjProp $app 'web'
            $pc  = Get-ObjProp $app 'publicClient'
            $spa = Get-ObjProp $app 'spa'
            if ($web) { $ru += @((Get-ObjProp $web 'redirectUris')) }
            if ($pc)  { $ru += @((Get-ObjProp $pc  'redirectUris')) }
            if ($spa) { $ru += @((Get-ObjProp $spa 'redirectUris')) }

            $rHits = @($ru | Where-Object { $_ -like "*$Domain*" })
            if ($rHits) { $hits.Add("ReplyUri($($rHits -join ';'))") }

            if ($hits.Count -gt 0) {
                $matchedApps.Add([PSCustomObject]@{
                    DisplayName    = $dn
                    AppId          = $appId
                    ObjectId       = $objId
                    IdentifierUris = ($idUris -join '; ')
                    MatchedOn      = ($hits -join ', ')
                }) | Out-Null
            }
        }

        Export-SafeCsv -Path (Join-Path $DiscoveryFolder '07_AppRegistrations.csv') -Data @($matchedApps) -Label 'App Registrations'
    }
} catch { Write-Log "Error: $($_.Exception.Message.Split("`n")[0])" -Level ERROR }
End-Section

# ─────────────────────────────────────────────────────────────
# 8. SERVICE PRINCIPALS (Graph)
# ─────────────────────────────────────────────────────────────
Start-Section "[8/22] Enterprise Applications / Service Principals (Graph)"
try {
    if (-not $GraphAvailable) { Write-Log "Graph unavailable — skipping Enterprise Applications." -Level WARN }
    else {
        $sps = Invoke-GraphGetAll -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$select=id,displayName,appId,replyUrls,servicePrincipalNames,loginUrl,homepage"
        $matchedSPs = [System.Collections.Generic.List[object]]::new()

        foreach ($sp in $sps) {
            $hits = [System.Collections.Generic.List[string]]::new()
            $dn = Get-ObjProp $sp 'displayName'
            $appId = Get-ObjProp $sp 'appId'
            $objId = Get-ObjProp $sp 'id'
            $homepage = Get-ObjProp $sp 'homepage'
            $loginUrl = Get-ObjProp $sp 'loginUrl'
            $replyUrls = @(Get-ObjProp $sp 'replyUrls')
            $spns = @(Get-ObjProp $sp 'servicePrincipalNames')

            if ($dn -and $dn -like "*$DomainPrefix*") { $hits.Add('DisplayName') }
            if ($homepage -and $homepage -like "*$Domain*") { $hits.Add('Homepage') }
            if ($loginUrl -and $loginUrl -like "*$Domain*") { $hits.Add('LoginUrl') }

            $rHits = @($replyUrls | Where-Object { $_ -like "*$Domain*" })
            if ($rHits) { $hits.Add("ReplyUrl($($rHits -join ';'))") }
            $sHits = @($spns | Where-Object { $_ -like "*$Domain*" })
            if ($sHits) { $hits.Add("SPN($($sHits -join ';'))") }

            if ($hits.Count -gt 0) {
                $matchedSPs.Add([PSCustomObject]@{
                    DisplayName = $dn
                    AppId       = $appId
                    ObjectId    = $objId
                    Homepage    = $homepage
                    MatchedOn   = ($hits -join ', ')
                }) | Out-Null
            }
        }

        Export-SafeCsv -Path (Join-Path $DiscoveryFolder '08_EnterpriseApps.csv') -Data @($matchedSPs) -Label 'Enterprise Applications'
    }
} catch { Write-Log "Error: $($_.Exception.Message.Split("`n")[0])" -Level ERROR }
End-Section

# ─────────────────────────────────────────────────────────────
# 9. SHAREPOINT SITES (Graph) — UPDATED to avoid dot queries
# ─────────────────────────────────────────────────────────────
Start-Section "[9/22] SharePoint Sites (Graph)"
try {
    if (-not $GraphAvailable) {
        Write-Log "Graph unavailable — skipping SharePoint Sites." -Level WARN
    } else {
        $select  = "`$select=id,displayName,webUrl,description,createdDateTime"
        $allSitesList = [System.Collections.Generic.List[object]]::new()
        $tenantWide = $false

        # ── Strategy 0 (fastest): use a pre-exported tenant sites cache, if present ──
        # The companion script export-tenant-sites.ps1 produces tenant-sites_<tag>.json
        # next to itself. When found and recent, we load it instead of calling SPO mid-run.
        $cacheBaseDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
        $cacheCandidates = @(Get-ChildItem -Path $cacheBaseDir -Filter 'tenant-sites.json' -File -ErrorAction SilentlyContinue)
        if ($cacheCandidates.Count -gt 0) {
            # Pick the most recent .json (in case multiple tenants are cached)
            $cacheFile = $cacheCandidates | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            $metaFile  = Join-Path $cacheBaseDir ($cacheFile.BaseName + '.meta.json')
            $exportedAt = $null
            $siteCount  = $null
            $cacheTenantTag = $null
            if (Test-Path $metaFile) {
                try {
                    $m = Get-Content $metaFile -Raw | ConvertFrom-Json
                    if ($m.ExportedAt) { try { $exportedAt = [datetime]$m.ExportedAt } catch {} }
                    if ($m.SiteCount)  { $siteCount  = [int]$m.SiteCount }
                    if ($m.TenantTag)  { $cacheTenantTag = $m.TenantTag }
                } catch {}
            }
            if (-not $exportedAt) { $exportedAt = $cacheFile.LastWriteTime }

            $age = (Get-Date) - $exportedAt
            if ($age.TotalDays -gt 7) {
                Write-Log ("Tenant-sites cache found but {0:N1} days old (threshold 7d): $($cacheFile.Name)" -f $age.TotalDays) -Level WARN
                Write-Log "  Cache will still be used. Delete tenant-sites.json to force a fresh SPO scan." -Level WARN
            } else {
                Write-Log ("Using tenant-sites cache: $($cacheFile.Name) (exported {0:N1}d ago, $siteCount site(s), tenant=$cacheTenantTag)" -f $age.TotalDays)
            }

            try {
                $cached = @(Get-Content $cacheFile.FullName -Raw -Encoding UTF8 | ConvertFrom-Json)
                if ($cached.Count -gt 0) {
                    $allSitesList.AddRange($cached)
                    $tenantWide = $true
                    Write-Log "Loaded $($cached.Count) site(s) from cache — skipping all live tenant enumeration strategies." -Level SUCCESS
                } else {
                    Write-Log "Cache file is empty — will fall through to live enumeration." -Level WARN
                }
            } catch {
                Write-Log "Failed to read cache file: $($_.Exception.Message.Split("`n")[0]). Falling through to live enumeration." -Level WARN
            }
        } else {
            Write-Detail "No tenant-sites_*.json cache file found in $cacheBaseDir — using live enumeration."
        }

        # ── Strategy 1 (primary): enumerate every site in the tenant ──
        # /sites/getAllSites returns every site including ones the search index has missed.
        # Requires Sites.Read.All. Paginated.
        if (-not $tenantWide) {
            try {
                Write-Log "Enumerating every site via /sites/getAllSites (most reliable)..."
                $all = Invoke-GraphGetAll -Uri "https://graph.microsoft.com/v1.0/sites/getAllSites?$select"
                if ($all -and $all.Count -gt 0) {
                    $allSitesList.AddRange(@($all))
                    $tenantWide = $true
                    Write-Log "Retrieved $($all.Count) site(s) tenant-wide."
                } else {
                    Write-Log "/sites/getAllSites returned 0 — falling back to search." -Level WARN
                }
            } catch {
                Write-Log "/sites/getAllSites failed (will fall back to search): $($_.Exception.Message.Split("`n")[0])" -Level WARN
            }
        }

        # ── Strategy 1.5: get sites linked to M365 groups already discovered ──
        # Group-backed sites are accessible to any group member, so this works even
        # when /getAllSites returns 403. Catches every Team/M365 group SharePoint site.
        # Build a script-scoped Graph ID cache so section 15 (Planner) doesn't re-resolve.
        if (-not (Get-Variable -Scope Script -Name 'M365GroupGraphIdCache' -ErrorAction SilentlyContinue)) {
            $Script:M365GroupGraphIdCache = @{}
        }
        if ($matchedM365 -and $matchedM365.Count -gt 0) {
            Write-Log "Resolving SharePoint sites for $($matchedM365.Count) M365 group(s) in parallel..."
            $resolveStart = Get-Date

            # Two Graph calls per group: resolve ID, then fetch root site.
            # ThrottleLimit 10 keeps within Graph's parallel ceiling.
            $resolved = $matchedM365 | ForEach-Object -ThrottleLimit 10 -Parallel {
                $smtp = $_.PrimarySmtpAddress
                if ([string]::IsNullOrWhiteSpace($smtp)) { return $null }
                try {
                    # Step 1: resolve group ID
                    $idResp = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=mail eq '$smtp'&`$select=id" `
                                                    -Method GET -OutputType PSObject -ErrorAction Stop
                    $vals = $idResp.PSObject.Properties['value']
                    if (-not $vals -or @($vals.Value).Count -eq 0) { return $null }
                    $g0 = @($vals.Value)[0]
                    $graphId = $g0.PSObject.Properties['id']
                    if (-not $graphId) { return $null }
                    $graphIdValue = [string]$graphId.Value

                    # Step 2: fetch the group's root site
                    $siteResp = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/groups/$graphIdValue/sites/root?`$select=id,displayName,webUrl,description,createdDateTime" `
                                                     -Method GET -OutputType PSObject -ErrorAction Stop
                    [PSCustomObject]@{ Smtp = $smtp; GraphId = $graphIdValue; Site = $siteResp }
                } catch {
                    # Group has no site or we can't see it — silent
                    return $null
                }
            }

            $resolvedFromGroups = 0
            foreach ($r in $resolved) {
                if (-not $r) { continue }
                if ($r.GraphId) { $Script:M365GroupGraphIdCache[$r.Smtp.ToLower()] = $r.GraphId }
                if ($r.Site) {
                    $allSitesList.Add($r.Site) | Out-Null
                    $resolvedFromGroups++
                }
            }

            $resolveElapsed = (Get-Date) - $resolveStart
            Write-Log ("Resolved $resolvedFromGroups site(s) from M365 group associations in {0:mm\:ss} (cached $($Script:M365GroupGraphIdCache.Count) Graph IDs for re-use)." -f $resolveElapsed)
        }

        # ── Strategy 1.7: SPO Management Shell tenant enumeration (child process) ──
        # The upfront sign-in block at startup runs the SPO scan and writes its JSON output;
        # this block just loads it. If for some reason that didn't run (module missing,
        # tenant cache hit, sign-in failure), we fall through to live enumeration.
        if (-not $tenantWide -and (Get-Variable -Scope Script -Name 'SPOSitesJsonPath' -ErrorAction SilentlyContinue) -and $Script:SPOSitesJsonPath -and (Test-Path $Script:SPOSitesJsonPath)) {
            try {
                $spoSites = @(Get-Content $Script:SPOSitesJsonPath -Raw | ConvertFrom-Json)
                if ($spoSites.Count -gt 0) {
                    foreach ($spoSite in $spoSites) { $allSitesList.Add($spoSite) | Out-Null }
                    $tenantWide = $true
                    Write-Log "Loaded $($spoSites.Count) site(s) from upfront SPO scan — skipping mid-run SPO and search fallback." -Level SUCCESS
                }
                Remove-Item $Script:SPOSitesJsonPath -Force -ErrorAction SilentlyContinue
            } catch {
                Write-Log "Failed to load upfront SPO results: $($_.Exception.Message.Split("`n")[0])" -Level WARN
            }
        }

        # Mid-run SPO — fallback if the upfront scan didn't succeed (module missing, etc.)
        if (-not $tenantWide) {
            $spoModule = Get-Module -ListAvailable -Name 'Microsoft.Online.SharePoint.PowerShell' -ErrorAction SilentlyContinue
            if (-not $spoModule) {
                Write-Log "Microsoft.Online.SharePoint.PowerShell not installed — skipping tenant-wide enumeration." -Level WARN
                Write-Log "Install: Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser" -Level WARN
            } else {
                # Tenant admin URL: derive from any site URL we already have, else ask the user
                $adminUrl = $null
                if ($allSitesList.Count -gt 0) {
                    $sampleUrl = Get-ObjProp $allSitesList[0] 'webUrl'
                    # Extract just the tenant host: e.g. https://contoso.sharepoint.com/... → contoso
                    if ($sampleUrl -match '^https://([^\.\-/]+)\.sharepoint\.com') {
                        $tenantHost = $Matches[1]
                        $adminUrl = "https://$tenantHost-admin.sharepoint.com"
                    }
                }
                if (-not $adminUrl) {
                    Write-Log "Could not auto-detect tenant admin URL — skipping SPO enumeration." -Level WARN
                    Write-Log "  (To enable: re-run after at least one site has been discovered, or run export-tenant-sites.ps1.)" -Level WARN
                }

                if ($adminUrl) {
                    Write-Log "Spawning isolated PowerShell child process for SPO tenant enumeration..."
                    Write-Log "An interactive sign-in prompt may appear in a new window — sign in with a SharePoint admin account."
                    Write-Log "  (Note: section timer below includes the interactive sign-in wait time.)"
                    Write-Log "  Tenant admin URL: $adminUrl"

                    $spoJsonPath = Join-Path $OutputFolder "_SPOTenantSites.json"
                    $spoLogPath  = Join-Path $OutputFolder ("_SPOTenantSites_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".log")
                    $spoScript   = Join-Path $OutputFolder "_SPOTenantSites_scan.ps1"

                    # Same child script layout as the upfront block
                    @'
param(
    [Parameter(Mandatory)][string]$AdminUrl,
    [Parameter(Mandatory)][string]$JsonPath,
    [Parameter(Mandatory)][string]$LogPath
)
function Write-ChildLog {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$ts] [$($Level.PadRight(7))] [spo-child] $Message"
    Add-Content -Path $LogPath -Value $entry -Encoding UTF8
    Write-Host $entry
}

# ── Diagnostic environment block ──
Write-ChildLog "Environment: PSEdition=$($PSVersionTable.PSEdition) PSVersion=$($PSVersionTable.PSVersion)"
Write-ChildLog "Process    : $([System.Diagnostics.Process]::GetCurrentProcess().ProcessName) PID=$PID"
Write-ChildLog "Hostname   : $env:COMPUTERNAME  User=$env:USERNAME"
Write-ChildLog "AdminUrl   : $AdminUrl"

try {
    Import-Module 'Microsoft.Online.SharePoint.PowerShell' -DisableNameChecking -ErrorAction Stop
    $spoMod = Get-Module 'Microsoft.Online.SharePoint.PowerShell' | Select-Object -First 1
    if ($spoMod) {
        Write-ChildLog "SPO module loaded: version=$($spoMod.Version) path=$($spoMod.ModuleBase)"
    } else {
        Write-ChildLog "SPO module loaded but not visible via Get-Module" -Level WARN
    }

    Write-ChildLog "Calling Connect-SPOService -Url $AdminUrl ... (interactive sign-in window will appear)"
    $connectStart = Get-Date
    try {
        Connect-SPOService -Url $AdminUrl -ErrorAction Stop
        $connectElapsed = ((Get-Date) - $connectStart).TotalSeconds
        Write-ChildLog ("Connected to {0} in {1:F1}s" -f $AdminUrl, $connectElapsed) -Level SUCCESS
    } catch {
        $connectElapsed = ((Get-Date) - $connectStart).TotalSeconds
        Write-ChildLog ("Connect-SPOService failed after {0:F1}s" -f $connectElapsed) -Level ERROR
        Write-ChildLog "  Exception type    : $($_.Exception.GetType().FullName)" -Level ERROR
        Write-ChildLog "  Exception message : $($_.Exception.Message)" -Level ERROR
        if ($_.Exception.InnerException) {
            Write-ChildLog "  Inner type        : $($_.Exception.InnerException.GetType().FullName)" -Level ERROR
            Write-ChildLog "  Inner message     : $($_.Exception.InnerException.Message)" -Level ERROR
        }
        # Probe known causes
        if ($_.Exception.Message -match 'No valid OAuth 2\.0 authentication session') {
            Write-ChildLog "DIAGNOSTIC: 'No valid OAuth 2.0 session' typically means one of:" -Level ERROR
            Write-ChildLog "  - The sign-in window was closed before completion" -Level ERROR
            Write-ChildLog "  - The signed-in account lacks the SharePoint Administrator role" -Level ERROR
            Write-ChildLog "  - Conditional Access blocked the session (location/device/MFA)" -Level ERROR
            Write-ChildLog "  - The admin URL is wrong (must be https://<tenant>-admin.sharepoint.com)" -Level ERROR
        }
        if ($_.Exception.Message -match 'AADSTS') {
            $stsCode = ([regex]::Match($_.Exception.Message,'AADSTS\d+')).Value
            Write-ChildLog "DIAGNOSTIC: Azure AD token error code: $stsCode (search Microsoft docs for that code)" -Level ERROR
        }
        if ($_.ScriptStackTrace) {
            Write-ChildLog "  Stack trace       : $($_.ScriptStackTrace -replace [Environment]::NewLine,' || ')" -Level ERROR
        }
        throw
    }

    Write-ChildLog "Calling Get-SPOSite -Limit All ..."
    $listStart = Get-Date
    $sites = @(Get-SPOSite -Limit All -ErrorAction Stop)
    $listElapsed = ((Get-Date) - $listStart).TotalSeconds
    Write-ChildLog ("Get-SPOSite returned $($sites.Count) site(s) in {0:F1}s" -f $listElapsed) -Level SUCCESS

    $projected = $sites | ForEach-Object {
        [PSCustomObject]@{
            id              = if ($_.SiteId) { $_.SiteId.ToString() } else { '' }
            displayName     = $_.Title
            webUrl          = $_.Url
            description     = ''
            createdDateTime = ''
            template        = $_.Template
            storageUsageMB  = $_.StorageUsageCurrent
            storageQuotaMB  = $_.StorageQuota
            owner           = $_.Owner
            ownerEmail      = $_.Owner
            sharingCapability = $_.SharingCapability
            lastContentModifiedDate = if ($_.LastContentModifiedDate) { $_.LastContentModifiedDate.ToString('o') } else { '' }
            status          = $_.Status
        }
    }
    $projected | ConvertTo-Json -Depth 6 -Compress | Set-Content -Path $JsonPath -Encoding UTF8
    Write-ChildLog "Wrote $($projected.Count) site(s) to $(Split-Path $JsonPath -Leaf)" -Level SUCCESS

    Disconnect-SPOService
}
catch {
    Write-ChildLog "SPO enumeration failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" -Level ERROR
    exit 1
}
'@ | Set-Content -Path $spoScript -Encoding UTF8

                    try {
                        $childArgs = @(
                            '-NoProfile','-NoLogo','-ExecutionPolicy','Bypass',
                            '-File', "`"$spoScript`"",
                            '-AdminUrl', "`"$adminUrl`"",
                            '-JsonPath', "`"$spoJsonPath`"",
                            '-LogPath', "`"$spoLogPath`""
                        )
                        $spoExe = Resolve-SPOExecutable
                        Write-Log "SPO child will be launched in: $spoExe"
                        $proc = Start-Process -FilePath $spoExe -ArgumentList $childArgs -Wait -NoNewWindow -PassThru
                        if ($proc.ExitCode -eq 0 -and (Test-Path $spoJsonPath)) {
                            $spoSites = @(Get-Content $spoJsonPath -Raw | ConvertFrom-Json)
                            if ($spoSites.Count -gt 0) {
                                foreach ($spoSite in $spoSites) { $allSitesList.Add($spoSite) | Out-Null }
                                $tenantWide = $true
                                Write-Log "SPO returned $($spoSites.Count) tenant site(s) — strategy 2 (search) will be skipped." -Level SUCCESS
                            } else {
                                Write-Log "SPO enumeration returned 0 sites." -Level WARN
                            }
                        } else {
                            Write-Log "SPO child process exited with code $($proc.ExitCode). Check $(Split-Path $spoLogPath -Leaf)." -Level WARN
                        }
                        if (Test-Path $spoLogPath) {
                            Get-Content $spoLogPath | ForEach-Object { Add-Content -Path $Script:LogPath -Value $_ -Encoding UTF8 }
                        }
                    } catch {
                        Write-Log "Failed to launch SPO child process: $($_.Exception.Message.Split("`n")[0])" -Level ERROR
                    } finally {
                        if (Test-Path $spoScript) { Remove-Item $spoScript -Force -ErrorAction SilentlyContinue }
                        if (Test-Path $spoJsonPath) { Remove-Item $spoJsonPath -Force -ErrorAction SilentlyContinue }
                    }
                }
            }
        }

        # ── Strategy 2 (fallback): keyword searches ──
        if (-not $tenantWide) {
            $searchTerms = @($DomainPrefix, ($Domain -replace '\.',' ')) | Select-Object -Unique
            foreach ($term in $searchTerms) {
                if ([string]::IsNullOrWhiteSpace($term)) { continue }
                $q = UrlEncode $term
                try {
                    $part = @(Invoke-GraphGetAll -Uri "https://graph.microsoft.com/v1.0/sites?search=$q&$select")
                    if ($part.Count -gt 0) { $allSitesList.AddRange($part) }
                    Write-Log "Search '$term' returned $($part.Count) site(s)."
                } catch {
                    Write-Log "Search '$term' failed: $($_.Exception.Message.Split("`n")[0])" -Level WARN
                }
            }
        }

        # ── Build context sets for broader matching ──
        # Sites can reference the domain via:
        #   - Linked M365 group (group's mail/displayName)
        #   - Site owner's UPN
        # Section 2 already collected mailboxes and section 6 collected M365 groups,
        # so we reuse those rather than re-querying.
        $domainUpnSet = @{}
        if ($matchedUsers) {
            foreach ($u in $matchedUsers) {
                if ($u.UserPrincipalName) { $domainUpnSet[$u.UserPrincipalName.ToLower()] = $true }
                if ($u.PrimarySmtpAddress) { $domainUpnSet[$u.PrimarySmtpAddress.ToLower()] = $true }
            }
        }
        $domainGroupSet = @{}
        if ($matchedM365) {
            foreach ($g in $matchedM365) {
                if ($g.PrimarySmtpAddress) { $domainGroupSet[$g.PrimarySmtpAddress.ToLower()] = $true }
            }
        }
        Write-Log "Cross-reference sets: $($domainUpnSet.Count) user UPN(s), $($domainGroupSet.Count) M365 group SMTP(s)."

        # ── Local filtering: URL/Title only (fast path) ──
        # On large tenants the previous multi-property + Graph-enrichment approach took hours.
        # We now match purely on URL/Title (case-insensitive) and use the cached storage
        # values from the SPO Management Shell directly. Trade-off: sites whose only domain
        # reference is in description, owner, or M365 group association will be missed.
        # Use direct property access ($_.webUrl etc.) — Get-ObjProp adds significant overhead
        # when called millions of times across 30k+ sites.
        $allSites = @($allSitesList | Sort-Object { if ($_.PSObject.Properties['id']) { $_.PSObject.Properties['id'].Value } else { $null } } -Unique)
        Write-Log "Filtering $($allSites.Count) candidate site(s) by URL/Title for '$Domain' or '$DomainPrefix'..."
        $filterStart = Get-Date

        $matchedSP = [System.Collections.Generic.List[object]]::new()
        $matchedOD = [System.Collections.Generic.List[object]]::new()
        $i = 0
        $progressEvery = [math]::Max(1000, [int]($allSites.Count / 30))

        foreach ($site in $allSites) {
            $i++
            if ($i % $progressEvery -eq 0) {
                Write-Progress -Activity "Filtering sites" -Status "$i / $($allSites.Count)" `
                    -PercentComplete (($i / [System.Math]::Max($allSites.Count,1)) * 100)
            }

            # StrictMode-safe property reads — accessing a missing property directly throws.
            # PSObject.Properties[name] returns $null for missing properties without throwing.
            $props = $site.PSObject.Properties
            $wu = if ($props['webUrl'])     { $props['webUrl'].Value }     else { $null }
            $dn = if ($props['displayName']){ $props['displayName'].Value }else { $null }

            # Two -like checks: URL and Title. Match on full domain OR prefix.
            $matched = $false
            $matchHit = $null
            if ($wu) {
                if     ($wu -like "*$Domain*")       { $matched = $true; $matchHit = 'WebUrl(domain)' }
                elseif ($wu -like "*$DomainPrefix*") { $matched = $true; $matchHit = 'WebUrl(prefix)' }
            }
            if (-not $matched -and $dn) {
                if     ($dn -like "*$Domain*")       { $matched = $true; $matchHit = 'Title(domain)' }
                elseif ($dn -like "*$DomainPrefix*") { $matched = $true; $matchHit = 'Title(prefix)' }
            }
            if (-not $matched) { continue }

            # All cached property reads — guarded for StrictMode
            $usedMB  = if ($props['storageUsageMB'])  { $props['storageUsageMB'].Value }  else { $null }
            $quotaMB = if ($props['storageQuotaMB']) { $props['storageQuotaMB'].Value } else { $null }
            $sid     = if ($props['id'])              { $props['id'].Value }              else { $null }
            $desc    = if ($props['description'])     { $props['description'].Value }     else { $null }
            $cdt     = if ($props['createdDateTime']) { $props['createdDateTime'].Value } else { $null }
            $oe      = if ($props['ownerEmail'])      { $props['ownerEmail'].Value }      else { $null }
            $lcm     = if ($props['lastContentModifiedDate']) { $props['lastContentModifiedDate'].Value } else { $null }

            $usedGB    = if ($usedMB)  { [math]::Round([double]$usedMB  / 1024, 2) } else { $null }
            $quotaGB   = if ($quotaMB) { [math]::Round([double]$quotaMB / 1024, 2) } else { $null }
            $usedBytes = if ($usedMB)  { [int64]([double]$usedMB * 1MB) } else { $null }

            $lastMod = $null
            if ($lcm) { try { $lastMod = [datetime]$lcm } catch {} }

            $isOneDrive = ($wu -like '*-my.sharepoint.com/personal/*')

            $row = [PSCustomObject]@{
                DisplayName     = $dn
                WebUrl          = $wu
                SiteId          = $sid
                Description     = $desc
                CreatedDateTime = $cdt
                StorageUsedGB   = $usedGB
                StorageQuotaGB  = $quotaGB
                StorageUsedBytes= $usedBytes
                ItemCount       = $null
                LastModified    = $lastMod
                Owner           = $oe
                Status          = 'OK'
                MatchedOn       = $matchHit
            }

            if ($isOneDrive) { $matchedOD.Add($row) | Out-Null }
            else             { $matchedSP.Add($row) | Out-Null }
        }
        Write-Progress -Activity "Filtering sites" -Completed
        $filterElapsed = (Get-Date) - $filterStart
        Write-Log ("Matched $($matchedSP.Count) SharePoint site(s); $($matchedOD.Count) OneDrive personal site(s) in {0:mm\:ss}." -f $filterElapsed)
        Write-Log "(Storage values from SPO cache; full OneDrive enumeration in section 10.)"

        # Strip the OneDrive-only ItemCount column for the SP CSV (always blank for SP)
        $spExport = @($matchedSP | Select-Object DisplayName, WebUrl, SiteId, Description, CreatedDateTime,
                                                StorageUsedGB, StorageQuotaGB, StorageUsedBytes, LastModified, Owner, Status, MatchedOn)
        Export-SafeCsv -Path (Join-Path $DiscoveryFolder '09_SharePointSites.csv') -Data @($spExport) -Label 'SharePoint Sites'

        # Stash for the next section so we don't have to re-discover
        $Script:MatchedOneDriveSites = $matchedOD
    }
} catch { Write-Log "Error: $($_.Exception.Message.Split("`n")[0])" -Level ERROR }
End-Section

# ─────────────────────────────────────────────────────────────
# 10. ONEDRIVES (Graph) — sizes + item counts
# ─────────────────────────────────────────────────────────────
Start-Section "[10/22] OneDrives (Graph)"
try {
    if (-not $GraphAvailable) {
        Write-Log "Graph unavailable — skipping OneDrives." -Level WARN
    } else {
        # ── Build union of users to scan: AD users (section 2) + Graph users (domain UPN) ──
        $userMap = @{}  # key = lowered UPN, value = @{ UPN, DisplayName, Source }

        # Source 1: AD users from section 2
        if (Get-Variable -Scope Script -Name 'matchedUsers' -ErrorAction SilentlyContinue) {
            foreach ($u in @($Script:matchedUsers)) {
                $upn = $u.UserPrincipalName
                if ([string]::IsNullOrWhiteSpace($upn)) { continue }
                $key = $upn.ToLower()
                $userMap[$key] = [PSCustomObject]@{
                    UPN         = $upn
                    DisplayName = $u.DisplayName
                    Source      = 'AD'
                }
            }
        } elseif ($matchedUsers) {
            foreach ($u in @($matchedUsers)) {
                $upn = $u.UserPrincipalName
                if ([string]::IsNullOrWhiteSpace($upn)) { continue }
                $key = $upn.ToLower()
                $userMap[$key] = [PSCustomObject]@{
                    UPN         = $upn
                    DisplayName = $u.DisplayName
                    Source      = 'AD'
                }
            }
        }

        # Source 2: shared/resource mailboxes from section 5 (avoids re-enumerating the whole tenant)
        $sharedCount = 0
        if (Get-Variable -Name 'matchedShared' -ErrorAction SilentlyContinue) {
            foreach ($s in @($matchedShared)) {
                $upn = $s.UserPrincipalName
                if ([string]::IsNullOrWhiteSpace($upn)) { continue }
                $key = $upn.ToLower()
                if (-not $userMap.ContainsKey($key)) {
                    $userMap[$key] = [PSCustomObject]@{
                        UPN         = $upn
                        DisplayName = $s.DisplayName
                        Source      = 'SharedMailbox'
                    }
                    $sharedCount++
                }
            }
        }
        Write-Log "OneDrive scope: $($userMap.Count - $sharedCount) user mailbox(es) + $sharedCount shared/resource mailbox(es) = $($userMap.Count) total."

        # ── Resolve each user's OneDrive in parallel ──────────────────────────
        # Sequential lookups at ~1-2 s/call become impractical beyond ~100 users.
        # ThrottleLimit 15 keeps well within Graph's per-app concurrency ceiling.
        Write-Log "Querying $($userMap.Count) OneDrive(s) in parallel (throttle = 15)..."
        $odStart = Get-Date

        $parallelResults = @($userMap.Values) | ForEach-Object -ThrottleLimit 15 -Parallel {
            $entry = $_
            $upn   = $entry.UPN

            $row = [PSCustomObject]@{
                UserPrincipalName = $upn
                DisplayName       = $entry.DisplayName
                Source            = $entry.Source
                WebUrl            = $null; DriveId = $null; Owner = $upn
                StorageUsedGB     = $null; StorageQuotaGB = $null; StorageUsedBytes = $null
                ItemCount         = $null; LastModified = $null; Status = 'NotChecked'
            }

            try {
                $encUpn = [System.Uri]::EscapeDataString($upn)
                $drive  = Invoke-MgGraphRequest `
                    -Uri "https://graph.microsoft.com/v1.0/users/$encUpn/drive?`$select=id,webUrl,quota,lastModifiedDateTime" `
                    -Method GET -OutputType PSObject -ErrorAction Stop

                # Inline property access — helper functions not available in parallel runspaces.
                $p = $drive.PSObject.Properties
                $row.DriveId = if ($p['id'])     { $p['id'].Value }     else { $null }
                $row.WebUrl  = if ($p['webUrl']) { $p['webUrl'].Value } else { $null }
                $lm = if ($p['lastModifiedDateTime']) { $p['lastModifiedDateTime'].Value } else { $null }
                if ($lm) { try { $row.LastModified = [datetime]$lm } catch {} }

                $q = if ($p['quota']) { $p['quota'].Value } else { $null }
                if ($q) {
                    $qp = $q.PSObject.Properties
                    $u = [int64](if ($qp['used'])  { $qp['used'].Value  } else { 0 })
                    $t = [int64](if ($qp['total']) { $qp['total'].Value } else { 0 })
                    $row.StorageUsedBytes = $u
                    $row.StorageUsedGB    = [math]::Round($u / 1GB, 2)
                    $row.StorageQuotaGB   = [math]::Round($t / 1GB, 2)
                }

                if ($row.DriveId) {
                    try {
                        $root = Invoke-MgGraphRequest `
                            -Uri "https://graph.microsoft.com/v1.0/drives/$($row.DriveId)/root?`$select=folder" `
                            -Method GET -OutputType PSObject -ErrorAction Stop
                        $fp = $root.PSObject.Properties['folder']
                        if ($fp -and $fp.Value) {
                            $cp = $fp.Value.PSObject.Properties['childCount']
                            if ($cp) { $row.ItemCount = $cp.Value }
                        }
                    } catch {}
                }

                $row.Status = 'OK'
            } catch {
                $code = $null
                try { $code = [int]$_.Exception.Response.StatusCode } catch {}
                $msg  = $_.ToString()
                $row.Status = if     ($code -eq 404 -or $msg -match 'mySite\s*Not\s*Found|ResourceNotFound|HTTP 404') { 'NotProvisioned' }
                              elseif ($code -eq 403 -or $msg -match 'Forbidden|HTTP 403')                             { 'PermissionDenied' }
                              else { "Error: $($_.Exception.Message.Split([Environment]::NewLine)[0])" }
            }
            $row
        }

        $odElapsed = (Get-Date) - $odStart
        Write-Log ("Parallel OneDrive lookup completed in {0:mm\:ss}." -f $odElapsed)

        $oneDriveResults = [System.Collections.Generic.List[object]]::new()
        $statOk = 0; $statNone = 0; $stat403 = 0; $statErr = 0
        foreach ($r in $parallelResults) {
            if (-not $r) { continue }
            $oneDriveResults.Add($r) | Out-Null
            switch ($r.Status) {
                'OK'              { $statOk++   }
                'NotProvisioned'  { $statNone++ }
                'PermissionDenied'{ $stat403++  }
                default           { if ($r.Status -ne 'NotChecked') { $statErr++ } }
            }
        }
        Write-Log "OneDrives — ${statOk} OK | ${statNone} not provisioned | ${stat403} 403 | ${statErr} other error(s)"

        Export-SafeCsv -Path (Join-Path $DiscoveryFolder '10_OneDrives.csv') -Data @($oneDriveResults) -Label 'OneDrives'
    }
} catch { Write-Log "Error: $($_.Exception.Message.Split("`n")[0])" -Level ERROR }
End-Section

# ─────────────────────────────────────────────────────────────
# 11. TEAMS (derived from M365 groups)
# ─────────────────────────────────────────────────────────────
Start-Section "[11/22] Microsoft Teams (derived from M365 Groups)"
try {
    $matchedTeams = @(
        $matchedM365 |
        Where-Object { $_.IsTeam -eq $true } |
        Select-Object DisplayName, PrimarySmtpAddress, Alias, SharePointSiteUrl, AccessType, MatchedOn
    )
    Export-SafeCsv -Path (Join-Path $DiscoveryFolder '11_Teams.csv') -Data @($matchedTeams) -Label 'Microsoft Teams'
} catch { Write-Log "Error: $($_.Exception.Message.Split("`n")[0])" -Level ERROR }
End-Section

# ─────────────────────────────────────────────────────────────
# 12. ENTRA REGISTERED DEVICES (Graph)
# ─────────────────────────────────────────────────────────────
Start-Section "[12/22] Entra Registered Devices (Graph)"
try {
    $DeviceResults = [System.Collections.Generic.List[object]]::new()
    if (-not $GraphAvailable) { Write-Log "Graph unavailable — skipping Devices." -Level WARN }
    else {
        # Retrieve all member users in the connected (source) tenant.
        $userUri = "https://graph.microsoft.com/v1.0/users?`$filter=userType eq 'Member'&`$select=id,displayName,userPrincipalName,accountEnabled"
        $domainUsers = @(Invoke-GraphGetAll -Uri $userUri)

        if ($domainUsers.Count -eq 0) {
            Write-Log "No users found matching '*$Domain' — no devices to scan." -Level WARN
        } else {
            Write-Log "Found $($domainUsers.Count) user(s) to scan for registered devices."

            $total             = $domainUsers.Count
            $i                 = 0
            $usersWithDevices  = 0
            $usersWithoutDevices = 0
            $totalDeviceCount  = 0

            foreach ($u in $domainUsers) {
                $i++
                $upn = Get-ObjProp $u 'userPrincipalName'
                $uid = Get-ObjProp $u 'id'
                $dn  = Get-ObjProp $u 'displayName'
                $enabled = Get-ObjProp $u 'accountEnabled'

                Write-Progress -Activity "Entra Registered Devices" `
                    -Status "$upn  ($i / $total)" `
                    -PercentComplete (($i / [System.Math]::Max($total, 1)) * 100)

                try {
                    # Fetch registered devices with full property set; handle pagination.
                    $devUri  = "https://graph.microsoft.com/v1.0/users/$uid/registeredDevices?`$select=id,deviceId,displayName,operatingSystem,operatingSystemVersion,trustType,approximateLastSignInDateTime"
                    $devices = [System.Collections.Generic.List[object]]::new()
                    $nextUri = $devUri

                    do {
                        $resp  = Invoke-MgGraphRequest -Uri $nextUri -Method GET -OutputType PSObject
                        $page  = @(Get-ObjProp $resp 'value')
                        if ($page) { $devices.AddRange($page) }
                        $nextUri = Get-ObjProp $resp '@odata.nextLink'
                    } while ($nextUri)

                    if ($devices.Count -eq 0) {
                        $usersWithoutDevices++
                    } else {
                        $usersWithDevices++
                        $totalDeviceCount += $devices.Count

                        foreach ($d in $devices) {
                            Write-Progress -Id 2 -ParentId 0 `
                                -Activity "Devices for $upn" `
                                -Status (Get-ObjProp $d 'displayName')

                            $DeviceResults.Add([PSCustomObject]@{
                                OwnerUPN       = $upn
                                OwnerName      = $dn
                                AccountEnabled = $enabled
                                DeviceName     = Get-ObjProp $d 'displayName'
                                DeviceObjectId = Get-ObjProp $d 'id'
                                EntraDeviceId  = Get-ObjProp $d 'deviceId'
                                OS             = Get-ObjProp $d 'operatingSystem'
                                OSVersion      = Get-ObjProp $d 'operatingSystemVersion'
                                TrustType      = Get-ObjProp $d 'trustType'
                                LastSignIn     = Get-ObjProp $d 'approximateLastSignInDateTime'
                                DiscoveredAt   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
                            }) | Out-Null
                        }
                    }
                } catch {
                    Write-Log "Device lookup failed for ${upn}: $($_.Exception.Message.Split("`n")[0])" -Level WARN
                }
            }

            Write-Progress -Activity "Entra Registered Devices" -Completed
            Write-Progress -Id 2 -Completed

            Write-Log "Devices complete — scanned: $total  |  with: $usersWithDevices  |  without: $usersWithoutDevices  |  total devices: $totalDeviceCount"
        }
    }
    Export-SafeCsv -Path (Join-Path $DiscoveryFolder '12_Devices.csv') `
                   -Data @($DeviceResults) `
                   -Label 'Entra Registered Devices'
} catch { Write-Log "Error in device scan: $_" -Level ERROR }
End-Section

# ─────────────────────────────────────────────────────────────
# 13. PROXY ADDRESSES TO REMOVE
# ─────────────────────────────────────────────────────────────
Start-Section "[13/22] Proxy Addresses to Remove"
try {
    $proxyFilter = "EmailAddresses -like '*$Domain*'"
    $recipients  = @(Get-Recipient -Filter $proxyFilter -ResultSize Unlimited -WarningAction SilentlyContinue)
    Write-Log "Found $($recipients.Count) recipient(s) with addresses matching '$Domain'"

    $proxyResults = [System.Collections.Generic.List[object]]::new()

    foreach ($r in $recipients) {
        $matchingAddresses = @($r.EmailAddresses | Where-Object { $_ -like "*$Domain*" })

        foreach ($addr in $matchingAddresses) {
            $addrStr = $addr.ToString()
            if ($addrStr -match '^([^:]+):(.+)$') {
                $addrType    = $Matches[1].ToUpper()
                $addrAddress = $Matches[2]
            } else {
                $addrType    = ''
                $addrAddress = $addrStr
            }

            $proxyResults.Add([PSCustomObject]@{
                RecipientType      = $r.RecipientTypeDetails
                DisplayName        = $r.DisplayName
                PrimarySmtpAddress = $r.PrimarySmtpAddress
                UserPrincipalName  = $r.WindowsLiveID
                ProxyAddress       = $addrAddress
                AddressType        = $addrType
                IsPrimary          = ($Matches[1] -ceq 'SMTP')
            }) | Out-Null
        }
    }

    Export-SafeCsv -Path (Join-Path $DiscoveryFolder '13_ProxyAddresses.csv') -Data @($proxyResults) -Label 'Proxy Addresses to Remove'
} catch { Write-Log "Error: $($_.Exception.Message.Split("`n")[0])" -Level ERROR }
End-Section


# ─────────────────────────────────────────────────────────────
# 14. EXCHANGE TRANSPORT RULES
# ─────────────────────────────────────────────────────────────
Start-Section "[14/22] Exchange Transport Rules"
try {
    $condProps = @(
        'SenderDomainIs','RecipientDomainIs',
        'FromAddressContainsWords','AnyOfRecipientAddressContainsWords',
        'FromAddressMatchesPatterns','RecipientAddressMatchesPatterns'
    )
    $allRules     = @(Get-TransportRule -ResultSize Unlimited -WarningAction SilentlyContinue)
    $matchedRules = [System.Collections.Generic.List[object]]::new()

    foreach ($rule in $allRules) {
        $hits = [System.Collections.Generic.List[string]]::new()
        foreach ($prop in $condProps) {
            try {
                $val = $rule.$prop
                if ($val -and @($val | Where-Object { $_ -like "*$Domain*" }).Count -gt 0) {
                    $hits.Add($prop)
                }
            } catch {}
        }
        if ($rule.Description -like "*$Domain*") { $hits.Add('Description') }
        if ($rule.Comments    -like "*$Domain*") { $hits.Add('Comments') }

        if ($hits.Count -gt 0) {
            $matchedRules.Add([PSCustomObject]@{
                Name        = $rule.Name
                Identity    = $rule.Identity
                State       = $rule.State
                Priority    = $rule.Priority
                Description = $rule.Description
                MatchedOn   = ($hits -join ', ')
            }) | Out-Null
        }
    }
    Export-SafeCsv -Path (Join-Path $DiscoveryFolder '14_TransportRules.csv') -Data @($matchedRules) -Label 'Transport Rules'
} catch { Write-Log "Error: $($_.Exception.Message.Split("`n")[0])" -Level ERROR }
End-Section

# ─────────────────────────────────────────────────────────────
# 15. PLANNER PLANS (Graph — via M365 Groups)
# ─────────────────────────────────────────────────────────────
Start-Section "[15/22] Planner Plans (via M365 Groups)"
try {
    if (-not $GraphAvailable) { Write-Log "Graph unavailable — skipping Planner." -Level WARN }
    elseif ($matchedM365.Count -eq 0) { Write-Log "No M365 Groups discovered — skipping Planner." -Level WARN }
    else {
        $plannerResults = [System.Collections.Generic.List[object]]::new()
        $i = 0
        $status403 = 0
        $statusOk  = 0
        $statusEmpty = 0

        foreach ($grp in $matchedM365) {
            $i++
            $smtp = $grp.PrimarySmtpAddress
            Write-Progress -Activity "Planner Plans" -Status "$i / $($matchedM365.Count): $smtp" `
                -PercentComplete (($i / [System.Math]::Max($matchedM365.Count,1)) * 100)
            try {
                # Try the cache populated by section 9 first; fall back to a Graph lookup
                $graphId = $null
                if (Get-Variable -Scope Script -Name 'M365GroupGraphIdCache' -ErrorAction SilentlyContinue) {
                    $cacheKey = $smtp.ToLower()
                    if ($Script:M365GroupGraphIdCache.ContainsKey($cacheKey)) {
                        $graphId = $Script:M365GroupGraphIdCache[$cacheKey]
                    }
                }
                if (-not $graphId) {
                    $gResult = @(Invoke-GraphGetAll -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=mail eq '$smtp'&`$select=id,displayName,mail")
                    if ($gResult.Count -gt 0) { $graphId = Get-ObjProp $gResult[0] 'id' }
                }
                if (-not $graphId) {
                    Write-Log "Could not resolve Graph ID for group: $smtp" -Level WARN
                    $plannerResults.Add([PSCustomObject]@{
                        GroupDisplayName = $grp.DisplayName
                        GroupEmail       = $smtp
                        GraphGroupId     = $null
                        PlanTitle        = '(group not resolvable)'
                        PlanId           = $null
                        BucketCount      = $null
                        TaskCount        = $null
                        CreatedDateTime  = $null
                        Status           = 'GroupNotFound'
                        Action           = 'Investigate — group does not exist in Graph'
                    }) | Out-Null
                    continue
                }

                # Direct call so we can distinguish 403 from "empty" reliably.
                $plans = @()
                $status = 'OK'
                try {
                    $resp  = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/groups/$graphId/planner/plans?`$select=id,title,createdDateTime" `
                                                    -Method GET -OutputType PSObject -ErrorAction Stop
                    $plans = @(Get-ObjProp $resp 'value')
                    $statusOk++
                } catch {
                    $code = $null
                    try { $code = [int]$_.Exception.Response.StatusCode } catch {}
                    if ($code -eq 403 -or $_.ToString() -match 'Forbidden|HTTP 403') {
                        $status = 'PermissionDenied'
                        $status403++
                    } elseif ($code -eq 404) {
                        $status = 'NotFound'
                    } else {
                        $status = "Error: $($_.Exception.Message.Split("`n")[0])"
                    }
                }

                if ($plans.Count -eq 0 -and $status -eq 'OK') { $statusEmpty++; continue }

                if ($status -ne 'OK') {
                    # Permission/access failure — still record the group so it shows up in the migration plan
                    $plannerResults.Add([PSCustomObject]@{
                        GroupDisplayName = $grp.DisplayName
                        GroupEmail       = $smtp
                        GraphGroupId     = $graphId
                        PlanTitle        = '(plan list unavailable)'
                        PlanId           = $null
                        BucketCount      = $null
                        TaskCount        = $null
                        CreatedDateTime  = $null
                        Status           = $status
                        Action           = 'Verify Planner contents in source tenant; recreate manually'
                    }) | Out-Null
                    continue
                }

                foreach ($plan in $plans) {
                    $planId  = Get-ObjProp $plan 'id'
                    $bucketCount = $null; $taskCount = $null
                    try { $bucketCount = @(Invoke-GraphGetAll -Uri "https://graph.microsoft.com/v1.0/planner/plans/$planId/buckets?`$select=id,name").Count } catch {}
                    try { $taskCount   = @(Invoke-GraphGetAll -Uri "https://graph.microsoft.com/v1.0/planner/plans/$planId/tasks?`$select=id,title,percentComplete,dueDateTime").Count } catch {}

                    $plannerResults.Add([PSCustomObject]@{
                        GroupDisplayName = $grp.DisplayName
                        GroupEmail       = $smtp
                        GraphGroupId     = $graphId
                        PlanTitle        = Get-ObjProp $plan 'title'
                        PlanId           = $planId
                        BucketCount      = $bucketCount
                        TaskCount        = $taskCount
                        CreatedDateTime  = Get-ObjProp $plan 'createdDateTime'
                        Status           = 'OK'
                        Action           = 'Recreate in new tenant'
                    }) | Out-Null
                }
            } catch {
                Write-Log "Planner lookup failed for ${smtp}: $($_.Exception.Message.Split("`n")[0])" -Level WARN
                $plannerResults.Add([PSCustomObject]@{
                    GroupDisplayName = $grp.DisplayName
                    GroupEmail       = $smtp
                    GraphGroupId     = $null
                    PlanTitle        = '(lookup failed)'
                    PlanId           = $null
                    BucketCount      = $null
                    TaskCount        = $null
                    CreatedDateTime  = $null
                    Status           = "Error: $($_.Exception.Message.Split("`n")[0])"
                    Action           = 'Investigate manually'
                }) | Out-Null
            }
        }
        Write-Progress -Activity "Planner Plans" -Completed
        Write-Log "Planner — ${statusOk} group(s) read OK, ${statusEmpty} had no plans, ${status403} returned 403 (permission denied)."
        Export-SafeCsv -Path (Join-Path $DiscoveryFolder '15_PlannerPlans.csv') -Data @($plannerResults) -Label 'Planner Plans'
    }
} catch { Write-Log "Error: $($_.Exception.Message.Split("`n")[0])" -Level ERROR }
End-Section

# ─────────────────────────────────────────────────────────────
# 16. POWER PLATFORM (Apps + Flows — optional module)
# ─────────────────────────────────────────────────────────────
Start-Section "[16/22] Power Platform (Apps and Flows)"
try {
    # Skip if the upfront scan already produced the CSV (always the case in normal runs).
    $ppDoneUpfront = $false
    if ((Get-Variable -Scope Script -Name 'PowerPlatformCsv' -ErrorAction SilentlyContinue) -and $Script:PowerPlatformCsv -and (Test-Path $Script:PowerPlatformCsv)) {
        Write-Log "Power Platform was scanned upfront — using cached CSV at $(Split-Path $Script:PowerPlatformCsv -Leaf)." -Level SUCCESS
        $ppDoneUpfront = $true
    } elseif ((Get-Variable -Scope Script -Name 'RunPowerPlatform' -ErrorAction SilentlyContinue) -and (-not $Script:RunPowerPlatform)) {
        Write-Log "Power Platform scan was declined at startup — skipping." -Level WARN
        $ppDoneUpfront = $true
    }

    if (-not $ppDoneUpfront) {
        $ppModule = Get-Module -ListAvailable -Name 'Microsoft.PowerApps.Administration.PowerShell' -ErrorAction SilentlyContinue
        if (-not $ppModule) {
            Write-Log "Microsoft.PowerApps.Administration.PowerShell not installed — skipping." -Level WARN
            Write-Log "Install: Install-Module Microsoft.PowerApps.Administration.PowerShell -Scope CurrentUser" -Level WARN
        } else {
            # The Power Platform module ships its own copy of Microsoft.Identity.Client.dll which
            # conflicts with the version Microsoft.Graph already loaded into this session.
            # Workaround: spawn a clean child pwsh.exe process that has no Graph SDK in it.
            Write-Log "Spawning isolated PowerShell child process to avoid Microsoft.Identity.Client conflict..."
            Write-Log "An interactive sign-in prompt may appear in a new window — sign in with a Power Platform admin account."
            Write-Log "  (Note: section timer below includes the interactive sign-in wait time.)"

            $ppCsvPath  = Join-Path $DiscoveryFolder '16_PowerPlatform.csv'
        $ppLogPath  = Join-Path $OutputFolder ("_PowerPlatform_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".log")
        $ppScript   = Join-Path $OutputFolder "_PowerPlatform_scan.ps1"

        # Write the child script. Single-quoted here-string so $Domain etc. are NOT expanded
        # in this parent process; we'll inject the actual values via -Args instead.
        @'
param(
    [Parameter(Mandatory)][string]$Domain,
    [Parameter(Mandatory)][string]$DomainPrefix,
    [Parameter(Mandatory)][string]$CsvPath,
    [Parameter(Mandatory)][string]$LogPath
)

function Write-ChildLog {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$ts] [$($Level.PadRight(7))] [child] $Message"
    Add-Content -Path $LogPath -Value $entry -Encoding UTF8
    Write-Host $entry
}

try {
    Import-Module 'Microsoft.PowerApps.Administration.PowerShell' -ErrorAction Stop
    Write-ChildLog "Power Platform module loaded in clean session."

    Add-PowerAppsAccount -ErrorAction Stop
    Write-ChildLog "Authenticated to Power Platform." -Level SUCCESS

    $results = [System.Collections.Generic.List[object]]::new()

    # Power Apps
    Write-ChildLog "Scanning Power Apps..."
    try {
        $apps = @(Get-AdminPowerApp -ErrorAction Stop | Where-Object {
            ($_.Owner.UserPrincipalName -like "*$Domain*") -or
            ($_.Internal.displayName    -like "*$DomainPrefix*")
        })
        foreach ($a in $apps) {
            $results.Add([PSCustomObject]@{
                ObjectType   = 'PowerApp'
                DisplayName  = $a.Internal.displayName
                ObjectId     = $a.AppName
                Owner        = $a.Owner.UserPrincipalName
                Environment  = $a.EnvironmentName
                CreatedTime  = $a.Internal.createdTime
                LastModified = $a.Internal.lastModifiedTime
                Action       = 'Export and re-import; reassign owner in new tenant'
            }) | Out-Null
        }
        Write-ChildLog "Found $($apps.Count) Power App(s) matching domain."
    } catch { Write-ChildLog "Power Apps scan failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" -Level WARN }

    # Power Automate Flows
    Write-ChildLog "Scanning Power Automate flows..."
    try {
        $envs = @(Get-AdminPowerAppEnvironment -ErrorAction Stop)
        foreach ($env in $envs) {
            try {
                $flows = @(Get-AdminFlow -EnvironmentName $env.EnvironmentName -ErrorAction Stop | Where-Object {
                    ($_.Internal.properties.creator.userPrincipalName -like "*$Domain*") -or
                    ($_.Internal.properties.displayName               -like "*$DomainPrefix*")
                })
                foreach ($f in $flows) {
                    $results.Add([PSCustomObject]@{
                        ObjectType   = 'PowerAutomateFlow'
                        DisplayName  = $f.Internal.properties.displayName
                        ObjectId     = $f.FlowName
                        Owner        = $f.Internal.properties.creator.userPrincipalName
                        Environment  = $env.DisplayName
                        CreatedTime  = $f.Internal.properties.creationTime
                        LastModified = $f.Internal.properties.lastModifiedTime
                        Action       = 'Export and re-import; re-authenticate all connections'
                    }) | Out-Null
                }
            } catch {
                Write-ChildLog "Flows in environment '$($env.DisplayName)' failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" -Level WARN
            }
        }
    } catch { Write-ChildLog "Power Automate scan failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" -Level WARN }

    if ($results.Count -gt 0) {
        $results | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8 -Force
        Write-ChildLog "Exported $($results.Count) record(s) -> $(Split-Path $CsvPath -Leaf)" -Level SUCCESS
    } else {
        # Placeholder row matches Export-SafeCsv format
        ([PSCustomObject]@{
            Status    = 'No Power Platform found'
            Domain    = $Domain
            Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        }) | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8 -Force
        Write-ChildLog "No Power Platform objects matched — wrote placeholder row." -Level WARN
    }
}
catch {
    Write-ChildLog "Power Platform scan failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" -Level ERROR
    exit 1
}
'@ | Set-Content -Path $ppScript -Encoding UTF8

        try {
            $childArgs = @(
                '-NoProfile','-NoLogo','-ExecutionPolicy','Bypass',
                '-File', "`"$ppScript`"",
                '-Domain', $Domain,
                '-DomainPrefix', $DomainPrefix,
                '-CsvPath', "`"$ppCsvPath`"",
                '-LogPath', "`"$ppLogPath`""
            )
            $proc = Start-Process -FilePath 'pwsh.exe' -ArgumentList $childArgs -Wait -NoNewWindow -PassThru
            if ($proc.ExitCode -eq 0) {
                Write-Log "Power Platform child process completed (exit 0). Detailed log: $(Split-Path $ppLogPath -Leaf)" -Level SUCCESS
            } else {
                Write-Log "Power Platform child process exited with code $($proc.ExitCode). Check $(Split-Path $ppLogPath -Leaf)." -Level WARN
            }

            # Surface child log into the main log so everything is in one place
            if (Test-Path $ppLogPath) {
                Get-Content $ppLogPath | ForEach-Object { Add-Content -Path $Script:LogPath -Value $_ -Encoding UTF8 }
            }
        } catch {
            Write-Log "Failed to launch Power Platform child process: $($_.Exception.Message.Split("`n")[0])" -Level ERROR
        } finally {
            # Clean up the temp script file
            if (Test-Path $ppScript) { Remove-Item $ppScript -Force -ErrorAction SilentlyContinue }
        }
        }
    }
} catch { Write-Log "Error: $($_.Exception.Message.Split("`n")[0])" -Level ERROR }
End-Section

# ─────────────────────────────────────────────────────────────
# 17. ENTRA CONDITIONAL ACCESS POLICIES
# ─────────────────────────────────────────────────────────────
Start-Section "[17/22] Entra Conditional Access Policies"
try {
    if (-not $GraphAvailable) { Write-Log "Graph unavailable — skipping CA Policies." -Level WARN }
    else {
        $caResults = [System.Collections.Generic.List[object]]::new()

        $policies = @(Invoke-GraphGetAll -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies?`$select=id,displayName,state,conditions")
        foreach ($pol in $policies) {
            $polJson = ($pol | ConvertTo-Json -Depth 20 -Compress)
            if ($polJson -like "*$Domain*" -or $polJson -like "*$DomainPrefix*") {
                $caResults.Add([PSCustomObject]@{
                    ObjectType  = 'ConditionalAccessPolicy'
                    DisplayName = Get-ObjProp $pol 'displayName'
                    ObjectId    = Get-ObjProp $pol 'id'
                    State       = Get-ObjProp $pol 'state'
                    Detail      = 'Policy JSON references domain — review conditions manually'
                    Action      = 'Review and update conditions in Entra portal'
                }) | Out-Null
            }
        }

        $locations = @(Invoke-GraphGetAll -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations?`$select=id,displayName,createdDateTime")
        foreach ($loc in $locations) {
            $dn = Get-ObjProp $loc 'displayName'
            if ($dn -like "*$DomainPrefix*") {
                $caResults.Add([PSCustomObject]@{
                    ObjectType  = 'NamedLocation'
                    DisplayName = $dn
                    ObjectId    = Get-ObjProp $loc 'id'
                    State       = 'N/A'
                    Detail      = 'Display name references domain prefix'
                    Action      = 'Rename or remove'
                }) | Out-Null
            }
        }

        Write-Log "Found $($caResults.Count) CA policy/location reference(s)."
        Export-SafeCsv -Path (Join-Path $DiscoveryFolder '17_EntraCA_Policies.csv') -Data @($caResults) -Label 'Entra CA Policies'
    }
} catch { Write-Log "Error: $($_.Exception.Message.Split("`n")[0])" -Level ERROR }
End-Section

# ─────────────────────────────────────────────────────────────
# 18. ENTRA AUTH POLICIES (HRD, Federation, Cross-Tenant)
# ─────────────────────────────────────────────────────────────
Start-Section "[18/22] Entra Auth Policies (HRD, Federation, Cross-Tenant)"
try {
    if (-not $GraphAvailable) { Write-Log "Graph unavailable — skipping Auth Policies." -Level WARN }
    else {
        $authResults = [System.Collections.Generic.List[object]]::new()

        # Domain federation configuration
        try {
            $fedResp = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/domains/$([Uri]::EscapeDataString($Domain))/federationConfiguration" -Method GET -OutputType PSObject -ErrorAction Stop
            foreach ($fc in @(Get-ObjProp $fedResp 'value')) {
                $authResults.Add([PSCustomObject]@{
                    PolicyType  = 'DomainFederation'
                    DisplayName = "Federation config for $Domain"
                    ObjectId    = Get-ObjProp $fc 'id'
                    State       = 'Federated'
                    Detail      = "IssuerUri: $(Get-ObjProp $fc 'issuerUri')"
                    Action      = 'Convert domain to Managed authentication'
                }) | Out-Null
            }
        } catch {
            if ($_.Exception.Message -notlike '*404*' -and $_.Exception.Message -notlike '*ResourceNotFound*') {
                Write-Log "Federation config check failed: $($_.Exception.Message.Split("`n")[0])" -Level WARN
            } else {
                Write-Log "Domain is not federated — no federation config to remediate."
            }
        }

        # Home Realm Discovery Policies
        try {
            $hrdPolicies = @(Invoke-GraphGetAll -Uri "https://graph.microsoft.com/v1.0/policies/homeRealmDiscoveryPolicies?`$select=id,displayName,definition,isOrganizationDefault")
            foreach ($hrd in $hrdPolicies) {
                $defJson = (@(Get-ObjProp $hrd 'definition') -join ' ')
                $dn      = Get-ObjProp $hrd 'displayName'
                if ($defJson -like "*$Domain*" -or $dn -like "*$DomainPrefix*") {
                    $authResults.Add([PSCustomObject]@{
                        PolicyType  = 'HomeRealmDiscovery'
                        DisplayName = $dn
                        ObjectId    = Get-ObjProp $hrd 'id'
                        State       = 'Active'
                        Detail      = 'Definition references domain'
                        Action      = 'Remove or update appliesTo assignment'
                    }) | Out-Null
                }
            }
        } catch { Write-Log "HRD policy scan failed: $($_.Exception.Message.Split("`n")[0])" -Level WARN }

        # Token Issuance Policies
        try {
            $tokenPolicies = @(Invoke-GraphGetAll -Uri "https://graph.microsoft.com/v1.0/policies/tokenIssuancePolicies?`$select=id,displayName,definition")
            foreach ($tp in $tokenPolicies) {
                $defJson = (@(Get-ObjProp $tp 'definition') -join ' ')
                if ($defJson -like "*$Domain*" -or (Get-ObjProp $tp 'displayName') -like "*$DomainPrefix*") {
                    $authResults.Add([PSCustomObject]@{
                        PolicyType  = 'TokenIssuance'
                        DisplayName = Get-ObjProp $tp 'displayName'
                        ObjectId    = Get-ObjProp $tp 'id'
                        State       = 'Active'
                        Detail      = 'Definition references domain'
                        Action      = 'Review and remove or update'
                    }) | Out-Null
                }
            }
        } catch { Write-Log "Token issuance policy scan failed: $($_.Exception.Message.Split("`n")[0])" -Level WARN }

        # Cross-Tenant Access Policy Partners
        try {
            $xtPartners = @(Invoke-GraphGetAll -Uri "https://graph.microsoft.com/v1.0/policies/crossTenantAccessPolicy/partners?`$select=tenantId,isServiceProvider")
            foreach ($p in $xtPartners) {
                $authResults.Add([PSCustomObject]@{
                    PolicyType  = 'CrossTenantAccess'
                    DisplayName = "Partner tenant: $(Get-ObjProp $p 'tenantId')"
                    ObjectId    = Get-ObjProp $p 'tenantId'
                    State       = 'Active'
                    Detail      = "Verify whether this partner relates to $Domain"
                    Action      = 'Verify and remove if linked to decommissioned domain/tenant'
                }) | Out-Null
            }
        } catch { Write-Log "Cross-tenant access policy scan failed: $($_.Exception.Message.Split("`n")[0])" -Level WARN }

        Write-Log "Found $($authResults.Count) auth policy reference(s)."
        Export-SafeCsv -Path (Join-Path $DiscoveryFolder '18_EntraAuthPolicies.csv') -Data @($authResults) -Label 'Entra Auth Policies'
    }
} catch { Write-Log "Error: $($_.Exception.Message.Split("`n")[0])" -Level ERROR }
End-Section


# ─────────────────────────────────────────────────────────────
# 19. ON-PREM AD USERS (Hybrid only)
# ─────────────────────────────────────────────────────────────
Start-Section "[19/22] On-Prem AD Users (Hybrid)"
try {
    if (-not $Hybrid) {
        Write-Log "Hybrid mode disabled — skipping (run with -Hybrid)." -Level WARN
    } elseif (-not $Script:HybridReady) {
        Write-Log "ActiveDirectory module not available — skipping." -Level WARN
    } else {
        $rows = [System.Collections.Generic.List[object]]::new()
        foreach ($key in $Script:ADCacheUPN.Keys) {
            $u = $Script:ADCacheUPN[$key]
            if ($u.Type -ne 'User') { continue }
            if (-not (Test-MatchesBusinessUnit $u.ExtensionAttribute7)) { continue }
            $rows.Add([PSCustomObject]@{
                DisplayName         = $u.DisplayName
                UserPrincipalName   = $u.UPN
                Mail                = $u.Mail
                sAMAccountName      = $u.sAMAccountName
                Enabled             = $u.Enabled
                ExtensionAttribute7 = $u.ExtensionAttribute7
                DistinguishedName   = $u.DN
                ProxyAddresses      = $u.ProxyAddresses
            }) | Out-Null
        }
        Write-Log "Found $($rows.Count) on-prem AD user(s) matching filters."
        Export-SafeCsv -Path (Join-Path $DiscoveryFolder '19_OnPrem_ADUsers.csv') -Data @($rows) -Label 'On-Prem AD Users'
    }
} catch { Write-Log "Error: $($_.Exception.Message.Split("`n")[0])" -Level ERROR }
End-Section

# ─────────────────────────────────────────────────────────────
# 20. ON-PREM AD GROUPS (Hybrid only)
# ─────────────────────────────────────────────────────────────
Start-Section "[20/22] On-Prem AD Groups (Hybrid)"
try {
    if (-not $Hybrid) {
        Write-Log "Hybrid mode disabled — skipping." -Level WARN
    } elseif (-not $Script:HybridReady) {
        Write-Log "ActiveDirectory module not available — skipping." -Level WARN
    } else {
        $seen = @{}
        $rows = [System.Collections.Generic.List[object]]::new()
        foreach ($entry in $Script:ADCacheMail.Values) {
            if ($entry.Type -ne 'Group') { continue }
            if ($seen.ContainsKey($entry.DN)) { continue }
            $seen[$entry.DN] = $true
            if (-not (Test-MatchesBusinessUnit $entry.ExtensionAttribute7)) { continue }
            $rows.Add([PSCustomObject]@{
                DisplayName         = $entry.DisplayName
                Mail                = $entry.Mail
                GroupCategory       = $entry.GroupCategory
                GroupScope          = $entry.GroupScope
                ExtensionAttribute7 = $entry.ExtensionAttribute7
                DistinguishedName   = $entry.DN
            }) | Out-Null
        }
        Write-Log "Found $($rows.Count) on-prem AD group(s) matching filters."
        Export-SafeCsv -Path (Join-Path $DiscoveryFolder '20_OnPrem_ADGroups.csv') -Data @($rows) -Label 'On-Prem AD Groups'
    }
} catch { Write-Log "Error: $($_.Exception.Message.Split("`n")[0])" -Level ERROR }
End-Section

# ─────────────────────────────────────────────────────────────
# 21. ON-PREM AD CONTACTS (Hybrid only)
# ─────────────────────────────────────────────────────────────
Start-Section "[21/22] On-Prem AD Contacts (Hybrid)"
try {
    if (-not $Hybrid) {
        Write-Log "Hybrid mode disabled — skipping." -Level WARN
    } elseif (-not $Script:HybridReady) {
        Write-Log "ActiveDirectory module not available — skipping." -Level WARN
    } else {
        $contactProps = @('mail','proxyAddresses','extensionAttribute7','DistinguishedName','DisplayName','targetAddress')
        $contacts = @(Get-ADObject -LDAPFilter "(&(objectClass=contact)(|(mail=*$Domain*)(proxyAddresses=*$Domain*)(targetAddress=*$Domain*)))" -Properties $contactProps -ErrorAction SilentlyContinue)
        $rows = [System.Collections.Generic.List[object]]::new()
        foreach ($c in $contacts) {
            $ext7 = $c.extensionAttribute7
            if (-not (Test-MatchesBusinessUnit $ext7)) { continue }
            $rows.Add([PSCustomObject]@{
                DisplayName         = $c.DisplayName
                Mail                = $c.mail
                TargetAddress       = $c.targetAddress
                ExtensionAttribute7 = $ext7
                DistinguishedName   = $c.DistinguishedName
                ProxyAddresses      = (@($c.proxyAddresses) -join '; ')
            }) | Out-Null
        }
        Write-Log "Found $($rows.Count) on-prem AD contact(s) matching filters."
        Export-SafeCsv -Path (Join-Path $DiscoveryFolder '21_OnPrem_ADContacts.csv') -Data @($rows) -Label 'On-Prem AD Contacts'
    }
} catch { Write-Log "Error: $($_.Exception.Message.Split("`n")[0])" -Level ERROR }
End-Section


# ─────────────────────────────────────────────────────────────
# 22. USER LICENSES (Graph)
# Enumerates assigned licenses for every user matched in section 2 via
# Graph /users/{upn}/licenseDetails, then writes two CSVs:
#   22_UserLicenses.csv   — one row per user with their assigned license list
#   22a_LicenseCounts.csv — aggregate count per SKU (sorted desc)
# Falls back to a Graph endsWith(UPN) enumeration if section 2 produced no rows.
# Uses UPN for the Graph lookup (works in both Hybrid and Cloud-only modes — the
# Hybrid branch of section 2 stores the on-prem ObjectGUID which is NOT the same
# as the Entra object id, so UPN is the safer key).
# ─────────────────────────────────────────────────────────────
Start-Section "[22/22] User Licenses (Graph)"
try {
    if (-not $GraphAvailable) {
        Write-Log "Graph unavailable — skipping license enumeration." -Level WARN
        Export-SafeCsv -Path (Join-Path $DiscoveryFolder '22_UserLicenses.csv')   -Data @() -Label 'User Licenses'
        Export-SafeCsv -Path (Join-Path $DiscoveryFolder '22a_LicenseCounts.csv') -Data @() -Label 'License Counts'
    } else {
        # ─── Build tenant SKU lookup (skuId → skuPartNumber + friendly name) ──
        Write-Log "Fetching tenant subscribedSkus to resolve SKU IDs..."
        $skuMap = @{}
        try {
            $skus = @(Invoke-GraphGetAll -Uri 'https://graph.microsoft.com/v1.0/subscribedSkus')
            foreach ($s in $skus) {
                $sid  = [string](Get-ObjProp $s 'skuId')
                $part = [string](Get-ObjProp $s 'skuPartNumber')
                if ($sid) {
                    $skuMap[$sid] = [PSCustomObject]@{
                        SkuPartNumber = $part
                        FriendlyName  = (Get-LicenseFriendlyName $part)
                    }
                }
            }
            Write-Log "Resolved $($skuMap.Count) tenant SKU(s) from subscribedSkus."
        } catch {
            Write-Log "subscribedSkus fetch failed (will fall back to per-licenseDetails values): $($_.Exception.Message.Split("`n")[0])" -Level WARN
        }

        # ─── Determine the user set to query ──────────────────────────────────
        # Prefer the already-filtered $matchedUsers from section 2 (honours -BusinessUnitId).
        # Fall back to a Graph UPN enumeration if section 2 produced nothing.
        $userSet = [System.Collections.Generic.List[object]]::new()
        if ($matchedUsers -and $matchedUsers.Count -gt 0) {
            Write-Log "Using $($matchedUsers.Count) user(s) from section 2."
            foreach ($u in $matchedUsers) {
                if ($u.UserPrincipalName) {
                    $userSet.Add([PSCustomObject]@{
                        UPN         = $u.UserPrincipalName
                        DisplayName = $u.DisplayName
                    }) | Out-Null
                }
            }
        } else {
            Write-Log "Section 2 produced no users — falling back to all members in connected tenant..."
            try {
                $graphUserUri = "https://graph.microsoft.com/v1.0/users?`$filter=userType eq 'Member'&`$select=id,userPrincipalName,displayName"
                $fallbackUsers = @(Invoke-GraphGetAll -Uri $graphUserUri)
                foreach ($g in $fallbackUsers) {
                    $upn = (Get-ObjProp $g 'userPrincipalName')
                    if ($upn) {
                        $userSet.Add([PSCustomObject]@{
                            UPN         = $upn
                            DisplayName = (Get-ObjProp $g 'displayName')
                        }) | Out-Null
                    }
                }
                Write-Log "Graph fallback returned $($userSet.Count) user(s)."
            } catch {
                Write-Log "Graph fallback enumeration failed: $($_.Exception.Message.Split("`n")[0])" -Level WARN
            }
        }

        if ($userSet.Count -eq 0) {
            Write-Log "No users to query — writing placeholder CSVs." -Level WARN
            Export-SafeCsv -Path (Join-Path $DiscoveryFolder '22_UserLicenses.csv')   -Data @() -Label 'User Licenses'
            Export-SafeCsv -Path (Join-Path $DiscoveryFolder '22a_LicenseCounts.csv') -Data @() -Label 'License Counts'
        } else {
            # ─── Parallel per-user licenseDetails lookup ──────────────────────
            Write-Log "Querying license details for $($userSet.Count) user(s) in parallel (throttle = 10)..."
            $licStart = Get-Date

            $perUserResults = $userSet | ForEach-Object -ThrottleLimit 10 -Parallel {
                $u = $_
                if ([string]::IsNullOrWhiteSpace($u.UPN)) { return $null }
                try {
                    # Graph context is inherited per-runspace in PS7 parallel; if missing, the call below will throw.
                    $upnEnc = [System.Uri]::EscapeDataString($u.UPN)
                    $uri    = "https://graph.microsoft.com/v1.0/users/$upnEnc/licenseDetails"
                    $resp   = Invoke-MgGraphRequest -Uri $uri -Method GET -OutputType PSObject -ErrorAction Stop
                    $vals   = @()
                    if ($resp.PSObject.Properties['value']) { $vals = @($resp.value) }
                    [PSCustomObject]@{
                        UPN         = $u.UPN
                        DisplayName = $u.DisplayName
                        Licenses    = $vals
                        Error       = $null
                    }
                } catch {
                    [PSCustomObject]@{
                        UPN         = $u.UPN
                        DisplayName = $u.DisplayName
                        Licenses    = @()
                        Error       = $_.Exception.Message.Split([Environment]::NewLine)[0]
                    }
                }
            }
            $licElapsed = (Get-Date) - $licStart
            Write-Log ("Parallel license lookup completed in {0:mm\:ss}." -f $licElapsed)

            # ─── Build per-user rows + aggregate counts ───────────────────────
            $licenseRows  = [System.Collections.Generic.List[object]]::new()
            $licenseCount = @{}   # skuPartNumber → count
            $licenseFriendlyByPart = @{}   # skuPartNumber → friendly name (cached for the count CSV)
            $usersWithLicenses = 0
            $usersWithoutLicenses = 0
            $errorCount = 0

            foreach ($r in $perUserResults) {
                if (-not $r) { continue }
                if ($r.Error) { $errorCount++ }

                $skuParts    = [System.Collections.Generic.List[string]]::new()
                $skuFriendly = [System.Collections.Generic.List[string]]::new()
                $skuIds      = [System.Collections.Generic.List[string]]::new()

                foreach ($lic in $r.Licenses) {
                    $sid  = [string](Get-ObjProp $lic 'skuId')
                    $part = [string](Get-ObjProp $lic 'skuPartNumber')
                    # licenseDetails returns skuPartNumber directly, but fall back to skuMap if needed.
                    if ([string]::IsNullOrWhiteSpace($part) -and $sid -and $skuMap.ContainsKey($sid)) {
                        $part = $skuMap[$sid].SkuPartNumber
                    }
                    if ([string]::IsNullOrWhiteSpace($part)) { continue }

                    $friendly = if ($sid -and $skuMap.ContainsKey($sid)) {
                        $skuMap[$sid].FriendlyName
                    } else {
                        Get-LicenseFriendlyName $part
                    }

                    $skuParts.Add($part)        | Out-Null
                    $skuFriendly.Add($friendly) | Out-Null
                    if ($sid) { $skuIds.Add($sid) | Out-Null }

                    if ($licenseCount.ContainsKey($part)) { $licenseCount[$part]++ } else { $licenseCount[$part] = 1 }
                    if (-not $licenseFriendlyByPart.ContainsKey($part)) { $licenseFriendlyByPart[$part] = $friendly }
                }

                if ($skuParts.Count -gt 0) { $usersWithLicenses++ } else { $usersWithoutLicenses++ }

                $licenseRows.Add([PSCustomObject]@{
                    DisplayName       = $r.DisplayName
                    UserPrincipalName = $r.UPN
                    LicenseCount      = $skuParts.Count
                    SkuPartNumbers    = ($skuParts    -join '; ')
                    LicenseNames      = ($skuFriendly -join '; ')
                    SkuIds            = ($skuIds      -join '; ')
                    LookupError       = $r.Error
                }) | Out-Null
            }

            # ─── Write per-user CSV ───────────────────────────────────────────
            Export-SafeCsv -Path (Join-Path $DiscoveryFolder '22_UserLicenses.csv') -Data @($licenseRows) -Label 'User Licenses'

            # ─── Build + write aggregate counts CSV (sorted by count desc) ────
            $countRows = [System.Collections.Generic.List[object]]::new()
            foreach ($k in $licenseCount.Keys) {
                # NOTE: assign friendly name to a variable BEFORE the hashtable
                # literal — using an inline `if ... else { Get-LicenseFriendlyName -SkuPartNumber $k }`
                # as a hashtable value tripped PowerShell's parser ("A positional
                # parameter cannot be found that accepts argument 'SkuPartNumber'").
                if ($licenseFriendlyByPart.ContainsKey($k)) {
                    $friendlyForCount = $licenseFriendlyByPart[$k]
                } else {
                    $friendlyForCount = Get-LicenseFriendlyName $k
                }
                $countRows.Add([PSCustomObject]@{
                    SkuPartNumber = $k
                    LicenseName   = $friendlyForCount
                    AssignedCount = $licenseCount[$k]
                }) | Out-Null
            }
            # ─── Sort by AssignedCount desc, with SkuPartNumber as tie-breaker.
            # NOTE: `Sort-Object AssignedCount -Descending, SkuPartNumber` does NOT
            # work — PowerShell parses the comma as the array operator, leaving
            # SkuPartNumber as an unbound positional arg ("A positional parameter
            # cannot be found that accepts argument 'SkuPartNumber'"). Use the
            # hashtable-property syntax to specify per-property sort direction.
            $countRowsSorted = @($countRows | Sort-Object -Property `
                @{Expression='AssignedCount'; Descending=$true}, `
                @{Expression='SkuPartNumber'; Descending=$false})
            Export-SafeCsv -Path (Join-Path $DiscoveryFolder '22a_LicenseCounts.csv') -Data $countRowsSorted -Label 'License Counts'

            Write-Log ("Licenses — users with licenses: {0} | users without: {1} | distinct SKUs: {2} | lookup errors: {3}" -f $usersWithLicenses, $usersWithoutLicenses, $licenseCount.Count, $errorCount)
        }
    }
} catch {
    # Verbose error reporting so any remaining bugs in this section surface with
    # a usable line/column/command for diagnostics. Without this, the catch
    # collapses the failure to a single line of text and the location is lost.
    $errLine = $_.InvocationInfo.ScriptLineNumber
    $errCol  = $_.InvocationInfo.OffsetInLine
    $errCmd  = ($_.InvocationInfo.Line -replace '\s+',' ').Trim()
    Write-Log "Error: $($_.Exception.Message.Split("`n")[0])" -Level ERROR
    Write-Log "  at line ${errLine}:${errCol} — $errCmd" -Level ERROR
    Write-Detail "Full exception: $($_.Exception.GetType().FullName) | $($_.Exception.Message)"
    if ($_.ScriptStackTrace) { Write-Detail "Stack: $($_.ScriptStackTrace -replace [Environment]::NewLine,' || ')" }
}
End-Section


# ─────────────────────────────────────────────────────────────
# SUMMARY + EXCEL MERGE
# ─────────────────────────────────────────────────────────────
Write-Log "Generating summary report..."
try {
    $summaryPath = Join-Path $DiscoveryFolder '_Summary.csv'
    $summaryData = @(Get-ChildItem -Path $DiscoveryFolder -Filter '*.csv' -File |
        Where-Object { $_.Name -ne '_Summary.csv' } |
        Sort-Object Name |
        ForEach-Object {
            $rows = Get-CsvRowCountFast -Path $_.FullName
            [PSCustomObject]@{
                Area        = ($_.BaseName -replace '^\d+\w?_','')
                File        = $_.Name
                RecordCount = $rows
            }
        })
    $summaryData | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8 -Force
} catch { Write-Log "Error generating summary: $_" -Level ERROR }

Write-Log "Merging CSV files into a single Excel workbook..."
try {
    $excelPath = Join-Path $OutputFolder "1. $DomainPrefix Discovery Objects.xlsx"
    Merge-CsvFolderToExcel -FolderPath $DiscoveryFolder -ExcelPath $excelPath -IncludeSummary
} catch { Write-Log "Excel merge failed: $_" -Level WARN }

# ─────────────────────────────────────────────────────────────
# DISCONNECT
# ─────────────────────────────────────────────────────────────
# NOTE: Disconnect-ExchangeOnline triggers MSAL.ClearAllTokensAsync() on a background
# thread, which on hosts with the broken WAM RuntimeBroker raises an unhandled
# NullReferenceException that terminates the PowerShell process. Skipping the explicit
# disconnect avoids this — the connection is reaped automatically when PowerShell exits,
# and keeping it alive between domains in batch mode is desirable for token re-use anyway.
Write-Log "Discovery complete; leaving sessions open (cleaned up on PowerShell exit)."

$totalElapsed = (Get-Date) - $Script:ScriptStart
Write-Log ("Script complete in {0:hh\:mm\:ss}. Output: $OutputFolder" -f $totalElapsed) -Level SUCCESS