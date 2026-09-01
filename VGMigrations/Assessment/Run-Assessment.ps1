#Requires -Version 7.0
<#
.SYNOPSIS
    M365 Tenant Assessment - modular replacement for search-domain.ps1.
.DESCRIPTION
    Collects AD, Exchange Online, Microsoft Graph, SharePoint, and Power Platform data for a
    VBU, writes one JSON file per collector plus an assessment workbook, and (always) a
    Domain Removal-compatible Discovery\ CSV folder so the existing removal scripts keep
    working unchanged. Also offers a "Generate Mapping Files" mode that turns a finished
    workbook into the six AvePoint Fly import files.

    Every prompt below falls back to Read-Host only when its value wasn't supplied as a
    parameter - this lets the standalone WinForms launcher (discovery-menu.ps1) and the
    Electron app's streamed Discovery panel drive a full run non-interactively (the latter
    has no stdin at all - a bare Read-Host would hang forever).
.PARAMETER Domain
    VBU domain (e.g. contoso.com). Supplying this switches every other prompt below to
    non-interactive mode (skips default to "proceed", not "ask").
.PARAMETER VBUSearchTerm
    Company name fragment used to scope AD/Graph/SharePoint matches (e.g. "Contoso").
    Defaults to the domain's first label when omitted in non-interactive mode.
.PARAMETER VBUId
    extensionAttribute7 value used to scope AD-sourced sections. Optional.
.PARAMETER SharePointAdminUrl
    SPO admin URL, e.g. https://tenant-admin.sharepoint.com. Defaults to
    https://ourvolaris-admin.sharepoint.com when omitted.
.PARAMETER SkipGraph
    Skips Microsoft Graph collection (Teams, M365 Groups, Intune devices, app registrations).
.PARAMETER SkipPowerPlatform
    Skips the Power Platform child-process scan (its own interactive sign-in).
.PARAMETER OutputPath
    Base directory for the assessment output folder. Defaults to the script's own folder.
.PARAMETER DeleteRawJson
    Deletes the Raw\*.json files after the workbook and compatibility CSVs are written.
    Defaults to keeping them when running non-interactively.
.PARAMETER KeepSession
    Skips signing out of SharePoint/Exchange/Graph at the end of the run, and skips signing
    back in if this process is already connected. Used by Run-MultiAssessment.ps1 so a batch
    of domains against the same source tenant only prompts for sign-in once instead of once
    per domain - Power Platform is unaffected (its own child-process sign-in still runs per
    domain regardless, since that scan itself is domain-scoped).
#>

[CmdletBinding()]
param(
    [string]$Domain,
    [string]$VBUSearchTerm,
    [string]$VBUId,
    [string]$SharePointAdminUrl,
    [switch]$SkipGraph,
    [switch]$SkipPowerPlatform,
    [string]$OutputPath,
    [switch]$DeleteRawJson,
    [switch]$KeepSession
)

# Any of these being supplied means we were launched by the GUI/Electron app, not typed by
# hand at a console - every remaining prompt below must not block on stdin.
$script:Unattended = [bool]$PSBoundParameters.ContainsKey('Domain')

# Define module root before any imports
$moduleRoot = Join-Path $PSScriptRoot 'Modules'

# Common.psm1 first - prefix constants and console helpers needed for the banner and menu
Import-Module (Join-Path $moduleRoot 'Common.psm1') -Force -DisableNameChecking -Global

# -----------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------
function Get-SafeCount {
    param([hashtable]$Counts, [string]$Key)
    if ($null -ne $Counts -and $Counts.ContainsKey($Key)) { return $Counts[$Key] }
    return 0
}

function Show-ModeMenu {
    $options  = @('Run Assessment', 'Generate Mapping Files')
    $selected = 0

    Write-Host 'Select mode (arrow keys, Enter to confirm):' -ForegroundColor Cyan
    while ($true) {
        for ($i = 0; $i -lt $options.Count; $i++) {
            if ($i -eq $selected) {
                Write-Host ('> ' + $options[$i]) -ForegroundColor Green
            }
            else {
                Write-Host ('  ' + $options[$i]) -ForegroundColor DarkGray
            }
        }

        $key = [System.Console]::ReadKey($true)
        switch ($key.Key) {
            'UpArrow'   { if ($selected -gt 0) { $selected-- } }
            'DownArrow' { if ($selected -lt $options.Count - 1) { $selected++ } }
            'Enter'     { return $options[$selected] }
        }

        # Redraw in place - move cursor back to the first option line
        [System.Console]::SetCursorPosition(0, [System.Console]::CursorTop - $options.Count)
    }
}

# -----------------------------------------------------------------------
# Banner
# -----------------------------------------------------------------------
Write-Host ''
Write-Host 'M365 Tenant Assessment' -ForegroundColor Cyan
Write-Host ('=' * 40) -ForegroundColor Cyan
Write-Host ''

# -----------------------------------------------------------------------
# Mode selection
# -----------------------------------------------------------------------
$mode = if ($script:Unattended) { 'Run Assessment' } else { Show-ModeMenu }
Write-Host ''

if ($mode -eq 'Generate Mapping Files') {
    # Mapping mode reads only the workbook - just ImportExcel and the MappingFiles module
    Import-Module ImportExcel -DisableNameChecking -ErrorAction Stop
    Import-Module (Join-Path $moduleRoot 'MappingFiles.psm1') -Force -DisableNameChecking -Global
    Invoke-MappingFileGeneration
    return
}

# -----------------------------------------------------------------------
# Service modules
# -----------------------------------------------------------------------
# ActiveDirectory is imported lazily inside AD.psm1's Invoke-ADCollection, guarded by
# Context.SkipAD - importing it here would hard-fail the whole run on a machine without RSAT,
# which is exactly the cloud-only scenario Prerequisites.psm1's soft-fail exists to support.

Import-Module ExchangeOnlineManagement -DisableNameChecking -ErrorAction Stop

# Pins Microsoft.Graph.* to 2.33.0 - the last SDK version before WAM became mandatory for
# interactive sign-in (see Ensure-GraphModules.ps1's own header for the full history). This is
# a proven fix carried over from search-domain.ps1; skipping it risks rediscovering the same
# WAM/Conditional-Access failures live against a customer tenant.
. (Join-Path $PSScriptRoot '..\Ensure-GraphModules.ps1') -GraphModules @(
    'Microsoft.Graph.Groups', 'Microsoft.Graph.Teams', 'Microsoft.Graph.Users', 'Microsoft.Graph.DeviceManagement'
)
Import-Module Microsoft.Graph.Groups           -DisableNameChecking -ErrorAction Stop
Import-Module Microsoft.Graph.Teams            -DisableNameChecking -ErrorAction Stop
Import-Module Microsoft.Graph.Users            -DisableNameChecking -ErrorAction Stop
Import-Module Microsoft.Graph.DeviceManagement -DisableNameChecking -ErrorAction Stop
Import-Module ImportExcel                      -DisableNameChecking -ErrorAction Stop

# SharePoint - soft-fail (installable by Test-Prerequisites; missing sets SkipSharePoint)
try {
    $prev = $WarningPreference
    $WarningPreference = 'SilentlyContinue'
    Import-Module Microsoft.Online.SharePoint.PowerShell `
        -UseWindowsPowerShell -DisableNameChecking -ErrorAction Stop
    $WarningPreference = $prev
}
catch {
    $WarningPreference = $prev
    Write-Host ($PREFIX_WARN + 'SPO module not available - SharePoint collection will be skipped') -ForegroundColor Yellow
}

# -----------------------------------------------------------------------
# Remaining platform modules
# -----------------------------------------------------------------------
Import-Module (Join-Path $moduleRoot 'Prerequisites.psm1')  -Force -DisableNameChecking -Global
Import-Module (Join-Path $moduleRoot 'AD.psm1')              -Force -DisableNameChecking -Global
Import-Module (Join-Path $moduleRoot 'Exchange.psm1')        -Force -DisableNameChecking -Global
Import-Module (Join-Path $moduleRoot 'Graph.psm1')           -Force -DisableNameChecking -Global
Import-Module (Join-Path $moduleRoot 'SharePoint.psm1')      -Force -DisableNameChecking -Global
Import-Module (Join-Path $moduleRoot 'PowerPlatform.psm1')   -Force -DisableNameChecking -Global
Import-Module (Join-Path $moduleRoot 'Workbook.psm1')        -Force -DisableNameChecking -Global
Import-Module (Join-Path $moduleRoot 'LegacyExport.psm1')    -Force -DisableNameChecking -Global

# -----------------------------------------------------------------------
# User inputs
# -----------------------------------------------------------------------
$vbuDomain     = if ($PSBoundParameters.ContainsKey('Domain'))        { $Domain.Trim() }              else { (Read-Host 'VBU Domain       (e.g. contoso.com)').Trim() }
$vbuSearchTerm = if ($PSBoundParameters.ContainsKey('VBUSearchTerm')) { $VBUSearchTerm.Trim() }
                 elseif ($script:Unattended)                          { ($vbuDomain -split '\.')[0] }
                 else                                                 { (Read-Host 'VBU Search Term  (e.g. Contoso)').Trim() }
$vbuId         = if ($PSBoundParameters.ContainsKey('VBUId'))         { $VBUId.Trim() }
                 elseif ($script:Unattended)                          { '' }
                 else                                                 { (Read-Host 'VBU ID           (extensionAttribute7 exact value)').Trim() }
$spoInput      = if ($PSBoundParameters.ContainsKey('SharePointAdminUrl')) { $SharePointAdminUrl.Trim() }
                 elseif ($script:Unattended)                               { '' }
                 else                                                      { (Read-Host 'SPO Admin URL    [https://ourvolaris-admin.sharepoint.com]').Trim() }
$spoAdminUrl   = if ($spoInput) { $spoInput } else { 'https://ourvolaris-admin.sharepoint.com' }

# -----------------------------------------------------------------------
# Folder structure
# -----------------------------------------------------------------------
# Temporary context call to derive VBUName via the shared TLD-strip logic in Common.psm1
$vbuName      = (New-AssessmentContext -VBUDomain $vbuDomain -VBUId $vbuId -VBUSearchTerm $vbuSearchTerm -RawPath 'TEMP').VBUName
$timestamp    = Get-Date -Format 'yyyyMMdd-HHmm'
$baseFolder   = if ($OutputPath) { $OutputPath } else { $PSScriptRoot }
$assessFolder = Join-Path $baseFolder "$vbuName-$timestamp"
$rawPath      = Join-Path $assessFolder 'Raw'
$discoveryPath = Join-Path $assessFolder 'Discovery'
$xlsxPath     = Join-Path $assessFolder "$vbuName-Assessment.xlsx"

New-Item -ItemType Directory -Path $rawPath                                 -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $assessFolder 'MappingFiles') -Force | Out-Null

Write-Host ($PREFIX_OK + "Output folder: $assessFolder") -ForegroundColor Green

# -----------------------------------------------------------------------
# Assessment context
# -----------------------------------------------------------------------
$ctx = New-AssessmentContext `
    -VBUDomain     $vbuDomain `
    -VBUId         $vbuId `
    -VBUSearchTerm $vbuSearchTerm `
    -RawPath       $rawPath `
    -SPOAdminUrl   $spoAdminUrl

if (-not (Get-Module -ListAvailable -Name 'Microsoft.Online.SharePoint.PowerShell')) {
    $ctx.SkipSharePoint = $true
}

# -----------------------------------------------------------------------
# Prerequisites
# -----------------------------------------------------------------------
Test-Prerequisites -Context $ctx

# -----------------------------------------------------------------------
# Authentication - all sessions connected upfront
# -----------------------------------------------------------------------
Write-SectionHeader 'Authentication'

# SPO must connect first - Exchange and Graph WAM tokens conflict with SPO auth if they connect before it
# No native "am I connected" query exists for Connect-SPOService, so KeepSession batches track
# it via a process-wide flag instead - reliable since the underlying session really does live at
# the process level (Run-MultiAssessment.ps1 re-invokes this script with '&' in the same process).
if (-not $ctx.SkipSharePoint) {
    if ($KeepSession -and $global:AssessmentSpoConnected) {
        Write-Host ($PREFIX_OK + 'SharePoint Online already connected (reusing session)') -ForegroundColor Green
    }
    else {
        Write-Host ($PREFIX_INFO + 'Connecting to SharePoint Online...') -ForegroundColor DarkGray
        try {
            Connect-SPOService -Url $ctx.SPOAdminUrl -ErrorAction Stop
            Write-Host ($PREFIX_OK + 'SharePoint Online connected') -ForegroundColor Green
            $global:AssessmentSpoConnected = $true
        }
        catch {
            Write-Host ($PREFIX_FAIL + 'SPO connection failed: ' + $_.Exception.Message) -ForegroundColor Red
            Write-Host ($PREFIX_WARN + 'SharePoint collection will be skipped') -ForegroundColor Yellow
            $ctx.SkipSharePoint = $true
        }
    }
}

if ($KeepSession -and (Get-ConnectionInformation | Where-Object { $_.State -eq 'Connected' })) {
    Write-Host ($PREFIX_OK + 'Exchange Online already connected (reusing session)') -ForegroundColor Green
}
else {
    Write-Host ($PREFIX_INFO + 'Connecting to Exchange Online...') -ForegroundColor DarkGray
    Connect-ExchangeOnline -ShowBanner:$false -DisableWAM -ErrorAction Stop
    Write-Host ($PREFIX_OK + 'Exchange Online connected') -ForegroundColor Green
}

$skipGraphSession = [bool]$SkipGraph
$graphStart = Get-Date
Write-Host ''
if (-not $skipGraphSession -and -not $script:Unattended) {
    $graphInput = (Read-Host 'Ready for Microsoft Graph authentication? [Enter to proceed, S to skip]').Trim()
    if ($graphInput.ToUpper() -eq 'S') { $skipGraphSession = $true }
}
if ($skipGraphSession) {
    Write-Host ($PREFIX_SKIP + 'Graph collection will be skipped') -ForegroundColor Yellow
}
elseif ($KeepSession -and (Get-MgContext)) {
    Write-Host ($PREFIX_OK + 'Microsoft Graph already connected (reusing session)') -ForegroundColor Green
}
else {
    Write-Host ($PREFIX_INFO + 'Connecting to Microsoft Graph...') -ForegroundColor DarkGray
    try {
        Connect-MgGraph -Scopes @(
            'Group.Read.All', 'Directory.Read.All', 'User.Read.All', 'Device.Read.All',
            'Sites.Read.All', 'Team.ReadBasic.All', 'TeamMember.Read.All',
            'DeviceManagementManagedDevices.Read.All', 'Policy.Read.All', 'Application.Read.All'
        ) -NoWelcome -ErrorAction Stop
        Write-Host ($PREFIX_OK + 'Microsoft Graph connected') -ForegroundColor Green
    }
    catch {
        Write-Host ($PREFIX_FAIL + 'Graph connection failed: ' + $_.Exception.Message) -ForegroundColor Red
        Write-Host ($PREFIX_WARN + 'Graph collection will be skipped') -ForegroundColor Yellow
        $skipGraphSession = $true
    }
}

$skipPowerPlatformSession = [bool]$SkipPowerPlatform
if (-not $skipPowerPlatformSession -and -not $script:Unattended) {
    Write-Host ''
    $ppInput = (Read-Host 'Ready for Power Platform authentication (separate sign-in window)? [Enter to proceed, S to skip]').Trim()
    if ($ppInput.ToUpper() -eq 'S') { $skipPowerPlatformSession = $true }
}
$ctx.SkipPowerPlatform = $skipPowerPlatformSession

# -----------------------------------------------------------------------
# Collection and workbook
# -----------------------------------------------------------------------
$adResult  = $null
$exResult  = $null
$grResult  = $null
$spoResult = $null
$ppResult  = $null

try {
    # Active Directory - non-critical (soft-fail; see Prerequisites.psm1/AD.psm1)
    try {
        $adResult = Invoke-ADCollection -Context $ctx
    }
    catch {
        Write-Host ($PREFIX_FAIL + 'AD collection error: ' + $_.Exception.Message) -ForegroundColor Red
        $adResult = New-CollectorResult -Success $false -ErrorMessage $_.Exception.Message
    }

    # Exchange Online - non-critical
    try {
        $exResult = Invoke-ExchangeCollection -Context $ctx
    }
    catch {
        Write-Host ($PREFIX_FAIL + 'Exchange collection error: ' + $_.Exception.Message) -ForegroundColor Red
        $exResult = New-CollectorResult -Success $false -ErrorMessage $_.Exception.Message
    }

    # Microsoft Graph - non-critical, user-skippable
    if ($skipGraphSession) {
        Update-CollectorStatus -CollectorName 'Graph Data' -Status 'Skipped' `
            -RawPath $ctx.RawPath -StartTime $graphStart
        $grResult = New-CollectorResult -Skipped $true
    }
    else {
        try {
            $grResult = Invoke-GraphCollection -Context $ctx
        }
        catch {
            Write-Host ($PREFIX_FAIL + 'Graph collection error: ' + $_.Exception.Message) -ForegroundColor Red
            $grResult = New-CollectorResult -Success $false -ErrorMessage $_.Exception.Message
        }
    }

    # SharePoint Online - non-critical
    try {
        $spoResult = Invoke-SharePointCollection -Context $ctx
    }
    catch {
        Write-Host ($PREFIX_FAIL + 'SharePoint collection error: ' + $_.Exception.Message) -ForegroundColor Red
        $spoResult = New-CollectorResult -Success $false -ErrorMessage $_.Exception.Message
    }

    # Power Platform - non-critical, its own child-process sign-in
    try {
        $ppResult = Invoke-PowerPlatformCollection -Context $ctx -SkipPowerPlatform:$skipPowerPlatformSession
    }
    catch {
        Write-Host ($PREFIX_FAIL + 'Power Platform collection error: ' + $_.Exception.Message) -ForegroundColor Red
        $ppResult = New-CollectorResult -Success $false -ErrorMessage $_.Exception.Message
    }

    # Workbook - non-critical (always runs with whatever JSON was written)
    try {
        Invoke-WorkbookGeneration -Context $ctx -XlsxPath $xlsxPath
    }
    catch {
        Write-Host ($PREFIX_FAIL + 'Workbook generation error: ' + $_.Exception.Message) -ForegroundColor Red
    }

    # Domain Removal compatibility CSVs - always runs, even on a partial/degraded assessment,
    # so the Discovery folder those scripts expect exists after every run.
    try {
        Export-LegacyDiscoveryCsvs -Context $ctx -DiscoveryFolder $discoveryPath
    }
    catch {
        Write-Host ($PREFIX_FAIL + 'Compatibility CSV export error: ' + $_.Exception.Message) -ForegroundColor Red
    }

    # Final summary
    $adCounts  = if ($adResult)  { $adResult.Counts }  else { $null }
    $exCounts  = if ($exResult)  { $exResult.Counts }  else { $null }
    $grCounts  = if ($grResult)  { $grResult.Counts }  else { $null }
    $spoCounts = if ($spoResult) { $spoResult.Counts } else { $null }
    $ppCounts  = if ($ppResult)  { $ppResult.Counts }  else { $null }

    Write-SectionHeader 'Run Complete'
    Write-ProgressLine -Label 'AD Users'                     -Count (Get-SafeCount -Counts $adCounts  -Key 'UserCount')
    Write-ProgressLine -Label 'AD Groups'                    -Count (Get-SafeCount -Counts $adCounts  -Key 'GroupCount')
    Write-ProgressLine -Label 'AD Group Memberships'         -Count (Get-SafeCount -Counts $adCounts  -Key 'MembershipCount')
    Write-ProgressLine -Label 'AD Devices'                   -Count (Get-SafeCount -Counts $adCounts  -Key 'DeviceCount')
    Write-ProgressLine -Label 'User Mailboxes'               -Count (Get-SafeCount -Counts $exCounts  -Key 'UserMailboxCount')
    Write-ProgressLine -Label 'Shared Mailboxes'             -Count (Get-SafeCount -Counts $exCounts  -Key 'SharedMailboxCount')
    Write-ProgressLine -Label 'Resource Mailboxes'           -Count (Get-SafeCount -Counts $exCounts  -Key 'ResourceMailboxCount')
    Write-ProgressLine -Label 'Distribution Groups'          -Count (Get-SafeCount -Counts $exCounts  -Key 'DGCount')
    Write-ProgressLine -Label 'Mail-Enabled Security Groups' -Count (Get-SafeCount -Counts $exCounts  -Key 'MailEnabledSecurityGroupCount')
    Write-ProgressLine -Label 'Teams'                        -Count (Get-SafeCount -Counts $grCounts  -Key 'TeamCount')
    Write-ProgressLine -Label 'M365 Groups'                  -Count (Get-SafeCount -Counts $grCounts  -Key 'M365GroupCount')
    Write-ProgressLine -Label 'Intune Devices'               -Count (Get-SafeCount -Counts $grCounts  -Key 'IntuneDeviceCount')
    Write-ProgressLine -Label 'SharePoint Sites'             -Count (Get-SafeCount -Counts $spoCounts -Key 'SPOSiteCount')
    Write-ProgressLine -Label 'OneDrives'                    -Count (Get-SafeCount -Counts $spoCounts -Key 'OneDriveCount')
    Write-ProgressLine -Label 'Power Apps'                   -Count (Get-SafeCount -Counts $ppCounts  -Key 'PowerAppCount')
    Write-ProgressLine -Label 'Power Automate Flows'         -Count (Get-SafeCount -Counts $ppCounts  -Key 'FlowCount')
    Write-Host ''
    Write-Host ($PREFIX_OK + 'Workbook: ' + $xlsxPath) -ForegroundColor Green
    Write-Host ($PREFIX_OK + 'Domain Removal CSVs: ' + $discoveryPath) -ForegroundColor Green
}
catch {
    Write-Host ($PREFIX_FAIL + 'Assessment stopped: ' + $_.Exception.Message) -ForegroundColor Red
}
finally {
    if (-not $KeepSession) {
        try { Disconnect-MgGraph              -ErrorAction SilentlyContinue } catch {}
        try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
        try { Disconnect-SPOService            -ErrorAction SilentlyContinue } catch {}
        $global:AssessmentSpoConnected = $false
    }
}

# -----------------------------------------------------------------------
# Raw JSON cleanup - always runs regardless of assessment outcome
# -----------------------------------------------------------------------
Write-Host ''
$deleteRaw = if ($PSBoundParameters.ContainsKey('DeleteRawJson')) { [bool]$DeleteRawJson }
             elseif ($script:Unattended)                          { $false }
             else { (Read-Host "Delete Raw JSON files from '$rawPath'? [Y/N]").Trim().ToUpper() -eq 'Y' }
if ($deleteRaw) {
    Get-ChildItem -Path $rawPath -Filter '*.json' | Remove-Item -Force
    Write-Host ($PREFIX_OK + 'Raw JSON files deleted') -ForegroundColor Green
}
else {
    Write-Host ($PREFIX_INFO + 'Raw JSON files retained') -ForegroundColor DarkGray
}
