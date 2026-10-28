#Requires -Version 7.0

# Define module root before any imports
$moduleRoot = Join-Path $PSScriptRoot 'Modules'

# Common.psm1 first - prefix constants and console helpers needed for the banner and menu
Import-Module (Join-Path $moduleRoot 'Common.psm1') -Force -DisableNameChecking -Global

# -----------------------------------------------------------------------
# Constants
# -----------------------------------------------------------------------
$script:GraphScopes = @(
    'Group.Read.All'
    'Directory.Read.All'
    'User.Read.All'
    'Device.Read.All'
    'Sites.Read.All'
    'Team.ReadBasic.All'
    'TeamMember.Read.All'
    'DeviceManagementManagedDevices.Read.All'
    'Policy.Read.All'
    'Application.Read.All'
    'Tasks.Read'
    'Organization.Read.All'
)

# -----------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------
function Read-MultiValue {
    param([Parameter(Mandatory)][string]$Prompt)

    $values = [System.Collections.Generic.List[string]]::new()

    # First entry is required - keep asking until something is entered.
    do {
        $first = (Read-Host $Prompt).Trim()
    } while (-not $first)
    $values.Add($first)

    # Subsequent entries are optional - a blank answer ends the list.
    while ($true) {
        $next = (Read-Host "$Prompt (blank to continue)").Trim()
        if (-not $next) { break }
        $values.Add($next)
    }

    return $values.ToArray()
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
$mode = Show-ModeMenu
Write-Host ''

if ($mode -eq 'Generate Mapping Files') {
    # Mapping mode reads only the workbook - just ImportExcel and the MappingFiles module
    Import-Module ImportExcel -DisableNameChecking -ErrorAction Stop
    Import-Module (Join-Path $moduleRoot 'MappingFiles.psm1') -Force -DisableNameChecking -Global
    Invoke-MappingFileGeneration
    return
}

# -----------------------------------------------------------------------
# Prerequisites - run before the service-module imports below so any missing
# module is installed first; only Prerequisites.psm1 itself needs to be loaded for this.
# -----------------------------------------------------------------------
Import-Module (Join-Path $moduleRoot 'Prerequisites.psm1') -Force -DisableNameChecking -Global
Test-Prerequisites

# -----------------------------------------------------------------------
# Service modules - hard-fail (all required except SPO)
# -----------------------------------------------------------------------
Import-Module ActiveDirectory                  -DisableNameChecking -ErrorAction Stop
Import-Module ExchangeOnlineManagement         -DisableNameChecking -ErrorAction Stop
Import-Module Microsoft.Graph.Authentication   -DisableNameChecking -ErrorAction Stop
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
Import-Module (Join-Path $moduleRoot 'AD.psm1')            -Force -DisableNameChecking -Global
Import-Module (Join-Path $moduleRoot 'Exchange.psm1')      -Force -DisableNameChecking -Global
Import-Module (Join-Path $moduleRoot 'Graph.psm1')         -Force -DisableNameChecking -Global
Import-Module (Join-Path $moduleRoot 'SharePoint.psm1')    -Force -DisableNameChecking -Global
Import-Module (Join-Path $moduleRoot 'Workbook.psm1')      -Force -DisableNameChecking -Global

# -----------------------------------------------------------------------
# User inputs
# -----------------------------------------------------------------------
$vbuDomains     = Read-MultiValue -Prompt 'VBU Domain       (e.g. contoso.com)'
$vbuSearchTerms = Read-MultiValue -Prompt 'VBU Search Term  (e.g. Contoso)'
$vbuIds         = Read-MultiValue -Prompt 'VBU ID           (extensionAttribute7 exact value)'
$spoInput       = (Read-Host 'SPO Admin URL    [https://ourvolaris-admin.sharepoint.com]').Trim()
$spoAdminUrl    = if ($spoInput) { $spoInput } else { 'https://ourvolaris-admin.sharepoint.com' }

# -----------------------------------------------------------------------
# Folder structure
# -----------------------------------------------------------------------
# Temporary context call to derive VBUName via the shared TLD-strip logic in Common.psm1
# (VBUName is folder/file naming only - derived from the first domain entered)
$vbuName      = (New-AssessmentContext -VBUDomains $vbuDomains -VBUIds $vbuIds -VBUSearchTerms $vbuSearchTerms -RawPath 'TEMP').VBUName
$timestamp    = Get-Date -Format 'yyyyMMdd-HHmm'
$assessFolder = Join-Path $PSScriptRoot "$vbuName-$timestamp"
$rawPath      = Join-Path $assessFolder 'Raw'
$xlsxPath     = Join-Path $assessFolder "$vbuName-Assessment.xlsx"

New-Item -ItemType Directory -Path $rawPath                                 -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $assessFolder 'MappingFiles') -Force | Out-Null

Write-Host ($PREFIX_OK + "Output folder: $assessFolder") -ForegroundColor Green

# -----------------------------------------------------------------------
# Assessment context
# -----------------------------------------------------------------------
$ctx = New-AssessmentContext `
    -VBUDomains     $vbuDomains `
    -VBUIds         $vbuIds `
    -VBUSearchTerms $vbuSearchTerms `
    -RawPath        $rawPath `
    -SPOAdminUrl    $spoAdminUrl

if (-not (Get-Module -ListAvailable -Name 'Microsoft.Online.SharePoint.PowerShell')) {
    $ctx.SkipSharePoint = $true
}

# -----------------------------------------------------------------------
# Authentication - all sessions connected upfront
# -----------------------------------------------------------------------
Write-SectionHeader 'Authentication'

# SPO must connect first - Exchange and Graph WAM tokens conflict with SPO auth if they connect before it
if (-not $ctx.SkipSharePoint) {
    Write-Host ($PREFIX_INFO + 'Connecting to SharePoint Online...') -ForegroundColor DarkGray
    try {
        Connect-SPOService -Url $ctx.SPOAdminUrl -ErrorAction Stop
        Write-Host ($PREFIX_OK + 'SharePoint Online connected') -ForegroundColor Green

        # Multi-Geo - a failed/empty lookup is not a connection failure, just a single-geo
        # tenant; falls back to the one admin URL already connected above.
        $geoRows = @()
        try { $geoRows = @(Get-SPOMultiGeoCompanyAllowedDataLocation -ErrorAction Stop) } catch { $geoRows = @() }

        if ($geoRows.Count -gt 0) {
            $spoGeos = [System.Collections.Generic.List[PSCustomObject]]::new()
            foreach ($geoRow in $geoRows) {
                $prefix   = ("$($geoRow.Domain)" -replace '(?i)\.onmicrosoft\.com$', '').ToLower()
                $geoEntry = [PSCustomObject]@{ Location = "$($geoRow.Location)"; AdminUrl = "https://$prefix-admin.sharepoint.com" }
                if ($geoRow.IsDefault) { $spoGeos.Insert(0, $geoEntry) } else { $spoGeos.Add($geoEntry) }
            }
            $ctx.SPOGeos = $spoGeos.ToArray()
        }
        else {
            $ctx.SPOGeos = @([PSCustomObject]@{ Location = 'Default'; AdminUrl = $ctx.SPOAdminUrl })
        }

        Write-Host ($PREFIX_OK + "SharePoint geos: $(($ctx.SPOGeos | ForEach-Object { $_.Location }) -join ', ')") -ForegroundColor Green
    }
    catch {
        Write-Host ($PREFIX_FAIL + 'SPO connection failed: ' + $_.Exception.Message) -ForegroundColor Red
        Write-Host ($PREFIX_WARN + 'SharePoint collection will be skipped') -ForegroundColor Yellow
        $ctx.SkipSharePoint = $true
    }
}

Write-Host ($PREFIX_INFO + 'Connecting to Exchange Online...') -ForegroundColor DarkGray
Connect-ExchangeOnline -ShowBanner:$false -DisableWAM -ErrorAction Stop
Write-Host ($PREFIX_OK + 'Exchange Online connected') -ForegroundColor Green

$skipGraph  = $false
$graphStart = Get-Date
Write-Host ($PREFIX_INFO + 'Connecting to Microsoft Graph...') -ForegroundColor DarkGray
try {
    Connect-MgGraph -Scopes $script:GraphScopes -NoWelcome -ErrorAction Stop
    Write-Host ($PREFIX_OK + 'Microsoft Graph connected') -ForegroundColor Green
}
catch {
    Write-Host ($PREFIX_FAIL + 'Graph connection failed: ' + $_.Exception.Message) -ForegroundColor Red
    Write-Host ($PREFIX_WARN + 'Graph collection will be skipped') -ForegroundColor Yellow
    $skipGraph = $true
}

# -----------------------------------------------------------------------
# Collection and workbook
# -----------------------------------------------------------------------
$adResult  = $null
$exResult  = $null
$grResult  = $null
$spoResult = $null

try {
    # Active Directory - critical: propagates on failure, stopping the run
    $adResult = Invoke-ADCollection -Context $ctx

    # Exchange Online - non-critical
    try {
        $exResult = Invoke-ExchangeCollection -Context $ctx
    }
    catch {
        Write-Host ($PREFIX_FAIL + 'Exchange collection error: ' + $_.Exception.Message) -ForegroundColor Red
        $exResult = New-CollectorResult -Success $false -ErrorMessage $_.Exception.Message
    }

    # Microsoft Graph - non-critical, user-skippable
    if ($skipGraph) {
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

    # Workbook - non-critical (always runs with whatever JSON was written)
    try {
        Invoke-WorkbookGeneration -Context $ctx -XlsxPath $xlsxPath
        Write-Host ''
        Write-Host ($PREFIX_OK + 'Workbook: ' + $xlsxPath) -ForegroundColor Green
    }
    catch {
        Write-Host ($PREFIX_FAIL + 'Workbook generation error: ' + $_.Exception.Message) -ForegroundColor Red
    }
}
catch {
    Write-Host ($PREFIX_FAIL + 'Assessment stopped: ' + $_.Exception.Message) -ForegroundColor Red
}
finally {
    try { Disconnect-MgGraph              -ErrorAction SilentlyContinue } catch {}
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    try { Disconnect-SPOService           -ErrorAction SilentlyContinue } catch {}
}

# -----------------------------------------------------------------------
# Raw JSON cleanup - always runs regardless of assessment outcome
# -----------------------------------------------------------------------
Write-Host ''
$cleanInput = (Read-Host "Delete Raw JSON files from '$rawPath'? [Y/N]").Trim().ToUpper()
if ($cleanInput -eq 'Y') {
    Get-ChildItem -Path $rawPath -Filter '*.json' | Remove-Item -Force
    Write-Host ($PREFIX_OK + 'Raw JSON files deleted') -ForegroundColor Green
}
else {
    Write-Host ($PREFIX_INFO + 'Raw JSON files retained') -ForegroundColor DarkGray
}
