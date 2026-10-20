#Requires -Version 7.0
<#
.SYNOPSIS
    Run M365 tenant assessment - collects AD, Exchange, Graph, SharePoint data and generates Excel workbook + CSV files.
.PARAMETER VbuDomain
    Source tenant domain (e.g. contoso.com). Prompted if omitted.
.PARAMETER VbuSearchTerm
    Company name fragment to filter objects (e.g. Contoso). Optional - leave blank to skip search term filtering. Prompted if omitted.
.PARAMETER VbuId
    extensionAttribute7 exact value for filtering. Optional - leave blank to skip VBU ID filtering. Prompted if omitted.
.PARAMETER SharePointAdminUrl
    SharePoint Online admin URL (e.g. https://contoso-admin.sharepoint.com). Prompted if omitted.
.PARAMETER OutputPath
    Output folder path. Defaults to VGMigrations\Assessment\{VBUName}-{timestamp} if omitted.
.EXAMPLE
    .\Run-Assessment.ps1
    (Interactive mode - prompts for all values)
.EXAMPLE
    .\Run-Assessment.ps1 -VbuDomain contoso.com -VbuSearchTerm Contoso -VbuId "12345"
    (Non-interactive mode - runs with provided parameters)
#>
param(
    [string]$VbuDomain = '',
    [string]$VbuSearchTerm = '',
    [string]$VbuId = '',
    [string]$SharePointAdminUrl = '',
    [string]$OutputPath = ''
)

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
)

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
# Mode selection (skip if VbuDomain parameter provided - run assessment directly)
# -----------------------------------------------------------------------
if ($VbuDomain) {
    $mode = 'Run Assessment'  # Skip menu when called with parameters from GUI
} else {
    $mode = Show-ModeMenu
    Write-Host ''
}

if ($mode -eq 'Generate Mapping Files') {
    # Mapping mode reads only the workbook - just ImportExcel and the MappingFiles module
    Import-Module ImportExcel -DisableNameChecking -ErrorAction Stop
    Import-Module (Join-Path $moduleRoot 'MappingFiles.psm1') -Force -DisableNameChecking -Global
    Invoke-MappingFileGeneration
    return
}

# -----------------------------------------------------------------------
# Service modules - hard-fail (all required except SPO)
# -----------------------------------------------------------------------
Import-Module ActiveDirectory                  -DisableNameChecking -ErrorAction Stop
Import-Module ExchangeOnlineManagement         -DisableNameChecking -ErrorAction Stop

# Pin ALL Microsoft.Graph.* modules to 2.33.0 to avoid MSAL assembly conflict with ExchangeOnlineManagement.
# Newer versions (2.34+) require WAM (Web Account Manager) which conflicts with Exchange's own MSAL usage.
# Version 2.33.0 uses browser-popup auth and coexists peacefully with Exchange.
Import-Module Microsoft.Graph.Authentication   -RequiredVersion 2.33.0 -DisableNameChecking -ErrorAction Stop
Import-Module Microsoft.Graph.Groups           -RequiredVersion 2.33.0 -DisableNameChecking -ErrorAction Stop
Import-Module Microsoft.Graph.Teams            -RequiredVersion 2.33.0 -DisableNameChecking -ErrorAction Stop
Import-Module Microsoft.Graph.Users            -RequiredVersion 2.33.0 -DisableNameChecking -ErrorAction Stop

# DeviceManagement v2.33.0 may not be installed - soft-fail (devices collection will be skipped)
try {
    Import-Module Microsoft.Graph.DeviceManagement -RequiredVersion 2.33.0 -DisableNameChecking -ErrorAction Stop
} catch {
    Write-Host ($PREFIX_WARN + 'Microsoft.Graph.DeviceManagement v2.33.0 not available - device collection will be skipped') -ForegroundColor Yellow
}

Import-Module ImportExcel -DisableNameChecking -ErrorAction Stop

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
Import-Module (Join-Path $moduleRoot 'Prerequisites.psm1') -Force -DisableNameChecking -Global
Import-Module (Join-Path $moduleRoot 'AD.psm1')            -Force -DisableNameChecking -Global
Import-Module (Join-Path $moduleRoot 'Exchange.psm1')      -Force -DisableNameChecking -Global
Import-Module (Join-Path $moduleRoot 'Graph.psm1')         -Force -DisableNameChecking -Global
Import-Module (Join-Path $moduleRoot 'SharePoint.psm1')    -Force -DisableNameChecking -Global
Import-Module (Join-Path $moduleRoot 'Workbook.psm1')      -Force -DisableNameChecking -Global

# -----------------------------------------------------------------------
# User inputs (use parameters if provided, otherwise prompt)
# -----------------------------------------------------------------------
$vbuDomain     = if ($VbuDomain) { $VbuDomain.Trim() } else { (Read-Host 'VBU Domain       (e.g. contoso.com)').Trim() }
$vbuSearchTerm = if ($PSBoundParameters.ContainsKey('VbuSearchTerm')) { $VbuSearchTerm.Trim() } else { (Read-Host 'VBU Search Term  (e.g. Contoso)').Trim() }
$vbuId         = if ($PSBoundParameters.ContainsKey('VbuId')) { $VbuId.Trim() } else { (Read-Host 'VBU ID           (extensionAttribute7 exact value)').Trim() }
$spoInput      = if ($SharePointAdminUrl) { $SharePointAdminUrl.Trim() } else { (Read-Host 'SPO Admin URL    [https://ourvolaris-admin.sharepoint.com]').Trim() }
$spoAdminUrl   = if ($spoInput) { $spoInput } else { 'https://ourvolaris-admin.sharepoint.com' }

# -----------------------------------------------------------------------
# Folder structure
# -----------------------------------------------------------------------
# Temporary context call to derive VBUName via the shared TLD-strip logic in Common.psm1
$vbuName      = (New-AssessmentContext -VBUDomain $vbuDomain -VBUId $vbuId -VBUSearchTerm $vbuSearchTerm -RawPath 'TEMP').VBUName
$timestamp    = Get-Date -Format 'yyyyMMdd-HHmm'
$assessFolder = if ($OutputPath) {
    # Create domain-specific subfolder: e.g. C:\M365Discovery\trapezegroup.me
    $basePath = $OutputPath.Trim()
    Join-Path $basePath $vbuDomain
} else {
    Join-Path $PSScriptRoot "$vbuName-$timestamp"
}
$rawPath      = Join-Path $assessFolder 'Raw'
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
if (-not $ctx.SkipSharePoint) {
    Write-Host ($PREFIX_INFO + 'Connecting to SharePoint Online...') -ForegroundColor DarkGray
    try {
        # -UseSystemBrowser required when running from Electron app (no window handle for popup)
        # Opens OS default browser instead - same fix as provision-onedrives.ps1
        try { Disconnect-SPOService -ErrorAction SilentlyContinue } catch {}
        Connect-SPOService -Url $ctx.SPOAdminUrl -UseSystemBrowser:$true -ErrorAction Stop
        Write-Host ($PREFIX_OK + 'SharePoint Online connected') -ForegroundColor Green
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
    }
    catch {
        Write-Host ($PREFIX_FAIL + 'Workbook generation error: ' + $_.Exception.Message) -ForegroundColor Red
    }

    # Export individual CSV files (search-domain.ps1 format) for scripts that need them
    try {
        Write-Host ($PREFIX_INFO + 'Exporting individual CSV files...') -ForegroundColor DarkGray
        $csvFolder = Join-Path $assessFolder 'CSV'
        New-Item -ItemType Directory -Path $csvFolder -Force | Out-Null

        # Map Assessment sheet names to search-domain.ps1 CSV filenames
        # Note: search-domain.ps1 merges AD Users + User Mailboxes into 02_ADUsers.csv
        # but Assessment has them as separate sheets, so we export User Mailboxes as 02b
        $sheetToCsv = @{
            'Accepted Domains'             = '01_AcceptedDomains.csv'
            'AD Users'                     = '02_ADUsers.csv'
            'User Mailboxes'               = '02b_UserMailboxes.csv'
            'Distribution Groups'          = '03_DistributionGroups.csv'
            'Mail Contacts'                = '04_MailContacts.csv'
            'Shared Mailboxes'             = '05_SharedMailboxes.csv'
            'M365 Groups'                  = '06_M365Groups.csv'
            'App Registrations'            = '07_AppRegistrations.csv'
            'Enterprise Apps'              = '08_EnterpriseApps.csv'
            'SharePoint Sites'             = '09_SharePointSites.csv'
            'OneDrives'                    = '10_OneDrives.csv'
            'Teams'                        = '11_Teams.csv'
            'Devices'                      = '12_Devices.csv'
            'AD Groups'                    = '13_ADGroups.csv'
            'Transport Rules'              = '14_TransportRules.csv'
            'Power Platform'               = '15_PowerPlatform.csv'
            'Connectors'                   = '16_Connectors.csv'
            'Remote Domains'               = '17_RemoteDomains.csv'
            'Journal Rules'                = '18_JournalRules.csv'
            'Mailbox Permissions'          = '19_MailboxPermissions.csv'
            'Mail-Enabled Security Groups' = '20_MailEnabledSecurityGroups.csv'
            'Dynamic Distribution Groups'  = '21_DynamicDistributionGroups.csv'
            'Resource Mailboxes'           = '22_ResourceMailboxes.csv'
            'AD Group Memberships'         = '23_ADGroupMemberships.csv'
        }

        foreach ($sheet in $sheetToCsv.Keys) {
            $csvName = $sheetToCsv[$sheet]
            $csvPath = Join-Path $csvFolder $csvName
            try {
                $data = Import-Excel -Path $xlsxPath -WorksheetName $sheet -ErrorAction Stop
                if ($data) {
                    $data | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
                    Write-Host "  [OK]   $csvName" -ForegroundColor DarkGray
                }
            }
            catch {
                # Sheet might not exist or be empty - skip silently
            }
        }
        Write-Host ($PREFIX_OK + "CSV files exported to: $csvFolder") -ForegroundColor Green
    }
    catch {
        Write-Host ($PREFIX_WARN + 'CSV export failed: ' + $_.Exception.Message) -ForegroundColor Yellow
    }

    # Final summary
    $adCounts  = if ($adResult)  { $adResult.Counts }  else { $null }
    $exCounts  = if ($exResult)  { $exResult.Counts }  else { $null }
    $grCounts  = if ($grResult)  { $grResult.Counts }  else { $null }
    $spoCounts = if ($spoResult) { $spoResult.Counts } else { $null }

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
    Write-Host ''
    Write-Host ($PREFIX_OK + 'Workbook: ' + $xlsxPath) -ForegroundColor Green
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
