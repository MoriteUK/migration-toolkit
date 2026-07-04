#Requires -Modules ImportExcel
#Requires -Modules Microsoft.Online.SharePoint.PowerShell

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$SourceWorkbook
)

$ErrorActionPreference = 'Stop'

# Read SPO admin URL from shared config; fall back to Volaris default
$TenantAdminUrl = 'https://ourvolaris-admin.sharepoint.com'
$_sharedCfg = Join-Path $env:LOCALAPPDATA 'FlyMigration\shared-config.json'
if (Test-Path $_sharedCfg) {
    try {
        $_cfg = Get-Content $_sharedCfg -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($_cfg.SharePointAdminUrl) { $TenantAdminUrl = $_cfg.SharePointAdminUrl.TrimEnd('/') }
    } catch {}
}

Add-Type -AssemblyName System.Windows.Forms

# ============================================================
# CONNECT TO SPO
# ============================================================

Write-Host "Connecting to SharePoint Online..." -ForegroundColor Cyan
Write-Host "(A browser sign-in tab will open — authenticate and then return here)" -ForegroundColor Yellow
Connect-SPOService -Url $TenantAdminUrl -UseWebLogin
Write-Host "Connected." -ForegroundColor Green

# ============================================================
# IMPORT WORKBOOK
# ============================================================

Write-Host "Loading workbook..." -ForegroundColor Cyan

$Teams  = Import-Excel -Path $SourceWorkbook -WorksheetName "Teams"
$Groups = Import-Excel -Path $SourceWorkbook -WorksheetName "M365Groups"
$Sites  = Import-Excel -Path $SourceWorkbook -WorksheetName "SharePointSites"

# ============================================================
# LOOKUPS
# ============================================================

$TeamSmtpLookup   = @{}
$ConsumedSiteUrls = @{}
$SiteSourceLookup = @{}

# ============================================================
# TEAMS
# ============================================================

$TeamsOutput = @()

foreach ($Team in $Teams)
{
    $smtp = $Team.PrimarySmtpAddress

    if ($smtp)
    {
        $smtp = $smtp.ToString().Trim().ToLower()
        $TeamSmtpLookup[$smtp] = $true
    }

    if ($Team.SharePointSiteUrl)
    {
        $siteUrl = $Team.SharePointSiteUrl.ToString().Trim().TrimEnd('/').ToLower()

        $ConsumedSiteUrls[$siteUrl] = $true
        $SiteSourceLookup[$siteUrl] = "Represented by Team"
    }

    $TeamsOutput += $Team
}

$TeamsOutput =
    $TeamsOutput |
    Sort-Object PrimarySmtpAddress -Unique

# ============================================================
# GROUPS
# ============================================================

$GroupsOutput = @()

foreach ($Group in $Groups)
{
    $smtp = $null

    if ($Group.PrimarySmtpAddress)
    {
        $smtp = $Group.PrimarySmtpAddress.ToString().Trim().ToLower()
    }

    $IsTeam = $false

    if ($null -ne $Group.IsTeam)
    {
        switch ($Group.IsTeam.ToString().ToLower())
        {
            "true" { $IsTeam = $true }
            "yes"  { $IsTeam = $true }
            "1"    { $IsTeam = $true }
        }
    }

    if ($IsTeam)
    {
        continue
    }

    if ($smtp -and $TeamSmtpLookup.ContainsKey($smtp))
    {
        continue
    }

    $GroupsOutput += $Group

    if ($Group.SharePointSiteUrl)
    {
        $siteUrl = $Group.SharePointSiteUrl.ToString().Trim().TrimEnd('/').ToLower()

        $ConsumedSiteUrls[$siteUrl] = $true
        $SiteSourceLookup[$siteUrl] = "Represented by M365 Group"
    }
}

$GroupsOutput =
    $GroupsOutput |
    Sort-Object PrimarySmtpAddress -Unique

# ============================================================
# SHAREPOINT VALIDATION
# ============================================================

$SharePointOutput = @()
$ExcludedSites    = @()

foreach ($Site in $Sites)
{
    if (-not $Site.WebUrl)
    {
        continue
    }

    $SiteUrl = $Site.WebUrl.ToString().Trim().TrimEnd('/').ToLower()

    Write-Host "Checking $SiteUrl" -ForegroundColor Cyan

    # Already represented by Team or Group
    if ($ConsumedSiteUrls.ContainsKey($SiteUrl))
    {
        $ExcludedSites += [PSCustomObject]@{
            SiteUrl         = $SiteUrl
            Title           = $Site.Title
            Template        = ""
            RelatedGroupId  = ""
            ExclusionReason = $SiteSourceLookup[$SiteUrl]
        }

        continue
    }

    # SPO Validation
    try
    {
        $SpoSite = Get-SPOSite -Identity $SiteUrl -Detailed -ErrorAction Stop
    }
    catch
    {
        $ExcludedSites += [PSCustomObject]@{
            SiteUrl         = $SiteUrl
            Title           = $Site.Title
            Template        = ""
            RelatedGroupId  = ""
            ExclusionReason = "Unable to query site"
        }

        continue
    }

    $Template = $SpoSite.Template
    $RelatedGroupId = $SpoSite.RelatedGroupId

    # Channel Sites
    if ($Template -match "CHANNEL")
    {
        $ExcludedSites += [PSCustomObject]@{
            SiteUrl         = $SiteUrl
            Title           = $SpoSite.Title
            Template        = $Template
            RelatedGroupId  = $RelatedGroupId
            ExclusionReason = "Private or Shared Channel Site"
        }

        continue
    }

    # Group Connected
    if ($RelatedGroupId `
        -and `
        $RelatedGroupId -ne "00000000-0000-0000-0000-000000000000")
    {
        $ExcludedSites += [PSCustomObject]@{
            SiteUrl         = $SiteUrl
            Title           = $SpoSite.Title
            Template        = $Template
            RelatedGroupId  = $RelatedGroupId
            ExclusionReason = "Group Connected Site"
        }

        continue
    }

    $SharePointOutput += $Site
}

$SharePointOutput =
    $SharePointOutput |
    Sort-Object WebUrl -Unique

# ============================================================
# EXPORT
# ============================================================

$OutputWorkbook = Join-Path `
    (Split-Path $SourceWorkbook -Parent) `
    "Discovery-Classified.xlsx"

if (Test-Path $OutputWorkbook)
{
    Remove-Item $OutputWorkbook -Force
}

$TeamsOutput |
    Export-Excel `
        -Path $OutputWorkbook `
        -WorksheetName "Teams" `
        -AutoSize `
        -FreezeTopRow `
        -BoldTopRow

$GroupsOutput |
    Export-Excel `
        -Path $OutputWorkbook `
        -WorksheetName "Groups" `
        -AutoSize `
        -FreezeTopRow `
        -BoldTopRow

$SharePointOutput |
    Export-Excel `
        -Path $OutputWorkbook `
        -WorksheetName "SharePoint" `
        -AutoSize `
        -FreezeTopRow `
        -BoldTopRow

$ExcludedSites |
    Export-Excel `
        -Path $OutputWorkbook `
        -WorksheetName "ExcludedSites" `
        -AutoSize `
        -FreezeTopRow `
        -BoldTopRow

Write-Host ""
Write-Host "Completed." -ForegroundColor Green
Write-Host ""
Write-Host "Output: $OutputWorkbook" -ForegroundColor Green
Write-Host ""
Write-Host "Teams         : $($TeamsOutput.Count)"
Write-Host "Groups        : $($GroupsOutput.Count)"
Write-Host "SharePoint    : $($SharePointOutput.Count)"
Write-Host "ExcludedSites : $($ExcludedSites.Count)"
