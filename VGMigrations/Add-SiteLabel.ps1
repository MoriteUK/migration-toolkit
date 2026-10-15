#Requires -Version 7.0
# ═════════════════════════════════════════════════════════════════════════════
# ADD SENSITIVITY LABEL TO SHAREPOINT SITE
# ═════════════════════════════════════════════════════════════════════════════
<#
.SYNOPSIS
    Adds a sensitivity label to a SharePoint Online site.

.DESCRIPTION
    Applies the "Internal Use" sensitivity label to a SharePoint site using PnP PowerShell.
    The label must already exist in the Microsoft Purview Information Protection catalog.

.PARAMETER SiteUrl
    The URL of the SharePoint site (e.g., https://mbuexpretio.sharepoint.com/sites/GRPProduct)

.PARAMETER LabelName
    The name of the sensitivity label to apply (default: "Internal Use")

.EXAMPLE
    .\Add-SiteLabel.ps1 -SiteUrl "https://mbuexpretio.sharepoint.com/sites/GRPProduct"

.NOTES
    Requirements:
    - PnP.PowerShell module (Install-Module PnP.PowerShell -Scope CurrentUser)
    - Global Admin or SharePoint Admin permissions
    - Sensitivity labels must be configured in Microsoft Purview
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl = "https://mbuexpretio.sharepoint.com/sites/GRPProduct",

    [Parameter(Mandatory=$false)]
    [string]$LabelName = "Internal Use"
)

function Add-SiteLabel {
    param(
        [string]$Url,
        [string]$Label
    )

    try {
        Write-Host "Connecting to site: $Url" -ForegroundColor Cyan

        # Connect to the site
        Connect-PnPOnline -Url $Url -Interactive -ErrorAction Stop
        Write-Host "✓ Connected successfully" -ForegroundColor Green

        # Get available labels
        Write-Host "`nRetrieving available sensitivity labels..." -ForegroundColor Cyan
        $labels = Get-PnPAvailableSensitivityLabel -ErrorAction SilentlyContinue

        if (-not $labels) {
            Write-Host "⚠ No sensitivity labels found in the tenant" -ForegroundColor Yellow
            Write-Host "  Ensure labels are configured in Microsoft Purview Information Protection" -ForegroundColor Yellow
            return
        }

        # Find matching label
        $matchedLabel = $labels | Where-Object { $_.DisplayName -eq $Label }

        if (-not $matchedLabel) {
            Write-Host "✗ Label '$Label' not found in tenant" -ForegroundColor Red
            Write-Host "`nAvailable labels:" -ForegroundColor Yellow
            $labels | ForEach-Object { Write-Host "  • $($_.DisplayName)" }
            return
        }

        Write-Host "✓ Label found: $($matchedLabel.DisplayName)" -ForegroundColor Green

        # Apply the label to the site
        Write-Host "`nApplying label to site..." -ForegroundColor Cyan
        Set-PnPSite -Identity $Url -Sensitivity $matchedLabel.DisplayName -ErrorAction Stop

        Write-Host "✓ Label '$Label' successfully applied to $Url" -ForegroundColor Green

        # Verify the label was applied
        Write-Host "`nVerifying label application..." -ForegroundColor Cyan
        $site = Get-PnPSite -Identity $Url -ErrorAction SilentlyContinue
        if ($site.SensitivityLabel) {
            Write-Host "✓ Verified: Site sensitivity label is now '$($site.SensitivityLabel)'" -ForegroundColor Green
        }

    } catch {
        Write-Host "✗ Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  at $($_.InvocationInfo.ScriptName):$($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
    } finally {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        Write-Host "`nDisconnected from SharePoint" -ForegroundColor Cyan
    }
}

# ── MAIN ──────────────────────────────────────────────────────────────────────

# Check if PnP.PowerShell module is installed
if (-not (Get-Module -Name PnP.PowerShell -ListAvailable -ErrorAction SilentlyContinue)) {
    Write-Host "PnP.PowerShell module not found" -ForegroundColor Red
    Write-Host "Installing PnP.PowerShell..." -ForegroundColor Yellow
    Install-Module PnP.PowerShell -Scope CurrentUser -Force
    Write-Host "✓ PnP.PowerShell installed" -ForegroundColor Green
}

# Run the label addition
Add-SiteLabel -Url $SiteUrl -Label $LabelName
