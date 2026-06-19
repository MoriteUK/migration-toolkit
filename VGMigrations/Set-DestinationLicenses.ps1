<#
.SYNOPSIS
    Assigns Microsoft 365 Business Basic or Business Premium licenses to users
    in a destination tenant, using the report from Get-UserMailboxLicenses.ps1.

.DESCRIPTION
    Reads the license-report.csv produced by Get-UserMailboxLicenses.ps1.
    Uses the DestinationUPN column to locate each user in the destination tenant,
    so cross-domain migrations (where the UPN changes) are handled correctly.

.PARAMETER ReportCsvPath
    Path to the license-report.csv produced by Get-UserMailboxLicenses.ps1.
    Must contain: SourceUPN, DestinationUPN, SuggestedTarget.

.PARAMETER TenantId
    The destination tenant ID or domain (e.g. fabrikam.onmicrosoft.com).

.PARAMETER DefaultLicense
    Fallback when SuggestedTarget is 'None' or blank.
    Values: BusinessBasic | BusinessPremium | Skip (default: Skip)

.PARAMETER OverrideLicense
    Assigns this tier to ALL users regardless of SuggestedTarget.
    Values: BusinessBasic | BusinessPremium

.PARAMETER WhatIf
    Show what would happen without making any changes.

.EXAMPLE
    # Use the suggested tiers from the report
    .\Set-DestinationLicenses.ps1 -ReportCsvPath C:\license-report.csv -TenantId fabrikam.onmicrosoft.com

    # Force everyone to Business Basic, dry run first
    .\Set-DestinationLicenses.ps1 -ReportCsvPath C:\license-report.csv -TenantId fabrikam.onmicrosoft.com -OverrideLicense BusinessBasic -WhatIf

    # Assign Business Premium to all; fall back to Basic for users with no suggested tier
    .\Set-DestinationLicenses.ps1 -ReportCsvPath C:\license-report.csv -TenantId fabrikam.onmicrosoft.com -DefaultLicense BusinessBasic
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$ReportCsvPath,

    [Parameter(Mandatory)]
    [string]$TenantId,

    [ValidateSet('BusinessBasic', 'BusinessPremium', 'Skip')]
    [string]$DefaultLicense = 'Skip',

    [ValidateSet('BusinessBasic', 'BusinessPremium')]
    [string]$OverrideLicense
)

#region --- Target SKU part numbers ------------------------------------------

$TargetSkuPartNumbers = @{
    BusinessBasic   = 'O365_BUSINESS_ESSENTIALS'
    BusinessPremium = 'SPB'
}

#endregion

#region --- Prerequisites ----------------------------------------------------

if (-not (Get-Module Microsoft.Graph.Users -ListAvailable) -or
    -not (Get-Module Microsoft.Graph.Identity.DirectoryManagement -ListAvailable)) {
    Write-Host "Installing Microsoft.Graph modules..." -ForegroundColor Cyan
    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
}

#endregion

#region --- Connect to destination tenant ------------------------------------

Write-Host "Connecting to destination tenant: $TenantId" -ForegroundColor Cyan
Connect-MgGraph -TenantId $TenantId -Scopes @('User.ReadWrite.All', 'Organization.Read.All') -NoWelcome

$tenantSkus = Get-MgSubscribedSku -All
$skuIdMap   = @{}

foreach ($tier in $TargetSkuPartNumbers.Keys) {
    $partNumber = $TargetSkuPartNumbers[$tier]
    $match = $tenantSkus | Where-Object { $_.SkuPartNumber -eq $partNumber }

    if (-not $match) {
        Write-Warning "SKU '$partNumber' ($tier) is not available in the destination tenant — users needing this tier will be skipped."
        continue
    }

    $available = $match.PrepaidUnits.Enabled - $match.ConsumedUnits
    Write-Host "  $tier ($partNumber): $($match.ConsumedUnits) used / $($match.PrepaidUnits.Enabled) total  ($available available)" -ForegroundColor Gray
    $skuIdMap[$tier] = $match.SkuId
}

if ($skuIdMap.Count -eq 0) {
    Write-Error "Neither target SKU is available in the destination tenant. Aborting."
    Disconnect-MgGraph | Out-Null
    exit 1
}

#endregion

#region --- Read report ------------------------------------------------------

if (-not (Test-Path $ReportCsvPath)) {
    Write-Error "Report CSV not found: $ReportCsvPath"
    Disconnect-MgGraph | Out-Null
    exit 1
}

$rows = Import-Csv $ReportCsvPath | Where-Object { $_.HasMailbox -ne 'ERROR' }
Write-Host "$($rows.Count) users loaded from report." -ForegroundColor Cyan

# Validate the report has a DestinationUPN column
$hasDestCol = $rows.Count -gt 0 -and $rows[0].PSObject.Properties['DestinationUPN']
if (-not $hasDestCol) {
    Write-Warning "Report does not have a 'DestinationUPN' column — falling back to SourceUPN for all users."
    Write-Warning "If UPNs differ between tenants, re-run Get-UserMailboxLicenses.ps1 with a DestinationUPN column in the input CSV."
}

#endregion

#region --- Process users ----------------------------------------------------

$results = [System.Collections.Generic.List[pscustomobject]]::new()
$i = 0

foreach ($row in $rows) {
    $i++

    $srcUpn = $row.SourceUPN.Trim()
    $dstUpn = if ($hasDestCol -and -not [string]::IsNullOrWhiteSpace($row.DestinationUPN)) {
                  $row.DestinationUPN.Trim()
              } else {
                  $srcUpn
              }

    Write-Progress -Activity "Assigning licenses" `
        -Status $(if ($dstUpn -ne $srcUpn) { "$srcUpn -> $dstUpn" } else { $dstUpn }) `
        -PercentComplete (($i / $rows.Count) * 100)

    $tier = if ($OverrideLicense)                                         { $OverrideLicense }
            elseif ($row.SuggestedTarget -in 'BusinessBasic','BusinessPremium') { $row.SuggestedTarget }
            else                                                          { $DefaultLicense }

    $status = [pscustomobject]@{
        SourceUPN         = $srcUpn
        DestinationUPN    = $dstUpn
        DisplayName       = $row.DisplayName
        SourceLicenses    = $row.MailboxLicenses
        TargetTier        = $tier
        Action            = ''
        Result            = ''
        Error             = ''
    }

    if ($tier -eq 'Skip') {
        $status.Action = 'Skipped'
        $status.Result = 'No target tier — use -DefaultLicense or -OverrideLicense to assign one'
        $results.Add($status)
        continue
    }

    $skuId = $skuIdMap[$tier]
    if (-not $skuId) {
        $status.Action = 'Skipped'
        $status.Result = "SKU for $tier not available in destination tenant"
        $results.Add($status)
        continue
    }

    # Look up the user by their DESTINATION UPN in the destination tenant
    try {
        $destUser = Get-MgUser -UserId $dstUpn `
            -Property 'Id,DisplayName,AssignedLicenses' `
            -ErrorAction Stop
    }
    catch {
        $status.Action = 'Error'
        $status.Result = "User '$dstUpn' not found in destination tenant"
        $status.Error  = $_.Exception.Message
        $results.Add($status)
        continue
    }

    $alreadyAssigned = $destUser.AssignedLicenses | Where-Object { $_.SkuId -eq $skuId }
    if ($alreadyAssigned) {
        $status.Action = 'AlreadyAssigned'
        $status.Result = "$tier already present — no change"
        $results.Add($status)
        continue
    }

    $status.Action = "Assign-$tier"

    if ($PSCmdlet.ShouldProcess($dstUpn, "Assign $tier license (SkuId $skuId)")) {
        try {
            Set-MgUserLicense -UserId $destUser.Id `
                -AddLicenses @(@{ SkuId = $skuId }) `
                -RemoveLicenses @() `
                -ErrorAction Stop

            $status.Result = 'Success'
        }
        catch {
            $status.Result = 'Failed'
            $status.Error  = $_.Exception.Message
        }
    }
    else {
        $status.Result = 'WhatIf — no change made'
    }

    $results.Add($status)
}

Write-Progress -Activity "Assigning licenses" -Completed

#endregion

#region --- Output results ---------------------------------------------------

$resultsPath = [System.IO.Path]::Combine(
    [System.IO.Path]::GetDirectoryName((Resolve-Path $ReportCsvPath)),
    'license-assignment-results.csv'
)
$results | Export-Csv -Path $resultsPath -NoTypeInformation -Encoding UTF8
Write-Host "`nResults: $resultsPath" -ForegroundColor Green

Write-Host "`nSummary:" -ForegroundColor Cyan
$results | Group-Object Result | Sort-Object Count -Descending | Format-Table Name, Count -AutoSize

$errors = $results | Where-Object { $_.Error }
if ($errors) {
    Write-Host "Errors:" -ForegroundColor Red
    $errors | Format-Table SourceUPN, DestinationUPN, Action, Error -AutoSize
}

Disconnect-MgGraph | Out-Null

#endregion
