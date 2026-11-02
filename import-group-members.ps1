<#
.SYNOPSIS
    Imports group members from CSV files with email domain rewriting.

.DESCRIPTION
    Reads member CSV files (exported by export-group-members.ps1) and adds members
    to the corresponding groups in the target tenant after rewriting email addresses
    to the target domain.

.PARAMETER MemberCSVs
    Path to the member CSV file(s). Supports wildcards (e.g., '*_Members_*.csv').

.PARAMETER TargetTenant
    The target tenant name (e.g., 'contoso'). The script will automatically append '.onmicrosoft.com'.
    All email addresses will be rewritten to <TargetTenant>.onmicrosoft.com.

.PARAMETER SkipExisting
    If set, skip members that are already in the group.

.EXAMPLE
    .\import-group-members.ps1 -MemberCSVs '.\MB CELCAT*_Members_*.csv' -TargetTenant 'aioneducation'

.EXAMPLE
    .\import-group-members.ps1 -MemberCSVs '.\MB CELCAT Finance_Members_20261101_120000.csv' -TargetTenant 'aioneducation' -WhatIf
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true)]
    [string]$MemberCSVs,

    [Parameter(Mandatory = $true)]
    [string]$TargetTenant,

    [switch]$SkipExisting
)

$ErrorActionPreference = 'Stop'

# Build target domain (always .onmicrosoft.com)
$targetDomain = if ($TargetTenant -like '*.onmicrosoft.com') {
    $TargetTenant
} else {
    "$TargetTenant.onmicrosoft.com"
}

Write-Host "`nTarget tenant domain: $targetDomain" -ForegroundColor Cyan

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    throw "ExchangeOnlineManagement module not found. Install with: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
}

if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:$false
}

# Helper function to rewrite email addresses to target domain
function Convert-ToTargetDomain {
    param([string]$Email)
    if (-not $Email) { return $null }
    $localPart = $Email.Split('@')[0]
    return "$localPart@$targetDomain"
}

# Find CSV files
$csvFiles = @(Get-ChildItem -Path $MemberCSVs -ErrorAction SilentlyContinue)

if ($csvFiles.Count -eq 0) {
    throw "No CSV files found matching: $MemberCSVs"
}

Write-Host "Found $($csvFiles.Count) CSV file(s) to process`n" -ForegroundColor Green

# Track results
$results = @{
    FilesProcessed = 0
    GroupsProcessed = 0
    MembersAdded = 0
    MembersSkipped = 0
    MembersFailed = 0
}

$failedMembers = [System.Collections.Generic.List[object]]::new()

foreach ($csvFile in $csvFiles) {
    Write-Host "Processing: $($csvFile.Name)" -ForegroundColor Cyan

    try {
        $members = Import-Csv -Path $csvFile.FullName

        if ($members.Count -eq 0) {
            Write-Warning "  No members found in CSV"
            continue
        }

        # Get group info from first row
        $sourceGroupEmail = $members[0].GroupEmail
        $groupName = $members[0].GroupName

        # Rewrite group email to target domain
        $targetGroupEmail = Convert-ToTargetDomain -Email $sourceGroupEmail

        Write-Host "  Group: $groupName" -ForegroundColor White
        Write-Host "  Source: $sourceGroupEmail" -ForegroundColor DarkGray
        Write-Host "  Target: $targetGroupEmail" -ForegroundColor White

        # Check if group exists in target tenant
        $targetGroup = Get-DistributionGroup -Identity $targetGroupEmail -ErrorAction SilentlyContinue

        if (-not $targetGroup) {
            Write-Warning "  Group not found in target tenant: $targetGroupEmail"
            Write-Warning "  Skipping all members for this group"
            continue
        }

        $results.FilesProcessed++
        $results.GroupsProcessed++

        # Get existing members once for efficiency
        $existingMembers = @(Get-DistributionGroupMember -Identity $targetGroupEmail -ErrorAction SilentlyContinue)
        $existingEmails = $existingMembers | ForEach-Object { $_.PrimarySmtpAddress.ToString().ToLower() }

        # Process each member
        $memberCount = 0
        foreach ($member in $members) {
            try {
                # Rewrite member email to target domain
                $sourceMemberEmail = $member.PrimarySmtpAddress
                $targetMemberEmail = Convert-ToTargetDomain -Email $sourceMemberEmail

                # Check if member exists in target tenant
                $targetMember = Get-Recipient -Identity $targetMemberEmail -ErrorAction SilentlyContinue

                if (-not $targetMember) {
                    Write-Warning "    Member not found in target tenant: $targetMemberEmail"
                    $results.MembersFailed++
                    $failedMembers.Add([pscustomobject]@{
                        CSVFile = $csvFile.Name
                        Group = $groupName
                        GroupEmail = $targetGroupEmail
                        OriginalMember = $sourceMemberEmail
                        TargetMember = $targetMemberEmail
                        Reason = 'Member not found in target tenant'
                    })
                    continue
                }

                # Check if already a member
                if ($existingEmails -contains $targetMemberEmail.ToLower()) {
                    if ($SkipExisting) {
                        $results.MembersSkipped++
                        continue
                    }
                }

                if ($WhatIfPreference) {
                    Write-Host "    [WHATIF] Would add: $targetMemberEmail" -ForegroundColor Cyan
                    $memberCount++
                } else {
                    Add-DistributionGroupMember -Identity $targetGroupEmail -Member $targetMemberEmail -ErrorAction Stop
                    $memberCount++
                    $results.MembersAdded++
                }

            } catch {
                Write-Warning "    Failed to add $targetMemberEmail : $($_.Exception.Message)"
                $results.MembersFailed++
                $failedMembers.Add([pscustomobject]@{
                    CSVFile = $csvFile.Name
                    Group = $groupName
                    GroupEmail = $targetGroupEmail
                    OriginalMember = $sourceMemberEmail
                    TargetMember = $targetMemberEmail
                    Reason = $_.Exception.Message
                })
            }
        }

        Write-Host "  ✓ Added $memberCount member(s)" -ForegroundColor Green

    } catch {
        Write-Warning "  Failed to process CSV: $($_.Exception.Message)"
    }

    Write-Host ""
}

# Summary
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "IMPORT SUMMARY" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host ""
Write-Host "CSV Files:" -ForegroundColor White
Write-Host "  Processed: $($results.FilesProcessed)" -ForegroundColor Green
Write-Host ""
Write-Host "Groups:" -ForegroundColor White
Write-Host "  Processed: $($results.GroupsProcessed)" -ForegroundColor Green
Write-Host ""
Write-Host "Members:" -ForegroundColor White
Write-Host "  Added:     $($results.MembersAdded)" -ForegroundColor Green
Write-Host "  Skipped:   $($results.MembersSkipped)" -ForegroundColor Yellow
Write-Host "  Failed:    $($results.MembersFailed)" -ForegroundColor Red
Write-Host ""

# Export failed members to CSV
if ($failedMembers.Count -gt 0) {
    $failedFile = Join-Path (Split-Path $csvFiles[0].FullName) "FailedMemberImports_${targetDomain}_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $failedMembers | Export-Csv -Path $failedFile -NoTypeInformation -Encoding UTF8
    Write-Host "Failed members exported to: $failedFile" -ForegroundColor Yellow
    Write-Host ""
}

Write-Host "Import complete!" -ForegroundColor Green
