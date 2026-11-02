<#
.SYNOPSIS
    Imports distribution lists and their members into a new tenant.

.DESCRIPTION
    Reads the CSVs produced by export-dlinfo.ps1 and creates:
      - Distribution groups (including mail-enabled security groups)
      - Group memberships

    Note: Dynamic distribution groups are skipped (they require RecipientFilter which may not work in new tenant).

.PARAMETER DistributionListsCSV
    Path to the DistributionLists_*.csv file from export-dlinfo.ps1

.PARAMETER MembersCSV
    Path to the DistributionListMembers_*.csv file from export-dlinfo.ps1

.PARAMETER TargetTenant
    The target tenant name (e.g., 'contoso'). The script will automatically append '.onmicrosoft.com'.
    All email addresses will be rewritten from the source domain to <TargetTenant>.onmicrosoft.com.

.PARAMETER CreateCloudOnly
    If set, creates dir-synced groups as cloud-only groups in the target tenant (ignores IsDirSynced flag).
    Use this when migrating from hybrid/on-prem to cloud-only.

.PARAMETER SkipExisting
    If set, skip groups that already exist instead of failing.

.PARAMETER WhatIf
    Show what would be created without actually creating anything.

.EXAMPLE
    .\import-dlinfo.ps1 -DistributionListsCSV '.\DistributionLists_contoso.com_20261101_120000.csv' -MembersCSV '.\DistributionListMembers_contoso.com_20261101_120000.csv' -TargetTenant 'contoso'

.EXAMPLE
    .\import-dlinfo.ps1 -DistributionListsCSV '.\DistributionLists_contoso.com_20261101_120000.csv' -MembersCSV '.\DistributionListMembers_contoso.com_20261101_120000.csv' -TargetTenant 'contoso' -CreateCloudOnly -SkipExisting -WhatIf
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ })]
    [string]$DistributionListsCSV,

    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ })]
    [string]$MembersCSV,

    [Parameter(Mandatory = $true)]
    [string]$TargetTenant,

    [switch]$CreateCloudOnly,
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

# Import CSVs
Write-Host "Importing CSVs..." -ForegroundColor Cyan
$groups = Import-Csv -Path $DistributionListsCSV
$members = Import-Csv -Path $MembersCSV

Write-Host "Found $($groups.Count) groups to create" -ForegroundColor Green
Write-Host "Found $($members.Count) memberships to add" -ForegroundColor Green
Write-Host "Source domain(s) will be rewritten to: $targetDomain`n" -ForegroundColor Yellow

# Track results
$results = @{
    GroupsCreated = 0
    GroupsSkipped = 0
    GroupsFailed = 0
    MembersAdded = 0
    MembersFailed = 0
}

$failedGroups = [System.Collections.Generic.List[object]]::new()
$failedMembers = [System.Collections.Generic.List[object]]::new()

# Create groups
Write-Host "Creating distribution groups..." -ForegroundColor Cyan
$i = 0
foreach ($group in $groups) {
    $i++
    Write-Progress -Activity "Creating distribution groups" -Status $group.DisplayName -PercentComplete (($i / $groups.Count) * 100)

    # Skip dynamic groups
    if ($group.GroupType -eq 'DynamicDistributionGroup') {
        Write-Warning "Skipping dynamic group: $($group.DisplayName) (RecipientFilter not supported in import)"
        $results.GroupsSkipped++
        continue
    }

    # Skip dir-synced groups unless -CreateCloudOnly is specified
    if ($group.IsDirSynced -eq 'True' -and -not $CreateCloudOnly) {
        Write-Warning "Skipping dir-synced group: $($group.DisplayName) (use -CreateCloudOnly to create as cloud-only)"
        $results.GroupsSkipped++
        continue
    }

    # Rewrite email addresses to target domain
    $targetEmail = Convert-ToTargetDomain -Email $group.PrimarySmtpAddress
    $targetAlias = $group.Alias

    # Check if group already exists
    $existing = Get-DistributionGroup -Identity $targetEmail -ErrorAction SilentlyContinue
    if ($existing) {
        if ($SkipExisting) {
            Write-Host "  ✓ Skipping existing group: $($group.DisplayName)" -ForegroundColor Yellow
            $results.GroupsSkipped++
            continue
        } else {
            Write-Warning "Group already exists: $($group.DisplayName)"
            $results.GroupsFailed++
            $failedGroups.Add([pscustomobject]@{
                DisplayName = $group.DisplayName
                OriginalEmail = $group.PrimarySmtpAddress
                TargetEmail = $targetEmail
                Reason = 'Already exists'
            })
            continue
        }
    }

    try {
        # Determine group type
        $isSecurityGroup = $group.GroupType -like '*Security*'

        # Parse email addresses (additional aliases) and rewrite to target domain
        $emailAddresses = @()
        if ($group.EmailAddresses) {
            $emailAddresses = $group.EmailAddresses -split '; ' |
                Where-Object { $_ -and $_ -ne $group.PrimarySmtpAddress } |
                ForEach-Object { Convert-ToTargetDomain -Email $_ }
        }

        # Create the group
        $createParams = @{
            Name = $group.DisplayName
            Alias = $targetAlias
            PrimarySmtpAddress = $targetEmail
            Type = if ($isSecurityGroup) { 'Security' } else { 'Distribution' }
        }

        if ($WhatIf) {
            Write-Host "  [WHATIF] Would create: $($group.DisplayName)" -ForegroundColor Cyan
            Write-Host "            Source: $($group.PrimarySmtpAddress)" -ForegroundColor DarkGray
            Write-Host "            Target: $targetEmail" -ForegroundColor Cyan
        } else {
            $newGroup = New-DistributionGroup @createParams -ErrorAction Stop

            # Set additional properties
            $setParams = @{
                Identity = $newGroup.Identity
                RequireSenderAuthenticationEnabled = [bool]::Parse($group.RequireSenderAuthenticationEnabled)
                HiddenFromAddressListsEnabled = [bool]::Parse($group.HiddenFromAddressListsEnabled)
            }

            # Add ManagedBy if present
            if ($group.ManagedBy) {
                $managers = $group.ManagedBy -split '; ' | Where-Object { $_ }
                if ($managers.Count -gt 0) {
                    # Only add managers that exist in the target tenant
                    $validManagers = @()
                    foreach ($mgr in $managers) {
                        if (Get-Recipient -Identity $mgr -ErrorAction SilentlyContinue) {
                            $validManagers += $mgr
                        } else {
                            Write-Warning "  Manager not found in target tenant: $mgr"
                        }
                    }
                    if ($validManagers.Count -gt 0) {
                        $setParams.ManagedBy = $validManagers
                    }
                }
            }

            Set-DistributionGroup @setParams -ErrorAction Stop

            # Add additional email addresses
            if ($emailAddresses.Count -gt 0) {
                foreach ($addr in $emailAddresses) {
                    Set-DistributionGroup -Identity $newGroup.Identity -EmailAddresses @{Add="smtp:$addr"} -ErrorAction SilentlyContinue
                }
            }

            Write-Host "  ✓ Created: $($group.DisplayName) ($targetEmail)" -ForegroundColor Green
            $results.GroupsCreated++
        }
    } catch {
        Write-Warning "Failed to create group $($group.DisplayName): $($_.Exception.Message)"
        $results.GroupsFailed++
        $failedGroups.Add([pscustomobject]@{
            DisplayName = $group.DisplayName
            OriginalEmail = $group.PrimarySmtpAddress
            TargetEmail = $targetEmail
            Reason = $_.Exception.Message
        })
    }
}
Write-Progress -Activity "Creating distribution groups" -Completed

# Add members
Write-Host "`nAdding group members..." -ForegroundColor Cyan

# Group members by group email for efficient processing
$membersByGroup = $members | Group-Object -Property GroupEmail

$i = 0
foreach ($groupMembers in $membersByGroup) {
    $i++
    $sourceGroupEmail = $groupMembers.Name
    $groupName = $groupMembers.Group[0].GroupDisplayName

    # Rewrite group email to target domain
    $targetGroupEmail = Convert-ToTargetDomain -Email $sourceGroupEmail

    Write-Progress -Activity "Adding group members" -Status $groupName -PercentComplete (($i / $membersByGroup.Count) * 100)

    # Check if group exists (might have been skipped or failed)
    $targetGroup = Get-DistributionGroup -Identity $targetGroupEmail -ErrorAction SilentlyContinue
    if (-not $targetGroup) {
        Write-Warning "  Group not found (skipped or failed): $groupName"
        foreach ($m in $groupMembers.Group) {
            $results.MembersFailed++
        }
        continue
    }

    foreach ($member in $groupMembers.Group) {
        try {
            # Rewrite member email to target domain
            $targetMemberEmail = Convert-ToTargetDomain -Email $member.MemberEmail

            # Check if member exists in target tenant
            $targetMember = Get-Recipient -Identity $targetMemberEmail -ErrorAction SilentlyContinue
            if (-not $targetMember) {
                Write-Warning "  Member not found in target tenant: $targetMemberEmail (group: $groupName)"
                $results.MembersFailed++
                $failedMembers.Add([pscustomobject]@{
                    Group = $groupName
                    GroupEmail = $targetGroupEmail
                    OriginalMember = $member.MemberEmail
                    TargetMember = $targetMemberEmail
                    Reason = 'Member not found in target tenant'
                })
                continue
            }

            # Check if member is already in group
            $existing = Get-DistributionGroupMember -Identity $targetGroupEmail -ErrorAction SilentlyContinue |
                Where-Object { $_.PrimarySmtpAddress -eq $targetMemberEmail }

            if ($existing) {
                # Already a member, skip
                continue
            }

            if ($WhatIf) {
                Write-Host "  [WHATIF] Would add: $targetMemberEmail -> $groupName" -ForegroundColor Cyan
            } else {
                Add-DistributionGroupMember -Identity $targetGroupEmail -Member $targetMemberEmail -ErrorAction Stop
                $results.MembersAdded++
            }
        } catch {
            Write-Warning "  Failed to add $targetMemberEmail to $groupName : $($_.Exception.Message)"
            $results.MembersFailed++
            $failedMembers.Add([pscustomobject]@{
                Group = $groupName
                GroupEmail = $targetGroupEmail
                OriginalMember = $member.MemberEmail
                TargetMember = $targetMemberEmail
                Reason = $_.Exception.Message
            })
        }
    }
}
Write-Progress -Activity "Adding group members" -Completed

# Summary
Write-Host "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "IMPORT SUMMARY" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host ""
Write-Host "Groups:" -ForegroundColor White
Write-Host "  Created:  $($results.GroupsCreated)" -ForegroundColor Green
Write-Host "  Skipped:  $($results.GroupsSkipped)" -ForegroundColor Yellow
Write-Host "  Failed:   $($results.GroupsFailed)" -ForegroundColor Red
Write-Host ""
Write-Host "Members:" -ForegroundColor White
Write-Host "  Added:    $($results.MembersAdded)" -ForegroundColor Green
Write-Host "  Failed:   $($results.MembersFailed)" -ForegroundColor Red
Write-Host ""

# Export failed items to CSVs for review
if ($failedGroups.Count -gt 0) {
    $failedGroupsFile = Join-Path (Split-Path $DistributionListsCSV) "FailedGroups_${targetDomain}_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $failedGroups | Export-Csv -Path $failedGroupsFile -NoTypeInformation -Encoding UTF8
    Write-Host "Failed groups exported to: $failedGroupsFile" -ForegroundColor Yellow
}

if ($failedMembers.Count -gt 0) {
    $failedMembersFile = Join-Path (Split-Path $MembersCSV) "FailedMembers_${targetDomain}_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $failedMembers | Export-Csv -Path $failedMembersFile -NoTypeInformation -Encoding UTF8
    Write-Host "Failed members exported to: $failedMembersFile" -ForegroundColor Yellow
}

Write-Host "`nImport complete!" -ForegroundColor Green
