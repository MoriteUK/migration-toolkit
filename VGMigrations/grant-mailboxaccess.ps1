#Requires -Version 7.0
#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Ensures every "SM AEP*" shared mailbox has a matching "MB AEP <name> Mailbox Access"
    security group, and that the group has Full Access + Send As on the mailbox.

.DESCRIPTION
    1. Finds all shared mailboxes whose -MatchProperty (DisplayName by default — Exchange's
       internal Name/Alias often doesn't match what's shown in Outlook/EAC) starts with
       -MailboxPrefix (default "SM AEP").
    2. For each one, builds the access group name by stripping the prefix off that same
       property and wrapping it as "<GroupPrefix> <remainder> <GroupSuffix>"
       e.g. mailbox DisplayName "SM AEP Finance" -> group "MB AEP Finance Mailbox Access".
    3. Creates the group (mail-enabled security group) if it doesn't already exist.
       A mail-enabled security group is required here, not a plain Entra ID security
       group, because Add-RecipientPermission (Send As) only accepts a mail-enabled
       trustee.
    4. Grants Full Access and Send As on the mailbox to the group, if not already granted.

    Everything is idempotent — safe to re-run. New mailboxes matching the prefix will
    pick up a group automatically on the next run; existing groups/permissions are
    left alone.

    Dry-run by default. Nothing is created or changed unless you pass -Commit.

.PARAMETER MailboxPrefix
    Prefix used to filter shared mailboxes. Default 'SM AEP'.

.PARAMETER MatchProperty
    Which mailbox property to match -MailboxPrefix against and to derive the group name's
    remainder from: 'DisplayName' (default) or 'Name'. Exchange's Name/Alias is an internal
    identifier and frequently differs from what people actually see and named things by
    (DisplayName) — if the script reports zero matches but you can see plenty of "SM AEP..."
    mailboxes in the EAC/Outlook, they're almost certainly named that way in DisplayName, not
    Name, which is why this defaults to DisplayName.

.PARAMETER GroupPrefix
    Prefix used when building the access group name. Default 'MB AEP'.

.PARAMETER GroupSuffix
    Suffix used when building the access group name. Default 'Mailbox Access'.

.PARAMETER AutoMapping
    Whether Full Access grants auto-map the mailbox into group members' Outlook.
    Default $false — with a group (rather than a single user) getting Full Access,
    leaving AutoMapping on tends to flood every member's Outlook profile with the
    mailbox whether they want it there or not. Members can still add it manually,
    or you can flip this to $true if you want it to just appear for everyone.

.PARAMETER Commit
    Without this switch the script only reports what it *would* do. Pass -Commit to
    actually create groups and grant permissions.

.EXAMPLE
    ./New-AEPMailboxAccessGroups.ps1
    Dry run — shows which groups/permissions are missing, writes a CSV report.

.EXAMPLE
    ./New-AEPMailboxAccessGroups.ps1 -Commit
    Creates the missing groups and grants the missing permissions.

.NOTES
    Author: generated for Andy White
    Requires the ExchangeOnlineManagement module and an account with Exchange admin
    rights (mailbox permissions) + a role that can create distribution/security
    groups (Recipient Management or similar).
#>

[CmdletBinding()]
param(
    [string]$MailboxPrefix = 'SM AEP',
    [ValidateSet('DisplayName', 'Name')]
    [string]$MatchProperty = 'DisplayName',
    [string]$GroupPrefix   = 'MB AEP',
    [string]$GroupSuffix   = 'Mailbox Access',
    [bool]$AutoMapping     = $false,
    [switch]$Commit
)

$ErrorActionPreference = 'Stop'
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$logPath   = Join-Path -Path (Get-Location) -ChildPath "AEP-MailboxAccessGroups_$timestamp.csv"
$results   = [System.Collections.Generic.List[pscustomobject]]::new()

function Write-Status {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'ACTION', 'OK', 'ERROR')]
        [string]$Level = 'INFO'
    )
    $colour = switch ($Level) {
        'ACTION' { 'Yellow' }
        'OK'     { 'Green' }
        'ERROR'  { 'Red' }
        default  { 'Gray' }
    }
    Write-Host "[$Level] $Message" -ForegroundColor $colour
}

# --- Connect to Exchange Online ---------------------------------------------
if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
    Write-Status 'Connecting to Exchange Online...'
    try {
        Connect-ExchangeOnline -ShowBanner:$false
    }
    catch {
        Write-Status "Standard sign-in failed ($($_.Exception.Message)), retrying with device code login..." 'ACTION'
        Connect-ExchangeOnline -ShowBanner:$false -Device
    }
}

if (-not $Commit) {
    Write-Status "Running in DRY-RUN mode — nothing will be created or changed. Re-run with -Commit to apply." 'ACTION'
}

# --- Find target shared mailboxes -------------------------------------------
Write-Status "Looking up shared mailboxes whose $MatchProperty matches '$MailboxPrefix*'..."
$sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited |
    Where-Object { $_.$MatchProperty -like "$MailboxPrefix*" } |
    Sort-Object $MatchProperty

if (-not $sharedMailboxes) {
    Write-Status "No shared mailboxes found where $MatchProperty starts with '$MailboxPrefix'. Try -MatchProperty Name if these mailboxes were matched by DisplayName instead." 'ERROR'
    return
}

Write-Status "Found $($sharedMailboxes.Count) matching mailbox(es)."

$defaultDomain = (Get-AcceptedDomain | Where-Object { $_.Default }).DomainName

foreach ($mbx in $sharedMailboxes) {

    # "SM AEP Finance" -> remainder "Finance" -> group "MB AEP Finance Mailbox Access"
    $matchValue = $mbx.$MatchProperty
    $suffix     = ($matchValue -replace "^$([regex]::Escape($MailboxPrefix))", '').Trim()
    $groupName  = ("$GroupPrefix $suffix $GroupSuffix" -replace '\s{2,}', ' ').Trim()
    $alias      = ($groupName -replace '[^a-zA-Z0-9]', '')
    if ($alias.Length -gt 64) { $alias = $alias.Substring(0, 64) }

    $row = [pscustomobject]@{
        Mailbox          = $matchValue
        MailboxUPN       = $mbx.PrimarySmtpAddress
        GroupName        = $groupName
        GroupAction      = 'None'
        FullAccessAction = 'None'
        SendAsAction     = 'None'
        Status           = 'OK'
        Notes            = ''
    }

    try {
        # --- Group: create if missing -------------------------------------
        $group = Get-DistributionGroup -Identity $groupName -ErrorAction SilentlyContinue

        if (-not $group) {
            $row.GroupAction = 'Create'
            Write-Status "$groupName : group does not exist" 'ACTION'
            if ($Commit) {
                $group = New-DistributionGroup -Name $groupName -Alias $alias -Type Security `
                    -PrimarySmtpAddress "$alias@$defaultDomain"
                Set-DistributionGroup -Identity $group.Identity `
                    -HiddenFromAddressListsEnabled $true `
                    -RequireSenderAuthenticationEnabled $true
                Write-Status "$groupName : group created (hidden from GAL, security-only)" 'OK'
            }
        }
        else {
            Write-Status "$groupName : group already exists" 'OK'
        }

        # --- Full Access ----------------------------------------------------
        $hasFullAccess = $false
        if ($group) {
            $hasFullAccess = [bool](Get-MailboxPermission -Identity $mbx.Identity -ErrorAction SilentlyContinue |
                Where-Object { $_.User -like "*$($group.Name)*" -and $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited })
        }

        if (-not $hasFullAccess) {
            $row.FullAccessAction = 'Grant'
            Write-Status "$matchValue : Full Access for '$groupName' missing" 'ACTION'
            if ($Commit -and $group) {
                Add-MailboxPermission -Identity $mbx.Identity -User $group.Identity `
                    -AccessRights FullAccess -InheritanceType All -AutoMapping:$AutoMapping -Confirm:$false | Out-Null
                Write-Status "$matchValue : Full Access granted to '$groupName'" 'OK'
            }
        }
        else {
            Write-Status "$matchValue : Full Access already granted to '$groupName'" 'OK'
        }

        # --- Send As ----------------------------------------------------------
        $hasSendAs = $false
        if ($group) {
            $hasSendAs = [bool](Get-RecipientPermission -Identity $mbx.Identity -Trustee $group.Identity -ErrorAction SilentlyContinue)
        }

        if (-not $hasSendAs) {
            $row.SendAsAction = 'Grant'
            Write-Status "$matchValue : Send As for '$groupName' missing" 'ACTION'
            if ($Commit -and $group) {
                Add-RecipientPermission -Identity $mbx.Identity -Trustee $group.Identity -AccessRights SendAs -Confirm:$false | Out-Null
                Write-Status "$matchValue : Send As granted to '$groupName'" 'OK'
            }
        }
        else {
            Write-Status "$matchValue : Send As already granted to '$groupName'" 'OK'
        }
    }
    catch {
        $row.Status = 'ERROR'
        $row.Notes  = $_.Exception.Message
        Write-Status "$matchValue : $($_.Exception.Message)" 'ERROR'
    }

    $results.Add($row)
}

$results | Export-Csv -Path $logPath -NoTypeInformation
Write-Status "Report written to $logPath" 'OK'

if (-not $Commit) {
    Write-Status 'Dry run complete. Re-run with -Commit to create the groups / grant the permissions listed above.' 'ACTION'
}