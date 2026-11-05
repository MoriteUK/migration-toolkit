#Requires -Version 7.0
<#
.SYNOPSIS
    Removes all references to a specified domain from on-premise Active Directory objects.

.DESCRIPTION
    Searches for every occurrence of a domain across AD user accounts, groups, contacts,
    and removes it from proxyAddresses, mail, and targetAddress attributes.
    If the domain is the primary address, it gets changed to @ourvolaris.onmicrosoft.com.

    Processes:
    - AD Users (UserPrincipalName, proxyAddresses, mail, targetAddress)
    - AD Groups (proxyAddresses, mail) - mail-enabled groups only
    - AD Contacts (proxyAddresses, mail, targetAddress)

.PARAMETER Domain
    The domain to remove (e.g., oldcompany.com). Do NOT include @ symbol.

.PARAMETER ReplacementDomain
    The replacement domain for primary addresses. Default: ourvolaris.onmicrosoft.com

.PARAMETER WhatIf
    Preview changes without making them.

.PARAMETER SkipUserPrincipalNames
    Don't change user UPNs, only remove from proxyAddresses.

.PARAMETER SearchBase
    Optional OU to limit search scope (e.g., "OU=Users,DC=domain,DC=com")

.PARAMETER Server
    Optional DC to connect to.

.PARAMETER Credential
    Optional credentials for AD access. If not specified, uses current Windows credentials.

.EXAMPLE
    .\Remove-DomainReferences.ps1 -Domain oldcompany.com -WhatIf
    Preview what would be removed

.EXAMPLE
    .\Remove-DomainReferences.ps1 -Domain oldcompany.com
    Remove all references to oldcompany.com from AD

.EXAMPLE
    .\Remove-DomainReferences.ps1 -Domain oldcompany.com -ReplacementDomain newtenant.onmicrosoft.com
    Use custom replacement domain

.EXAMPLE
    .\Remove-DomainReferences.ps1 -Domain oldcompany.com -SearchBase "OU=Users,DC=contoso,DC=com"
    Only process objects in specific OU

.EXAMPLE
    .\Remove-DomainReferences.ps1 -Domain oldcompany.com -Server DC01.contoso.local
    Use specific DC

.EXAMPLE
    $cred = Get-Credential
    .\Remove-DomainReferences.ps1 -Domain oldcompany.com -Credential $cred
    Use alternate credentials
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$Domain,

    [string]$ReplacementDomain = 'ourvolaris.onmicrosoft.com',

    [switch]$SkipUserPrincipalNames,

    [string]$SearchBase,

    [string]$Server,

    [PSCredential]$Credential
)

$ErrorActionPreference = 'Stop'

# Import Active Directory module
Write-Host 'Loading Active Directory module...' -ForegroundColor Cyan
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Write-Host '  ✓ Active Directory module loaded' -ForegroundColor Green
}
catch {
    Write-Host '  ✗ Failed to load Active Directory module' -ForegroundColor Red
    Write-Host '    Install RSAT tools: Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0' -ForegroundColor Yellow
    throw
}

# Clean domain input (remove @ if present)
$Domain = $Domain.TrimStart('@').ToLowerInvariant()
$ReplacementDomain = $ReplacementDomain.TrimStart('@').ToLowerInvariant()

# Test AD connection
Write-Host 'Testing Active Directory connection...' -ForegroundColor Cyan
try {
    $testParams = @{}
    if ($Server) { $testParams.Server = $Server }
    if ($Credential) { $testParams.Credential = $Credential }

    $testDomain = Get-ADDomain @testParams -ErrorAction Stop
    Write-Host "  ✓ Connected to domain: $($testDomain.DNSRoot)" -ForegroundColor Green
    Write-Host "  ✓ Domain Controller: $($testDomain.PDCEmulator)" -ForegroundColor Green
}
catch {
    Write-Host '  ✗ Cannot connect to Active Directory' -ForegroundColor Red
    Write-Host "    Error: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host ''
    Write-Host 'Troubleshooting:' -ForegroundColor Yellow
    Write-Host '  1. Are you on a domain-joined machine?' -ForegroundColor DarkGray
    Write-Host '  2. Do you have RSAT Active Directory tools installed?' -ForegroundColor DarkGray
    Write-Host '  3. Try specifying -Server <DomainController> and/or -Credential' -ForegroundColor DarkGray
    throw
}

Write-Host ''
Write-Host '╔═══════════════════════════════════════════════════════════╗' -ForegroundColor Cyan
Write-Host '║     Remove Domain References from Active Directory       ║' -ForegroundColor Cyan
Write-Host '╚═══════════════════════════════════════════════════════════╝' -ForegroundColor Cyan
Write-Host ''
Write-Host "Domain to remove:       @$Domain" -ForegroundColor Yellow
Write-Host "Replacement domain:     @$ReplacementDomain" -ForegroundColor Green
Write-Host "WhatIf mode:            $($WhatIfPreference)" -ForegroundColor $(if ($WhatIfPreference) { 'Yellow' } else { 'White' })
if ($SearchBase) {
    Write-Host "Search base:            $SearchBase" -ForegroundColor White
}
if ($Server) {
    Write-Host "Domain Controller:      $Server" -ForegroundColor White
}
Write-Host ''

# Build AD cmdlet parameters
$adParams = @{}
if ($SearchBase) { $adParams.SearchBase = $SearchBase }
if ($Server) { $adParams.Server = $Server }
if ($Credential) { $adParams.Credential = $Credential }

# Statistics
$stats = @{
    UsersProcessed = 0
    UsersChanged = 0
    GroupsProcessed = 0
    GroupsChanged = 0
    ContactsProcessed = 0
    ContactsChanged = 0
    TotalProxyAddressesRemoved = 0
    TotalUPNsChanged = 0
    TotalMailAttributesChanged = 0
    TotalTargetAddressesChanged = 0
}

# Helper function to process AD object
function Remove-DomainFromADObject {
    param(
        [Parameter(Mandatory)]
        $ADObject,

        [Parameter(Mandatory)]
        [string]$ObjectType,

        [Parameter(Mandatory)]
        [string]$DomainToRemove,

        [Parameter(Mandatory)]
        [string]$ReplacementDomain,

        [switch]$ChangeUPN
    )

    $changed = $false
    $dn = $ADObject.DistinguishedName

    Write-Host ''
    Write-Host "  Processing: $($ADObject.Name) ($($ADObject.mail))" -ForegroundColor White

    # ── Process proxyAddresses ──────────────────────────────────────────────
    $currentProxies = @($ADObject.proxyAddresses)
    if ($currentProxies.Count -gt 0) {
        $newProxies = [System.Collections.Generic.List[string]]::new()
        $removedCount = 0
        $primaryChanged = $false

        foreach ($proxy in $currentProxies) {
            $proxyLower = $proxy.ToLowerInvariant()

            # Keep addresses that DON'T match the domain (or are X500)
            if ($proxyLower -notlike "*@$DomainToRemove" -or $proxyLower -like 'x500:*') {
                $newProxies.Add($proxy)
            }
            else {
                Write-Host "    Removing proxyAddress: $proxy" -ForegroundColor Red
                $removedCount++
                $changed = $true
                $script:stats.TotalProxyAddressesRemoved++

                # Check if we're removing the primary
                if ($proxy -clike 'SMTP:*') {
                    $primaryChanged = $true
                }
            }
        }

        # If primary was removed, add new primary
        if ($primaryChanged) {
            $samAccountName = $ADObject.SamAccountName
            $newPrimary = "SMTP:$samAccountName@$ReplacementDomain"
            $newProxies.Add($newPrimary)
            Write-Host "    New PRIMARY: $newPrimary" -ForegroundColor Green
        }

        # Update proxyAddresses if changed
        if ($removedCount -gt 0) {
            if ($PSCmdlet.ShouldProcess($dn, "Remove $removedCount proxyAddress(es)")) {
                try {
                    Set-ADObject -Identity $dn @adParams `
                                -Replace @{ proxyAddresses = $newProxies.ToArray() } `
                                -ErrorAction Stop
                    Write-Host "    ✓ Updated proxyAddresses" -ForegroundColor Green
                }
                catch {
                    Write-Host "    ✗ proxyAddresses Error: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            else {
                Write-Host "    [WhatIf] Would remove $removedCount proxyAddress(es)" -ForegroundColor Yellow
            }
        }
    }

    # ── Process mail attribute ───────────────────────────────────────────────
    if ($ADObject.mail -and $ADObject.mail -like "*@$DomainToRemove") {
        $newMail = "$($ADObject.SamAccountName)@$ReplacementDomain"
        Write-Host "    Changing mail: $($ADObject.mail) → $newMail" -ForegroundColor Yellow

        if ($PSCmdlet.ShouldProcess($dn, "Change mail to $newMail")) {
            try {
                Set-ADObject -Identity $dn @adParams `
                            -Replace @{ mail = $newMail } `
                            -ErrorAction Stop
                Write-Host "    ✓ Updated mail" -ForegroundColor Green
                $changed = $true
                $script:stats.TotalMailAttributesChanged++
            }
            catch {
                Write-Host "    ✗ mail Error: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        else {
            Write-Host "    [WhatIf] Would change mail" -ForegroundColor Yellow
        }
    }

    # ── Process targetAddress ───────────────────────────────────────────────
    if ($ADObject.targetAddress -and $ADObject.targetAddress -like "*@$DomainToRemove") {
        $newTarget = "$($ADObject.SamAccountName)@$ReplacementDomain"
        Write-Host "    Changing targetAddress: $($ADObject.targetAddress) → $newTarget" -ForegroundColor Yellow

        if ($PSCmdlet.ShouldProcess($dn, "Change targetAddress to $newTarget")) {
            try {
                Set-ADObject -Identity $dn @adParams `
                            -Replace @{ targetAddress = $newTarget } `
                            -ErrorAction Stop
                Write-Host "    ✓ Updated targetAddress" -ForegroundColor Green
                $changed = $true
                $script:stats.TotalTargetAddressesChanged++
            }
            catch {
                Write-Host "    ✗ targetAddress Error: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        else {
            Write-Host "    [WhatIf] Would change targetAddress" -ForegroundColor Yellow
        }
    }

    # ── Process UserPrincipalName (Users only) ──────────────────────────────
    if ($ChangeUPN -and -not $SkipUserPrincipalNames -and $ObjectType -eq 'User' -and $ADObject.UserPrincipalName) {
        if ($ADObject.UserPrincipalName -like "*@$DomainToRemove") {
            $newUPN = "$($ADObject.SamAccountName)@$ReplacementDomain"
            Write-Host "    Changing UPN: $($ADObject.UserPrincipalName) → $newUPN" -ForegroundColor Yellow

            if ($PSCmdlet.ShouldProcess($dn, "Change UPN to $newUPN")) {
                try {
                    Set-ADUser -Identity $dn @adParams `
                              -UserPrincipalName $newUPN `
                              -ErrorAction Stop
                    Write-Host "    ✓ Updated UPN" -ForegroundColor Green
                    $changed = $true
                    $script:stats.TotalUPNsChanged++
                }
                catch {
                    Write-Host "    ✗ UPN Error: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            else {
                Write-Host "    [WhatIf] Would change UPN" -ForegroundColor Yellow
            }
        }
    }

    if (-not $changed) {
        Write-Host "    No changes needed" -ForegroundColor DarkGray
    }

    return $changed
}

# ═══════════════════════════════════════════════════════════════════════════
# AD USERS
# ═══════════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host '═══ AD Users ═══' -ForegroundColor Cyan

$filter = {(proxyAddresses -like "*@$Domain") -or (mail -like "*@$Domain") -or (targetAddress -like "*@$Domain") -or (UserPrincipalName -like "*@$Domain")}
$users = @(Get-ADUser -Filter $filter @adParams `
    -Properties Name, SamAccountName, UserPrincipalName, mail, proxyAddresses, targetAddress)

Write-Host "Found $($users.Count) users with references to @$Domain" -ForegroundColor White

foreach ($user in $users) {
    $stats.UsersProcessed++

    if (Remove-DomainFromADObject -ADObject $user -ObjectType 'User' `
                                   -DomainToRemove $Domain `
                                   -ReplacementDomain $ReplacementDomain `
                                   -ChangeUPN) {
        $stats.UsersChanged++
    }
}

# ═══════════════════════════════════════════════════════════════════════════
# AD GROUPS (mail-enabled only)
# ═══════════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host '═══ AD Groups (mail-enabled) ═══' -ForegroundColor Cyan

$filter = {(proxyAddresses -like "*@$Domain") -or (mail -like "*@$Domain")}
$groups = @(Get-ADGroup -Filter $filter @adParams `
    -Properties Name, SamAccountName, mail, proxyAddresses)

Write-Host "Found $($groups.Count) groups with references to @$Domain" -ForegroundColor White

foreach ($group in $groups) {
    $stats.GroupsProcessed++

    if (Remove-DomainFromADObject -ADObject $group -ObjectType 'Group' `
                                   -DomainToRemove $Domain `
                                   -ReplacementDomain $ReplacementDomain) {
        $stats.GroupsChanged++
    }
}

# ═══════════════════════════════════════════════════════════════════════════
# AD CONTACTS
# ═══════════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host '═══ AD Contacts ═══' -ForegroundColor Cyan

$filter = {(proxyAddresses -like "*@$Domain") -or (mail -like "*@$Domain") -or (targetAddress -like "*@$Domain")}
$contacts = @(Get-ADObject -Filter $filter @adParams -Properties objectClass `
    | Where-Object { $_.objectClass -eq 'contact' })

if ($contacts.Count -gt 0) {
    # Get full contact objects with all properties
    $fullContacts = @()
    foreach ($contact in $contacts) {
        $fullContacts += Get-ADObject -Identity $contact.DistinguishedName @adParams `
                            -Properties Name, mail, proxyAddresses, targetAddress
    }

    Write-Host "Found $($fullContacts.Count) contacts with references to @$Domain" -ForegroundColor White

    foreach ($contact in $fullContacts) {
        # Contacts don't have SamAccountName, use first part of mail or name
        if ($contact.mail) {
            $contact | Add-Member -NotePropertyName 'SamAccountName' -NotePropertyValue $contact.mail.Split('@')[0] -Force
        }
        else {
            $contact | Add-Member -NotePropertyName 'SamAccountName' -NotePropertyValue $contact.Name.Replace(' ', '') -Force
        }

        $stats.ContactsProcessed++

        if (Remove-DomainFromADObject -ADObject $contact -ObjectType 'Contact' `
                                       -DomainToRemove $Domain `
                                       -ReplacementDomain $ReplacementDomain) {
            $stats.ContactsChanged++
        }
    }
}
else {
    Write-Host "Found 0 contacts with references to @$Domain" -ForegroundColor White
}

# ═══════════════════════════════════════════════════════════════════════════
# SUMMARY
# ═══════════════════════════════════════════════════════════════════════════
Write-Host ''
Write-Host '╔═══════════════════════════════════════════════════════════╗' -ForegroundColor Green
Write-Host '║                    SUMMARY REPORT                         ║' -ForegroundColor Green
Write-Host '╚═══════════════════════════════════════════════════════════╝' -ForegroundColor Green
Write-Host ''
Write-Host "Domain removed:              @$Domain" -ForegroundColor Yellow
Write-Host "Replacement domain:          @$ReplacementDomain" -ForegroundColor Green
Write-Host ''
Write-Host 'Object Type          Processed    Changed' -ForegroundColor White
Write-Host '────────────────────────────────────────' -ForegroundColor DarkGray
Write-Host ("AD Users             {0,-12} {1}" -f $stats.UsersProcessed, $stats.UsersChanged) -ForegroundColor White
Write-Host ("AD Groups            {0,-12} {1}" -f $stats.GroupsProcessed, $stats.GroupsChanged) -ForegroundColor White
Write-Host ("AD Contacts          {0,-12} {1}" -f $stats.ContactsProcessed, $stats.ContactsChanged) -ForegroundColor White
Write-Host ''
Write-Host "Total proxyAddresses removed:    $($stats.TotalProxyAddressesRemoved)" -ForegroundColor Yellow
Write-Host "Total UPNs changed:              $($stats.TotalUPNsChanged)" -ForegroundColor Yellow
Write-Host "Total mail attributes changed:   $($stats.TotalMailAttributesChanged)" -ForegroundColor Yellow
Write-Host "Total targetAddresses changed:   $($stats.TotalTargetAddressesChanged)" -ForegroundColor Yellow
Write-Host ''

if ($WhatIfPreference) {
    Write-Host '⚠ WhatIf mode - NO CHANGES WERE MADE' -ForegroundColor Yellow
    Write-Host '  Run without -WhatIf to apply changes' -ForegroundColor Yellow
    Write-Host ''
}
