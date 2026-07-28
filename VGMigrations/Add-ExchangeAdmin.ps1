#Requires -Version 7.0
<#
.SYNOPSIS
    Add-ExchangeAdmin.ps1 — Assigns the Exchange Administrator Entra role to a user.

.DESCRIPTION
    Connects to Microsoft Graph and permanently assigns the Exchange Administrator
    directory role to the specified account. This is required before granting Fly
    consent for Exchange migration — the consent will fail if the account does not
    already hold this role.

    If the user already has the role the script exits cleanly with no error.

.PARAMETER UPN
    User Principal Name of the account to assign the role to
    (the same account that will approve the Fly consent).
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$UPN
)

$ErrorActionPreference = 'Stop'

Write-Host "=== Assign Exchange Administrator Role ==="
Write-Host "Account : $UPN"
Write-Host ''

. (Join-Path $PSScriptRoot 'Ensure-GraphModules.ps1') -GraphModules @('Microsoft.Graph.Identity.Governance','Microsoft.Graph.Users')

Write-Host 'Connecting to Microsoft Graph — sign in with a Global Administrator account...'
Connect-MgGraph -Scopes 'RoleManagement.ReadWrite.Directory','User.Read.All' -NoWelcome -ErrorAction Stop
Write-Host 'Connected.'
Write-Host ''

# Resolve user object
Write-Host "Looking up user: $UPN"
$user = Get-MgUser -UserId $UPN -Property 'Id,DisplayName,UserPrincipalName' -ErrorAction Stop
Write-Host "  Found: $($user.DisplayName)  [$($user.Id)]"

# Find Exchange Administrator role definition
Write-Host "Looking up Exchange Administrator role definition..."
$roleDef = Get-MgRoleManagementDirectoryRoleDefinition `
    -Filter "displayName eq 'Exchange Administrator'" -ErrorAction Stop
if (-not $roleDef) {
    Write-Host "ERROR: 'Exchange Administrator' role definition not found in tenant."
    exit 1
}
Write-Host "  Role definition ID: $($roleDef.Id)"

# Check for existing assignment
Write-Host "Checking existing role assignments..."
$existing = Get-MgRoleManagementDirectoryRoleAssignment `
    -Filter "principalId eq '$($user.Id)' and roleDefinitionId eq '$($roleDef.Id)'" `
    -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "  $($user.DisplayName) already has the Exchange Administrator role — nothing to do."
    Write-Host ''
    Write-Host "=== Done (already assigned) ==="
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    exit 0
}

# Assign the role (tenant-wide scope)
Write-Host "Assigning Exchange Administrator to $($user.DisplayName)..."
New-MgRoleManagementDirectoryRoleAssignment `
    -PrincipalId   $user.Id `
    -RoleDefinitionId $roleDef.Id `
    -DirectoryScopeId '/' `
    -ErrorAction Stop

Write-Host "  Assigned OK."
Write-Host ''
Write-Host "=== Done: Exchange Administrator role assigned to $($user.DisplayName) ==="
Write-Host "You can now sign in to AOS and run the App Profile & Consent setup."

try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
