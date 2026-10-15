<#
.SYNOPSIS
    Ensure-GraphModules.ps1 — dot-source this before any Connect-MgGraph / Mg* cmdlet call.

.DESCRIPTION
    Microsoft Graph SDK >= 2.34.0 makes the WAM (Web Account Manager) broker mandatory for
    interactive sign-in on Windows. Set-MgGraphOption -DisableLoginByWAM is a documented no-op
    unless paired with a custom app registration
    (github.com/microsoftgraph/msgraph-sdk-powershell/issues/3518) — it does NOT actually stop
    WAM from being tried. WAM failures used to be caught and retried with
    -UseDeviceCode/-UseDeviceAuthentication, but device-code flow can now be blocked tenant-wide
    by a Conditional Access "Authentication flows" policy, so relying on device-code — whether as
    a fallback or as the sole/primary sign-in method — is no longer reliable.

    Pinning every Microsoft.Graph.* submodule a script uses to 2.33.0 — the last version before
    WAM became mandatory — restores plain browser-popup interactive sign-in. ALL submodules a
    script touches must be pinned together (not just Authentication), otherwise mismatched
    assembly versions get loaded into the same process, which is its own source of the
    NullReferenceException/RuntimeBroker failures this codebase has chased for over a year.
    Installs side-by-side if missing; does not remove or affect any newer version already
    installed on the machine.

.PARAMETER GraphModules
    Microsoft.Graph.* submodule names this script calls cmdlets from, beyond Authentication
    (which is always pinned). E.g. @('Microsoft.Graph.Users', 'Microsoft.Graph.Groups').

.EXAMPLE
    . (Join-Path $PSScriptRoot 'Ensure-GraphModules.ps1') -GraphModules @('Microsoft.Graph.Users')
#>
param(
    [string[]]$GraphModules = @()
)

$Script:GraphPinnedVersion = '2.33.0'
$modulesToPin = @('Microsoft.Graph.Authentication') + $GraphModules | Select-Object -Unique

foreach ($mod in $modulesToPin) {
    try {
        if (-not (Get-Module -ListAvailable -Name $mod | Where-Object { $_.Version -eq $Script:GraphPinnedVersion })) {
            Write-Host "Installing $mod $Script:GraphPinnedVersion (pinned — avoids the SDK's mandatory-WAM regression in >= 2.34.0)..." -ForegroundColor Yellow
            Install-Module -Name $mod -RequiredVersion $Script:GraphPinnedVersion -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
        Import-Module -Name $mod -RequiredVersion $Script:GraphPinnedVersion -Force -ErrorAction Stop
    } catch {
        Write-Host "WARNING: Could not load pinned ${mod} ${Script:GraphPinnedVersion} — $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "Falling back to whatever version of $mod is already installed — interactive sign-in may hit the WAM/device-code issue this pin exists to avoid." -ForegroundColor Yellow
    }
}
