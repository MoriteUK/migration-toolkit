#Requires -Version 7.0

Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -DisableNameChecking -Force -Global

$script:GraphSubmodules = @(
    'Microsoft.Graph.Authentication'
    'Microsoft.Graph.Groups'
    'Microsoft.Graph.Teams'
    'Microsoft.Graph.Users'
    'Microsoft.Graph.DeviceManagement'
)

# -----------------------------------------------------------------------
# Private functions
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Installs missing PSGallery modules for the current user, bootstrapping the NuGet provider if needed.
#>
function Install-Prerequisites {
    param([string[]]$ModuleNames)

    $nuget = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue
    if (-not $nuget) {
        Write-Host ($PREFIX_INFO + 'Bootstrapping NuGet provider...') -ForegroundColor DarkGray
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
    }

    foreach ($name in $ModuleNames) {
        Write-Host ($PREFIX_INFO + "Installing $name...") -ForegroundColor DarkGray
        try {
            Install-Module -Name $name -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
            Write-Host ($PREFIX_OK + "$name installed") -ForegroundColor Green
        }
        catch {
            Write-Host ($PREFIX_FAIL + "Failed to install ${name}: " + $_.Exception.Message) -ForegroundColor Red
            throw
        }
    }
}

# -----------------------------------------------------------------------
# Public functions
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Validates required modules and AD connectivity before an assessment run.
.DESCRIPTION
    Hard-fails if the ActiveDirectory module (RSAT) is missing or the domain is
    unreachable. Missing PSGallery modules are installed via Install-Prerequisites;
    a failed SharePoint module install is non-fatal and sets SkipSharePoint on the
    context instead.
#>
function Test-Prerequisites {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSCustomObject]$Context
    )

    Write-SectionHeader 'Prerequisites'

    # --- ActiveDirectory ---
    # Cannot be installed from PSGallery - requires RSAT via Windows optional features
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-Host ($PREFIX_FAIL + 'ActiveDirectory module not found.') -ForegroundColor Red
        Write-Host ($PREFIX_INFO + 'Install RSAT via Settings -> Apps -> Optional Features -> RSAT: Active Directory Domain Services and Lightweight Directory Services Tools') -ForegroundColor Yellow
        throw 'ActiveDirectory module not found. Install RSAT to continue.'
    }

    try {
        Get-ADDomain -ErrorAction Stop | Out-Null
        Write-Host ($PREFIX_OK + 'ActiveDirectory module present and domain reachable') -ForegroundColor Green
    }
    catch {
        Write-Host ($PREFIX_FAIL + 'AD connectivity check failed: ' + $_.Exception.Message) -ForegroundColor Red
        throw
    }

    # --- ExchangeOnlineManagement ---
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host ($PREFIX_WARN + 'ExchangeOnlineManagement not found - installing...') -ForegroundColor Yellow
        Install-Prerequisites -ModuleNames @('ExchangeOnlineManagement')
    }
    else {
        Write-Host ($PREFIX_OK + 'ExchangeOnlineManagement present') -ForegroundColor Green
    }

    # --- Microsoft.Graph submodules ---
    # Install individual submodules - the meta-package is too large for Install-Module
    $missing = @($script:GraphSubmodules | Where-Object { -not (Get-Module -ListAvailable -Name $_) })
    if ($missing.Count -gt 0) {
        Write-Host ($PREFIX_WARN + "Missing Graph submodules ($($missing.Count)) - installing...") -ForegroundColor Yellow
        Install-Prerequisites -ModuleNames $missing
    }
    else {
        Write-Host ($PREFIX_OK + 'Microsoft.Graph submodules present') -ForegroundColor Green
    }

    # --- Microsoft.Online.SharePoint.PowerShell ---
    # Install failure is non-fatal - sets SkipSharePoint on the context
    if (-not (Get-Module -ListAvailable -Name 'Microsoft.Online.SharePoint.PowerShell')) {
        Write-Host ($PREFIX_WARN + 'Microsoft.Online.SharePoint.PowerShell not found - installing...') -ForegroundColor Yellow
        try {
            Install-Prerequisites -ModuleNames @('Microsoft.Online.SharePoint.PowerShell')
        }
        catch {
            Write-Host ($PREFIX_WARN + 'SPO module install failed - SharePoint collection will be skipped') -ForegroundColor Yellow
            $Context.SkipSharePoint = $true
        }
    }
    else {
        Write-Host ($PREFIX_OK + 'Microsoft.Online.SharePoint.PowerShell present') -ForegroundColor Green
    }

    # --- ImportExcel ---
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host ($PREFIX_WARN + 'ImportExcel not found - installing...') -ForegroundColor Yellow
        Install-Prerequisites -ModuleNames @('ImportExcel')
    }
    else {
        Write-Host ($PREFIX_OK + 'ImportExcel present') -ForegroundColor Green
    }

    Write-Host ''
    Write-Host ($PREFIX_OK + 'Prerequisites check complete') -ForegroundColor Green
}

# -----------------------------------------------------------------------
# Exports
# -----------------------------------------------------------------------

Export-ModuleMember -Function 'Test-Prerequisites'
