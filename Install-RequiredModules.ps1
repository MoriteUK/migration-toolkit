<#
.SYNOPSIS
  Checks and installs all PowerShell modules required for Migration Toolkit and FLY.

.DESCRIPTION
  This script verifies the installation of all required PowerShell modules and offers
  to install any missing ones. It handles both user-scope and system-wide installations.

.PARAMETER AutoInstall
  Automatically install missing modules without prompting.

.PARAMETER Scope
  Installation scope: CurrentUser (default) or AllUsers (requires admin).

.EXAMPLE
  .\Install-RequiredModules.ps1

.EXAMPLE
  .\Install-RequiredModules.ps1 -AutoInstall -Scope CurrentUser

.EXAMPLE
  .\Install-RequiredModules.ps1 -Scope AllUsers
  # Requires running as Administrator
#>

[CmdletBinding()]
param(
    [switch]$AutoInstall,
    [ValidateSet('CurrentUser', 'AllUsers')]
    [string]$Scope = 'CurrentUser'
)

$ErrorActionPreference = 'Stop'

# Check if running as administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

$isAdmin = Test-Administrator

if ($Scope -eq 'AllUsers' -and -not $isAdmin) {
    Write-Error "Installing modules with -Scope AllUsers requires running PowerShell as Administrator."
    exit 1
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Migration Toolkit - Module Installation  " -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Scope: $Scope" -ForegroundColor Gray
if ($isAdmin) {
    Write-Host "Running as: Administrator" -ForegroundColor Green
} else {
    Write-Host "Running as: Current User" -ForegroundColor Yellow
}
Write-Host ""

# Define all required modules with descriptions
$requiredModules = @(
    @{
        Name = 'Microsoft.Graph'
        Description = 'Microsoft Graph PowerShell SDK for Azure AD, Users, Groups, etc.'
        Required = $true
        MinVersion = '2.0.0'
    },
    @{
        Name = 'Microsoft.Graph.Users'
        Description = 'Microsoft Graph Users module'
        Required = $true
        MinVersion = '2.0.0'
    },
    @{
        Name = 'Microsoft.Graph.Groups'
        Description = 'Microsoft Graph Groups module'
        Required = $true
        MinVersion = '2.0.0'
    },
    @{
        Name = 'Microsoft.Online.SharePoint.PowerShell'
        Description = 'SharePoint Online Management Shell'
        Required = $true
        MinVersion = '16.0.0'
    },
    @{
        Name = 'ExchangeOnlineManagement'
        Description = 'Exchange Online PowerShell V2'
        Required = $true
        MinVersion = '3.0.0'
    },
    @{
        Name = 'MicrosoftTeams'
        Description = 'Microsoft Teams PowerShell module'
        Required = $true
        MinVersion = '5.0.0'
    },
    @{
        Name = 'AzureAD'
        Description = 'Azure Active Directory PowerShell (legacy, still used by some scripts)'
        Required = $false
        MinVersion = '2.0.0'
    },
    @{
        Name = 'MSOnline'
        Description = 'Microsoft Online Services (legacy, still used by some scripts)'
        Required = $false
        MinVersion = '1.1.0'
    },
    @{
        Name = 'Fly.Client'
        Description = 'AvePoint FLY PowerShell module for migration automation'
        Required = $true
        MinVersion = $null
    },
    @{
        Name = 'ImportExcel'
        Description = 'Excel file import/export without Excel installed'
        Required = $true
        MinVersion = '7.0.0'
    },
    @{
        Name = 'PSWriteColor'
        Description = 'Enhanced console output with colors'
        Required = $false
        MinVersion = $null
    }
)

# Check PowerShell version
Write-Host "Checking PowerShell version..." -ForegroundColor Cyan
$psVersion = $PSVersionTable.PSVersion
Write-Host "  PowerShell Version: $($psVersion.Major).$($psVersion.Minor).$($psVersion.Build)" -ForegroundColor White

if ($psVersion.Major -lt 7) {
    Write-Host "  ⚠️  PowerShell 7+ recommended (you have $($psVersion.Major).$($psVersion.Minor))" -ForegroundColor Yellow
    Write-Host "     Download from: https://aka.ms/powershell" -ForegroundColor Gray
} else {
    Write-Host "  ✅ PowerShell version OK" -ForegroundColor Green
}
Write-Host ""

# Check NuGet provider
Write-Host "Checking NuGet package provider..." -ForegroundColor Cyan
$nuget = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
if (-not $nuget) {
    Write-Host "  ⚠️  NuGet provider not found, installing..." -ForegroundColor Yellow
    try {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope $Scope | Out-Null
        Write-Host "  ✅ NuGet provider installed" -ForegroundColor Green
    } catch {
        Write-Host "  ❌ Failed to install NuGet provider: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "  ✅ NuGet provider installed (v$($nuget.Version))" -ForegroundColor Green
}
Write-Host ""

# Check PSGallery repository
Write-Host "Checking PowerShell Gallery..." -ForegroundColor Cyan
$gallery = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
if (-not $gallery) {
    Write-Host "  ⚠️  PSGallery not found, registering..." -ForegroundColor Yellow
    Register-PSRepository -Default
    Write-Host "  ✅ PSGallery registered" -ForegroundColor Green
} elseif ($gallery.InstallationPolicy -ne 'Trusted') {
    Write-Host "  ⚠️  PSGallery not trusted, setting as trusted..." -ForegroundColor Yellow
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    Write-Host "  ✅ PSGallery set as trusted" -ForegroundColor Green
} else {
    Write-Host "  ✅ PSGallery configured and trusted" -ForegroundColor Green
}
Write-Host ""

# Check each required module
Write-Host "Checking required modules..." -ForegroundColor Cyan
Write-Host ("=" * 80) -ForegroundColor Gray

$results = @()
$missingModules = @()
$outdatedModules = @()

foreach ($module in $requiredModules) {
    $moduleName = $module.Name
    $minVersion = $module.MinVersion
    $required = $module.Required

    Write-Host ""
    Write-Host "[$moduleName]" -ForegroundColor White
    Write-Host "  $($module.Description)" -ForegroundColor Gray

    # Check if module is installed
    $installed = Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1

    if ($installed) {
        $installedVersion = $installed.Version
        Write-Host "  Installed: v$installedVersion" -ForegroundColor White

        # Check version requirement
        if ($minVersion) {
            $requiredVer = [version]$minVersion
            if ($installedVersion -lt $requiredVer) {
                Write-Host "  ⚠️  Version $minVersion or higher required" -ForegroundColor Yellow
                $outdatedModules += $module
                $results += [PSCustomObject]@{
                    Module = $moduleName
                    Status = 'Outdated'
                    Installed = $installedVersion.ToString()
                    Required = $minVersion
                    Action = 'Update needed'
                }
            } else {
                Write-Host "  ✅ OK" -ForegroundColor Green
                $results += [PSCustomObject]@{
                    Module = $moduleName
                    Status = 'OK'
                    Installed = $installedVersion.ToString()
                    Required = $minVersion
                    Action = 'None'
                }
            }
        } else {
            Write-Host "  ✅ OK" -ForegroundColor Green
            $results += [PSCustomObject]@{
                Module = $moduleName
                Status = 'OK'
                Installed = $installedVersion.ToString()
                Required = 'Any'
                Action = 'None'
            }
        }
    } else {
        if ($required) {
            Write-Host "  ❌ Not installed (REQUIRED)" -ForegroundColor Red
        } else {
            Write-Host "  ⚠️  Not installed (optional)" -ForegroundColor Yellow
        }
        $missingModules += $module
        $results += [PSCustomObject]@{
            Module = $moduleName
            Status = 'Missing'
            Installed = 'Not installed'
            Required = if ($minVersion) { $minVersion } else { 'Any' }
            Action = if ($required) { 'Install required' } else { 'Install optional' }
        }
    }
}

Write-Host ""
Write-Host ("=" * 80) -ForegroundColor Gray
Write-Host ""

# Summary
Write-Host "SUMMARY" -ForegroundColor Cyan
Write-Host ("=" * 80) -ForegroundColor Gray
$results | Format-Table -AutoSize

$missingCount = ($missingModules | Where-Object { $_.Required }).Count
$optionalMissingCount = ($missingModules | Where-Object { -not $_.Required }).Count
$outdatedCount = $outdatedModules.Count

Write-Host ""
Write-Host "Missing required modules: $missingCount" -ForegroundColor $(if ($missingCount -gt 0) { 'Red' } else { 'Green' })
Write-Host "Missing optional modules: $optionalMissingCount" -ForegroundColor $(if ($optionalMissingCount -gt 0) { 'Yellow' } else { 'Green' })
Write-Host "Outdated modules: $outdatedCount" -ForegroundColor $(if ($outdatedCount -gt 0) { 'Yellow' } else { 'Green' })
Write-Host ""

# Install missing/outdated modules
if ($missingModules.Count -gt 0 -or $outdatedModules.Count -gt 0) {
    if ($AutoInstall) {
        $install = $true
        Write-Host "Auto-install enabled, proceeding..." -ForegroundColor Yellow
    } else {
        Write-Host "Would you like to install/update the missing modules now?" -ForegroundColor Cyan
        Write-Host "  Scope: $Scope" -ForegroundColor Gray
        $response = Read-Host "Install? (Y/N)"
        $install = $response -eq 'Y' -or $response -eq 'y'
    }

    if ($install) {
        Write-Host ""
        Write-Host "Installing/Updating modules..." -ForegroundColor Cyan
        Write-Host ("=" * 80) -ForegroundColor Gray

        $allModules = $missingModules + $outdatedModules | Sort-Object -Property Name -Unique

        foreach ($module in $allModules) {
            $moduleName = $module.Name
            Write-Host ""
            Write-Host "Processing: $moduleName" -ForegroundColor White

            try {
                # Check if already installed (for updates)
                $existing = Get-Module -ListAvailable -Name $moduleName

                if ($existing) {
                    Write-Host "  Updating to latest version..." -ForegroundColor Yellow
                    Update-Module -Name $moduleName -Scope $Scope -Force -ErrorAction Stop
                    Write-Host "  ✅ Updated successfully" -ForegroundColor Green
                } else {
                    Write-Host "  Installing..." -ForegroundColor Yellow
                    $installParams = @{
                        Name = $moduleName
                        Scope = $Scope
                        Force = $true
                        AllowClobber = $true
                    }

                    if ($module.MinVersion) {
                        $installParams.MinimumVersion = $module.MinVersion
                    }

                    Install-Module @installParams -ErrorAction Stop
                    Write-Host "  ✅ Installed successfully" -ForegroundColor Green
                }
            } catch {
                Write-Host "  ❌ Failed: $($_.Exception.Message)" -ForegroundColor Red

                # Special handling for common issues
                if ($moduleName -eq 'Fly.Client') {
                    Write-Host ""
                    Write-Host "  NOTE: Fly.Client may require:" -ForegroundColor Yellow
                    Write-Host "  1. Access to AvePoint's PowerShell Gallery" -ForegroundColor Gray
                    Write-Host "  2. Manual installation from AvePoint" -ForegroundColor Gray
                    Write-Host "  3. Contact AvePoint support for installation package" -ForegroundColor Gray
                }
            }
        }

        Write-Host ""
        Write-Host ("=" * 80) -ForegroundColor Gray
        Write-Host "Installation complete!" -ForegroundColor Green
        Write-Host ""
        Write-Host "IMPORTANT: Close and reopen PowerShell/Migration Toolkit for changes to take effect." -ForegroundColor Yellow
        Write-Host ""
    } else {
        Write-Host ""
        Write-Host "Installation skipped." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "To install manually later, run:" -ForegroundColor Gray
        foreach ($module in ($missingModules + $outdatedModules | Sort-Object -Property Name -Unique)) {
            Write-Host "  Install-Module -Name $($module.Name) -Scope $Scope" -ForegroundColor Gray
        }
        Write-Host ""
    }
} else {
    Write-Host "✅ All required modules are installed and up to date!" -ForegroundColor Green
    Write-Host ""
}

# Additional checks
Write-Host ""
Write-Host "ADDITIONAL CHECKS" -ForegroundColor Cyan
Write-Host ("=" * 80) -ForegroundColor Gray

# Check Node.js (required for FLY briefs)
Write-Host ""
Write-Host "Node.js (required for FLY Executive Brief):" -ForegroundColor White
try {
    $nodeVersion = & node --version 2>$null
    if ($nodeVersion) {
        Write-Host "  ✅ Installed: $nodeVersion" -ForegroundColor Green
    } else {
        Write-Host "  ❌ Not found on PATH" -ForegroundColor Red
        Write-Host "     Download from: https://nodejs.org/" -ForegroundColor Gray
    }
} catch {
    Write-Host "  ❌ Not installed" -ForegroundColor Red
    Write-Host "     Download from: https://nodejs.org/" -ForegroundColor Gray
}

# Check for Chrome/Edge (required for PDF generation)
Write-Host ""
Write-Host "Chrome/Edge (required for PDF generation):" -ForegroundColor White
$browsers = @(
    @{ Name = 'Microsoft Edge'; Path = "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe" },
    @{ Name = 'Microsoft Edge'; Path = "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe" },
    @{ Name = 'Google Chrome'; Path = "$env:ProgramFiles\Google\Chrome\Application\chrome.exe" }
)

$foundBrowser = $false
foreach ($browser in $browsers) {
    if (Test-Path $browser.Path) {
        Write-Host "  ✅ $($browser.Name) found" -ForegroundColor Green
        $foundBrowser = $true
        break
    }
}

if (-not $foundBrowser) {
    Write-Host "  ❌ No suitable browser found" -ForegroundColor Red
    Write-Host "     Install Microsoft Edge or Google Chrome for PDF generation" -ForegroundColor Gray
}

Write-Host ""
Write-Host ("=" * 80) -ForegroundColor Gray
Write-Host ""
Write-Host "Module check complete!" -ForegroundColor Cyan
Write-Host ""

# Output results to file
$outputFile = Join-Path $PSScriptRoot "module-check-results.txt"
$results | Out-File -FilePath $outputFile -Encoding UTF8
Write-Host "Results saved to: $outputFile" -ForegroundColor Gray
Write-Host ""
