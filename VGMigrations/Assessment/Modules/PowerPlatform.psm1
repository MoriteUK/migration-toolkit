#Requires -Version 7.0

Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -DisableNameChecking -Force -Global

# -----------------------------------------------------------------------
# Private functions
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Writes the Power Platform child scan script to disk and returns its path.
.DESCRIPTION
    Power Platform's admin SDK conflicts with the Microsoft.Identity.Client assembly
    version already loaded by the Graph SDK in this process, so the scan (including its
    own interactive sign-in) always runs in a separate pwsh.exe child process - same
    reason search-domain.ps1 did this. The child writes JSON (not CSV) so the parent can
    read it straight back into Write-JsonOutput.
#>
function New-PowerPlatformChildScript {
    param([Parameter(Mandatory)][string]$Path)

    @'
param(
    [Parameter(Mandatory)][string]$Domain,
    [Parameter(Mandatory)][string]$DomainPrefix,
    [Parameter(Mandatory)][string]$JsonPath,
    [Parameter(Mandatory)][string]$LogPath
)

function Write-ChildLog {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry = "[$ts] [$($Level.PadRight(7))] [pp-child] $Message"
    Add-Content -Path $LogPath -Value $entry -Encoding UTF8
    Write-Host $entry
}

try {
    Import-Module 'Microsoft.PowerApps.Administration.PowerShell' -ErrorAction Stop
    Add-PowerAppsAccount -ErrorAction Stop
    Write-ChildLog "Authenticated to Power Platform." -Level SUCCESS

    $results = [System.Collections.Generic.List[object]]::new()
    Write-ChildLog "Scanning Power Apps..."
    try {
        $apps = @(Get-AdminPowerApp -ErrorAction Stop | Where-Object {
            ($_.Owner.UserPrincipalName -like "*$Domain*") -or
            ($_.Internal.displayName    -like "*$DomainPrefix*")
        })
        foreach ($a in $apps) {
            $results.Add([PSCustomObject]@{
                ObjectType   = 'PowerApp'
                DisplayName  = $a.Internal.displayName
                ObjectId     = $a.AppName
                Owner        = $a.Owner.UserPrincipalName
                Environment  = $a.EnvironmentName
                CreatedTime  = $a.Internal.createdTime
                LastModified = $a.Internal.lastModifiedTime
                Action       = 'Export and re-import; reassign owner in new tenant'
            }) | Out-Null
        }
        Write-ChildLog "Found $($apps.Count) Power App(s) matching domain."
    } catch { Write-ChildLog "Power Apps scan failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" -Level WARN }

    Write-ChildLog "Scanning Power Automate flows..."
    try {
        $envs = @(Get-AdminPowerAppEnvironment -ErrorAction Stop)
        foreach ($env in $envs) {
            try {
                $flows = @(Get-AdminFlow -EnvironmentName $env.EnvironmentName -ErrorAction Stop | Where-Object {
                    ($_.Internal.properties.creator.userPrincipalName -like "*$Domain*") -or
                    ($_.Internal.properties.displayName               -like "*$DomainPrefix*")
                })
                foreach ($f in $flows) {
                    $results.Add([PSCustomObject]@{
                        ObjectType   = 'PowerAutomateFlow'
                        DisplayName  = $f.Internal.properties.displayName
                        ObjectId     = $f.FlowName
                        Owner        = $f.Internal.properties.creator.userPrincipalName
                        Environment  = $env.DisplayName
                        CreatedTime  = $f.Internal.properties.creationTime
                        LastModified = $f.Internal.properties.lastModifiedTime
                        Action       = 'Export and re-import; re-authenticate all connections'
                    }) | Out-Null
                }
            } catch {
                Write-ChildLog "Flows in environment '$($env.DisplayName)' failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" -Level WARN
            }
        }
    } catch { Write-ChildLog "Power Automate scan failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" -Level WARN }

    $results | ConvertTo-Json -Depth 6 -AsArray | Set-Content -Path $JsonPath -Encoding UTF8
    Write-ChildLog "Wrote $($results.Count) record(s) -> $(Split-Path $JsonPath -Leaf)" -Level SUCCESS
}
catch {
    Write-ChildLog "Power Platform scan failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" -Level ERROR
    exit 1
}
'@ | Set-Content -Path $Path -Encoding UTF8
}

# -----------------------------------------------------------------------
# Public functions
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Orchestrates Power Platform (Power Apps + Power Automate) collection via a child process.
.DESCRIPTION
    Returns a skipped result immediately when SkipPowerPlatform is set. Otherwise writes and
    runs a child pwsh.exe script that signs in independently (Add-PowerAppsAccount) and scans
    Power Apps/Flows matching VBUDomain/VBUSearchTerm, then reads its JSON output back and
    writes PowerPlatform.json. Failures (including a missing module or a failed child run) are
    non-critical and are returned as a failed/skipped collector result - never propagated.
#>
function Invoke-PowerPlatformCollection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][PSCustomObject]$Context,
        [switch]$SkipPowerPlatform
    )

    $start = Get-Date
    Write-SectionHeader 'Power Platform'

    if ($SkipPowerPlatform -or $Context.SkipPowerPlatform) {
        Write-Host ($PREFIX_SKIP + 'Power Platform collection skipped') -ForegroundColor Yellow
        Update-CollectorStatus -CollectorName 'Power Platform Data' -Status 'Skipped' `
            -RawPath $Context.RawPath -StartTime $start
        Write-JsonOutput -FileName 'PowerPlatform.json' -Data @() -RawPath $Context.RawPath
        return New-CollectorResult -Skipped $true
    }

    if (-not (Get-Module -ListAvailable -Name 'Microsoft.PowerApps.Administration.PowerShell')) {
        Write-Host ($PREFIX_WARN + 'Microsoft.PowerApps.Administration.PowerShell not available - Power Platform collection skipped') -ForegroundColor Yellow
        Update-CollectorStatus -CollectorName 'Power Platform Data' -Status 'Skipped' `
            -RawPath $Context.RawPath -StartTime $start -Message 'Module not available'
        Write-JsonOutput -FileName 'PowerPlatform.json' -Data @() -RawPath $Context.RawPath
        return New-CollectorResult -Skipped $true
    }

    $childScript = Join-Path $Context.RawPath '_PowerPlatform_scan.ps1'
    $childJson   = Join-Path $Context.RawPath '_PowerPlatform_scan.json'
    $childLog    = Join-Path $Context.RawPath ("_PowerPlatform_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.log')

    try {
        Write-Host ($PREFIX_INFO + 'Launching Power Platform scan (separate sign-in window)...') -ForegroundColor DarkGray
        New-PowerPlatformChildScript -Path $childScript

        $childArgs = @(
            '-NoProfile', '-NoLogo', '-ExecutionPolicy', 'Bypass',
            '-File', "`"$childScript`"",
            '-Domain', $Context.VBUDomain,
            '-DomainPrefix', $Context.VBUSearchTerm,
            '-JsonPath', "`"$childJson`"",
            '-LogPath', "`"$childLog`""
        )
        $proc = Start-Process -FilePath 'pwsh.exe' -ArgumentList $childArgs -Wait -NoNewWindow -PassThru

        if (Test-Path $childLog) {
            Get-Content $childLog | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkGray }
        }

        if ($proc.ExitCode -ne 0 -or -not (Test-Path $childJson)) {
            Write-Host ($PREFIX_WARN + "Power Platform scan failed (exit $($proc.ExitCode))") -ForegroundColor Yellow
            Update-CollectorStatus -CollectorName 'Power Platform Data' -Status 'Failed' `
                -RawPath $Context.RawPath -StartTime $start -Message "Child exit code $($proc.ExitCode)"
            Write-JsonOutput -FileName 'PowerPlatform.json' -Data @() -RawPath $Context.RawPath
            return New-CollectorResult -Success $false -ErrorMessage "Power Platform child process exit code $($proc.ExitCode)"
        }

        $results = @(Get-Content $childJson -Raw | ConvertFrom-Json)
        $powerApps = @($results | Where-Object { $_.ObjectType -eq 'PowerApp' })
        $flows     = @($results | Where-Object { $_.ObjectType -eq 'PowerAutomateFlow' })

        Write-ProgressLine -Label 'Power Apps'            -Count $powerApps.Count
        Write-ProgressLine -Label 'Power Automate Flows'  -Count $flows.Count

        Write-JsonOutput -FileName 'PowerPlatform.json' -Data $results -RawPath $Context.RawPath

        $counts = @{
            PowerAppCount = $powerApps.Count
            FlowCount     = $flows.Count
        }
        Update-CollectorStatus -CollectorName 'Power Platform Data' -Status 'Complete' `
            -RawPath $Context.RawPath -StartTime $start -Message "Power Apps: $($counts.PowerAppCount), Flows: $($counts.FlowCount)"

        return New-CollectorResult -Success $true -Counts $counts
    }
    catch {
        Write-Host ($PREFIX_FAIL + 'Power Platform collection failed: ' + $_.Exception.Message) -ForegroundColor Red
        Update-CollectorStatus -CollectorName 'Power Platform Data' -Status 'Failed' `
            -RawPath $Context.RawPath -StartTime $start -Message $_.Exception.Message
        Write-JsonOutput -FileName 'PowerPlatform.json' -Data @() -RawPath $Context.RawPath
        return New-CollectorResult -Success $false -ErrorMessage $_.Exception.Message
    }
    finally {
        if (Test-Path $childScript) { Remove-Item $childScript -Force -ErrorAction SilentlyContinue }
        if (Test-Path $childJson)   { Remove-Item $childJson   -Force -ErrorAction SilentlyContinue }
    }
}

# -----------------------------------------------------------------------
# Exports
# -----------------------------------------------------------------------

Export-ModuleMember -Function 'Invoke-PowerPlatformCollection'
