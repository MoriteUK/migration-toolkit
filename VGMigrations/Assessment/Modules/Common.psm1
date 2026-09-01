#Requires -Version 7.0

#region Prefix constants
$script:PREFIX_OK      = '[OK]   '
$script:PREFIX_WARN    = '[!!]   '
$script:PREFIX_FAIL    = '[FAIL] '
$script:PREFIX_SKIP    = '[--]   '
$script:PREFIX_INFO    = '[>>]   '
$script:PREFIX_PENDING = '[ ]    '
#endregion

#region Module-scope state
$script:CollectorStatus = [ordered]@{}
#endregion

# -----------------------------------------------------------------------
# Private functions
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Derives the VBU folder name from a domain by stripping single or compound ccTLD trailing segments.
#>
function Get-VBUName {
    param([string]$Domain)
    $parts = $Domain.Split('.')
    if ($parts.Count -lt 2) { return $Domain }
    $second = $parts[-2]
    # Compound ccTLD: second-to-last segment is 3 chars or fewer and there
    # are at least 3 parts (e.g. mydomain.co.uk). Strip two trailing segments.
    if ($second.Length -le 3 -and $parts.Count -gt 2) {
        return ($parts[0..($parts.Count - 3)] -join '.')
    }
    # Single TLD: strip last segment only.
    return ($parts[0..($parts.Count - 2)] -join '.')
}

<#
.SYNOPSIS
    Writes the module-scoped collector status table to CollectorStatus.json in the Raw output folder.
#>
function Save-CollectorStatus {
    param([Parameter(Mandatory)][string]$RawPath)
    $path = Join-Path $RawPath 'CollectorStatus.json'
    $script:CollectorStatus | ConvertTo-Json -Depth 5 | Set-Content -Path $path -Encoding UTF8
}

# -----------------------------------------------------------------------
# Public functions
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Creates the assessment context object passed to all collectors.
.DESCRIPTION
    Builds a PSCustomObject holding the VBU identifiers, Raw output path, SPO admin URL,
    assessment start timestamp, and the SkipSharePoint/SkipAD/SkipPowerPlatform flags.
    VBUName is derived from VBUDomain via TLD stripping and is used for folder naming only.
#>
function New-AssessmentContext {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$VBUDomain,
        [Parameter(Mandatory)][string]$VBUId,
        [Parameter(Mandatory)][string]$VBUSearchTerm,
        [Parameter(Mandatory)][string]$RawPath,
        [string]$SPOAdminUrl = ''
    )

    [PSCustomObject]@{
        VBUDomain         = $VBUDomain
        VBUName           = Get-VBUName -Domain $VBUDomain
        VBUId             = $VBUId
        VBUSearchTerm     = $VBUSearchTerm
        RawPath           = $RawPath
        SPOAdminUrl       = $SPOAdminUrl
        AssessmentDate    = Get-Date
        SkipSharePoint    = $false
        SkipAD            = $false
        SkipPowerPlatform = $false
    }
}

# Callers pass $ctx.RawPath directly - not the full context object. Collectors follow this same pattern.
<#
.SYNOPSIS
    Serializes collector data to a JSON file in the Raw output folder.
.DESCRIPTION
    Writes the object as UTF-8 JSON (depth 10, enums as strings) to RawPath\FileName.
    Null data is serialized as an empty array so downstream readers always get valid JSON.
#>
function Write-JsonOutput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$FileName,
        [Parameter(Mandatory)][AllowNull()]$Data,
        [Parameter(Mandatory)][string]$RawPath
    )

    $path        = Join-Path $RawPath $FileName
    $toSerialize = $null -ne $Data ? $Data : @()
    ConvertTo-Json -InputObject $toSerialize -Depth 10 -EnumsAsStrings |
        Set-Content -Path $path -Encoding UTF8
    Write-Host ($script:PREFIX_OK + $FileName + ' written') -ForegroundColor Green
}

<#
.SYNOPSIS
    Loads a JSON file from a Raw output folder, returning an empty array if the file is missing, empty, or null.
.DESCRIPTION
    Shared by Workbook.psm1 (assessment workbook) and LegacyExport.psm1 (Domain Removal
    CSV compatibility layer) so both read collector output the same way.
#>
function Import-AssessmentJson {
    param(
        [Parameter(Mandatory)][string]$FileName,
        [Parameter(Mandatory)][string]$RawPath
    )
    # The leading comma on every return below is deliberate, not decorative: PowerShell
    # unrolls an array written to the pipeline into its individual elements, so a bare
    # "return @()" (zero elements) hands the caller literally nothing - and `$x = Import-...`
    # then assigns $null, not an empty array. `,@()` wraps the array as a single pipeline
    # object, which is what actually survives the return boundary as an empty array. Confirmed
    # live: without this, every direct (unwrapped) caller of this function crashed on
    # Export-Csv with "Cannot bind argument to parameter 'InputObject' because it is null"
    # whenever the underlying collector genuinely found zero of something.
    $path = Join-Path $RawPath $FileName
    if (-not (Test-Path $path)) { return ,@() }
    $content = Get-Content $path -Raw -ErrorAction SilentlyContinue
    if ([string]::IsNullOrWhiteSpace($content)) { return ,@() }
    $parsed = $content | ConvertFrom-Json
    if ($null -eq $parsed) { return ,@() }
    return ,@($parsed)
}

<#
.SYNOPSIS
    Builds the standard result object returned by every collector.
.DESCRIPTION
    Returns a PSCustomObject with Success, ErrorMessage, Counts, and Skipped fields.
    The parameter is named ErrorMessage rather than Error because $Error is a reserved
    PowerShell automatic variable.
#>
function New-CollectorResult {
    [CmdletBinding()]
    param(
        [bool]$Success        = $true,
        [string]$ErrorMessage = '',
        [hashtable]$Counts    = @{},
        [bool]$Skipped        = $false
    )

    [PSCustomObject]@{
        Success      = $Success
        ErrorMessage = $ErrorMessage
        Counts       = $Counts
        Skipped      = $Skipped
    }
}

<#
.SYNOPSIS
    Derives the persistent cache folder for a given source tenant, keyed by its SPO admin URL.
.DESCRIPTION
    Shared by SharePoint.psm1 and TeamMemberships.psm1's cache read/write functions, and by
    Run-Assessment.ps1's pre-flight freshness check, so all three agree on the same folder for
    the same tenant. Every VBU split from the same source tenant reuses this one cache -
    SharePointAdminUrl is already the one tenant-identifying input Run-Assessment.ps1 collects,
    so it doubles as the cache key rather than inventing a second one.
#>
function Get-DiscoveryCacheFolder {
    param(
        [Parameter(Mandatory)][string]$SharePointAdminUrl,
        [Parameter(Mandatory)][string]$CacheRoot
    )
    $safe = ($SharePointAdminUrl -replace '^https?://', '') -replace '[^a-zA-Z0-9\.\-]', '-'
    if (-not $safe) { $safe = 'default' }
    return Join-Path $CacheRoot $safe
}

<#
.SYNOPSIS
    Checks whether the SharePoint Sites and Teams/Channels/Members caches exist and are fresh.
.DESCRIPTION
    Run-Assessment.ps1 calls this before doing any tenant work and stops the whole run if either
    cache is missing or older than MaxAgeDays - both files are produced by the standalone
    Update-SharePointSitesCache.ps1 / Update-TeamsChannelsCache.ps1 scripts, which walk the
    entire tenant (the slow part) once so every VBU's Discovery run can just read and
    locally filter the result instead of repeating that walk every time.
#>
function Test-DiscoveryCachePrereqs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SharePointAdminUrl,
        [Parameter(Mandatory)][string]$CacheRoot,
        [int]$MaxAgeDays = 7
    )

    $folder = Get-DiscoveryCacheFolder -SharePointAdminUrl $SharePointAdminUrl -CacheRoot $CacheRoot
    $checks = @(
        @{ Name = 'SharePoint Sites (+ OneDrive)'; FileName = 'SharePointSites.json';        Script = 'Update-SharePointSitesCache.ps1' }
        @{ Name = 'Teams Channels & Members';      FileName = 'TeamsChannelsMembers.json';   Script = 'Update-TeamsChannelsCache.ps1' }
    )

    $problems = [System.Collections.Generic.List[string]]::new()
    foreach ($c in $checks) {
        $path = Join-Path $folder $c.FileName
        if (-not (Test-Path $path)) {
            $problems.Add("$($c.Name) cache not found. Run: VGMigrations\$($c.Script) -SharePointAdminUrl `"$SharePointAdminUrl`"")
            continue
        }
        $ageDays = ((Get-Date) - (Get-Item $path).LastWriteTime).TotalDays
        if ($ageDays -gt $MaxAgeDays) {
            $problems.Add("$($c.Name) cache is $([math]::Floor($ageDays)) day(s) old (max $MaxAgeDays). Run: VGMigrations\$($c.Script) -SharePointAdminUrl `"$SharePointAdminUrl`"")
        }
    }

    [PSCustomObject]@{
        IsFresh     = ($problems.Count -eq 0)
        CacheFolder = $folder
        Problems    = $problems.ToArray()
    }
}

<#
.SYNOPSIS
    Records a collector's status and persists CollectorStatus.json.
.DESCRIPTION
    Updates the module-scoped status table with the status, optional message, elapsed
    duration from StartTime, and an update timestamp, then writes the file to RawPath
    via Save-CollectorStatus.
#>
function Update-CollectorStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$CollectorName,
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][string]$RawPath,
        [Parameter(Mandatory)][datetime]$StartTime,
        [string]$Message = ''
    )

    $duration = Format-Duration -Start $StartTime
    $script:CollectorStatus[$CollectorName] = [ordered]@{
        Status   = $Status
        Message  = $Message
        Duration = $duration
        Updated  = (Get-Date -Format 'o')
    }
    Save-CollectorStatus -RawPath $RawPath
}

<#
.SYNOPSIS
    Converts EXO/SPO size values to gigabytes.
.DESCRIPTION
    Handles null, empty, and 'Unlimited' (all return 0), ByteQuantifiedSize strings like
    "1.5 GB (1,610,612,736 bytes)", unit-tagged strings, and raw numeric values - bytes
    by default, megabytes when -FromMB is set. Returns a double rounded to 2 decimal places.
#>
function Convert-SizeToGB {
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyString()]$Value,
        [switch]$FromMB
    )

    if ($null -eq $Value) { return [double]0 }
    $str = "$Value".Trim()
    if ($str -eq '' -or $str -eq 'Unlimited') { return [double]0 }

    # ByteQuantifiedSize string: "1.5 GB (1,610,612,736 bytes)"
    if ($str -match '\(([0-9,]+)\s+bytes\)') {
        $bytes = [long]($Matches[1] -replace ',', '')
        return [math]::Round($bytes / 1GB, 2)
    }

    # Unit-tagged string: "1.5 GB", "500 MB", "100 KB"
    if ($str -match '^([\d.]+)\s*(GB|MB|KB|B|bytes)$') {
        $num = [double]$Matches[1]
        switch ($Matches[2]) {
            'GB'    { return [math]::Round($num, 2) }
            'MB'    { return [math]::Round($num / 1024, 2) }
            'KB'    { return [math]::Round($num / 1MB, 2) }
            default { return [math]::Round($num / 1GB, 2) }
        }
    }

    # Raw numeric value: bytes by default, MB if -FromMB
    if ($str -match '^[\d.]+$') {
        $num = [double]$str
        if ($FromMB) { return [math]::Round($num / 1024, 2) }
        return [math]::Round($num / 1GB, 2)
    }

    return [double]0
}

<#
.SYNOPSIS
    Formats the elapsed time since a start datetime as an mm:ss string.
.DESCRIPTION
    Computes whole seconds elapsed between Start and now. Returns a zero-padded
    minutes:seconds string; minutes are not capped at 60.
#>
function Format-Duration {
    [CmdletBinding()]
    param([Parameter(Mandatory)][datetime]$Start)

    $total = [int]((Get-Date) - $Start).TotalSeconds
    '{0:d2}:{1:d2}' -f ([int]($total / 60)), ($total % 60)
}

<#
.SYNOPSIS
    Writes a consistent label/count progress line to the console.
.DESCRIPTION
    Prints the label in dark gray with the info prefix, followed by the count in green
    on the same line. Used by collectors to report per-workload result counts.
#>
function Write-ProgressLine {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Label,
        [Parameter(Mandatory)][int]$Count
    )

    Write-Host ($script:PREFIX_INFO + $Label + ': ') -ForegroundColor DarkGray -NoNewline
    Write-Host $Count -ForegroundColor Green
}

<#
.SYNOPSIS
    Writes a color-coded section header to the console.
.DESCRIPTION
    Prints a blank line followed by the title in cyan. Used to delimit each
    phase of the run in console output.
#>
function Write-SectionHeader {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Title)

    Write-Host ''
    Write-Host $Title -ForegroundColor Cyan
}

# -----------------------------------------------------------------------
# Exports
# -----------------------------------------------------------------------

Export-ModuleMember -Function @(
    'New-AssessmentContext'
    'Write-JsonOutput'
    'Import-AssessmentJson'
    'New-CollectorResult'
    'Update-CollectorStatus'
    'Convert-SizeToGB'
    'Format-Duration'
    'Write-ProgressLine'
    'Write-SectionHeader'
    'Get-DiscoveryCacheFolder'
    'Test-DiscoveryCachePrereqs'
) -Variable @(
    'PREFIX_OK'
    'PREFIX_WARN'
    'PREFIX_FAIL'
    'PREFIX_SKIP'
    'PREFIX_INFO'
    'PREFIX_PENDING'
)
