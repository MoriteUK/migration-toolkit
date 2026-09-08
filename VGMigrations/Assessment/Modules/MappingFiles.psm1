#Requires -Version 7.0

Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -DisableNameChecking -Force -Global
Import-Module ImportExcel -DisableNameChecking -ErrorAction Stop

# -----------------------------------------------------------------------
# Private functions
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Shows a file open dialog for selecting an assessment workbook, returning the path or null if cancelled.
#>
function Select-AssessmentWorkbook {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog        = [System.Windows.Forms.OpenFileDialog]::new()
    $dialog.Filter = 'Excel files (*.xlsx)|*.xlsx'
    $dialog.Title  = 'Select assessment workbook'
    if ($dialog.ShowDialog() -eq 'OK') { return $dialog.FileName }
    return $null
}

<#
.SYNOPSIS
    Rewrites an email address onto the destination domain by keeping the local part and replacing everything after '@'.
#>
function Convert-EmailToDestination {
    param(
        [AllowNull()][AllowEmptyString()][string]$Address,
        [string]$DestinationDomain
    )
    if ([string]::IsNullOrWhiteSpace($Address)) { return '' }
    return $Address.Split('@')[0] + '@' + $DestinationDomain
}

<#
.SYNOPSIS
    Rewrites a SharePoint URL onto the destination tenant by prepending the destination base URL to the source URI path.
#>
function Convert-UrlToDestination {
    param(
        [AllowNull()][AllowEmptyString()][string]$Url,
        [string]$DestinationBaseUrl
    )
    if ([string]::IsNullOrWhiteSpace($Url)) { return '' }
    $path = ([Uri]$Url).AbsolutePath
    return $DestinationBaseUrl.TrimEnd('/') + $path
}

<#
.SYNOPSIS
    Reads a worksheet from the assessment workbook via Import-Excel, returning an empty array with a warning on failure.
#>
function Import-WorkbookSheet {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$SheetName
    )
    try {
        return @(Import-Excel -Path $Path -WorksheetName $SheetName -ErrorAction Stop)
    }
    catch {
        Write-Host ($PREFIX_WARN + "Could not read sheet '$SheetName': " + $_.Exception.Message) -ForegroundColor Yellow
        return @()
    }
}

<#
.SYNOPSIS
    Tests whether a workbook row is flagged for migration via its Migrate column.
.DESCRIPTION
    Returns true when Migrate = 'Yes' (case-insensitive). Sheets from older workbooks
    that predate the Migrate column have no such property - those rows are treated as
    Migrate = Yes for backward compatibility.
#>
function Test-MigrateFlag {
    param([Parameter(Mandatory)][object]$Row)
    if (-not $Row.PSObject.Properties['Migrate']) { return $true }
    return ($Row.Migrate -eq 'Yes')
}

<#
.SYNOPSIS
    Writes mapping rows to a plain xlsx with a 'Migration mappings' sheet and prints the record count; skips the file when there are no rows.
#>
function Write-MappingFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [AllowNull()][object[]]$Rows,
        [Parameter(Mandatory)][string]$Label
    )

    $count = if ($null -ne $Rows) { @($Rows).Count } else { 0 }
    if ($count -gt 0) {
        # Plain output - no formatting flags, header row then data rows only
        @($Rows) | Export-Excel -Path $Path -WorksheetName 'Migration mappings'
    }
    else {
        Write-Host ($PREFIX_WARN + "$Label - no source rows, file not written") -ForegroundColor Yellow
    }
    Write-ProgressLine -Label $Label -Count $count
}

# -----------------------------------------------------------------------
# Public functions
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Generates the six AvePoint Fly import files from an existing assessment workbook.
.DESCRIPTION
    Prompts for the workbook (file dialog), destination domain, and destination SPO base
    URL, then reads the source sheets with Import-Excel and writes the teams, m365groups,
    sharepoint, onedrive, teamschat, and exchange mapping files to a timestamped folder
    adjacent to the workbook. Reads only the workbook - no live services are queried.
#>
function Invoke-MappingFileGeneration {
    [CmdletBinding()]
    param()

    Write-SectionHeader 'Mapping File Generation'

    # --- Inputs ---
    Write-Host ($PREFIX_INFO + 'Select the assessment workbook...') -ForegroundColor DarkGray
    $workbookPath = Select-AssessmentWorkbook
    if (-not $workbookPath) {
        Write-Host ($PREFIX_SKIP + 'No workbook selected - mapping file generation cancelled') -ForegroundColor Yellow
        return
    }
    Write-Host ($PREFIX_OK + "Workbook: $workbookPath") -ForegroundColor Green

    $destDomain = (Read-Host 'Destination domain        (e.g. newtenant.com)').Trim()
    $destSpoUrl = (Read-Host 'Destination SPO base URL  (e.g. https://newtenant.sharepoint.com)').Trim()

    # --- Output folder - adjacent to the workbook ---
    $vbuName   = [IO.Path]::GetFileNameWithoutExtension($workbookPath) -replace '-Assessment$', ''
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmm'
    $outFolder = Join-Path (Split-Path $workbookPath -Parent) 'MappingFiles' "$vbuName-$timestamp"
    New-Item -ItemType Directory -Path $outFolder -Force | Out-Null
    Write-Host ($PREFIX_OK + "Output folder: $outFolder") -ForegroundColor Green

    # --- Mapping file names - sanitized VBU name from the Assessment Summary sheet ---
    $summarySheet = Import-WorkbookSheet -Path $workbookPath -SheetName 'Assessment Summary'
    $vbuNameValue = ($summarySheet | Where-Object { $_.Section -eq 'VBU Name' } | Select-Object -First 1 -ExpandProperty Value)
    $safeName     = ("$vbuNameValue" -replace '[^\w\-]', '')

    $teamsFileName      = "$safeName-teams-mapping.xlsx"
    $m365GroupsFileName = "$safeName-m365groups-mapping.xlsx"
    $sharepointFileName = "$safeName-sharepoint-mapping.xlsx"
    $onedriveFileName   = "$safeName-onedrive-mapping.xlsx"
    $teamschatFileName  = "$safeName-teamschat-mapping.xlsx"
    $exchangeFileName   = "$safeName-exchange-mapping.xlsx"

    # --- Read source sheets ---
    Write-Host ($PREFIX_INFO + 'Reading workbook sheets...') -ForegroundColor DarkGray
    $teams           = Import-WorkbookSheet -Path $workbookPath -SheetName 'Teams'
    $m365Groups      = Import-WorkbookSheet -Path $workbookPath -SheetName 'M365 Groups'
    $spoSites        = Import-WorkbookSheet -Path $workbookPath -SheetName 'SharePoint Sites'
    $oneDrives       = Import-WorkbookSheet -Path $workbookPath -SheetName 'OneDrives'
    $adUsers         = Import-WorkbookSheet -Path $workbookPath -SheetName 'Users'
    $userMailboxes   = Import-WorkbookSheet -Path $workbookPath -SheetName 'User Mailboxes'
    $sharedMailboxes = Import-WorkbookSheet -Path $workbookPath -SheetName 'Shared Mailboxes'

    # Teams sheet has no PrimarySmtpAddress - resolve via M365 Groups on GroupId
    $groupIdToEmail = @{}
    foreach ($g in $m365Groups) {
        if ($g.GroupId -and $g.PrimarySmtpAddress) { $groupIdToEmail[$g.GroupId] = $g.PrimarySmtpAddress }
    }

    # --- Teams ---
    $teamRows = @($teams |
        Where-Object { $_.MigrationObjectType -eq 'Migrated as Team' -and (Test-MigrateFlag -Row $_) } |
        ForEach-Object {
            $srcEmail = if ($_.GroupId -and $groupIdToEmail.ContainsKey($_.GroupId)) { $groupIdToEmail[$_.GroupId] } else { '' }
            [PSCustomObject]@{
                'Source team name'               = $_.DisplayName
                'Source team email address'      = $srcEmail
                'Destination team name'          = $_.DisplayName
                'Destination team email address' = Convert-EmailToDestination -Address $srcEmail -DestinationDomain $destDomain
            }
        })

    # --- M365 Groups ---
    $groupRows = @($m365Groups |
        Where-Object { $_.MigrationObjectType -eq 'Migrated as M365 Group' } |
        ForEach-Object {
            [PSCustomObject]@{
                'Source group name'               = $_.DisplayName
                'Source group email address'      = $_.PrimarySmtpAddress
                'Destination group name'          = $_.DisplayName
                'Destination group email address' = Convert-EmailToDestination -Address $_.PrimarySmtpAddress -DestinationDomain $destDomain
            }
        })

    # --- SharePoint ---
    $spoRows = @($spoSites |
        Where-Object { $_.MigrationObjectType -eq 'Migrate as SharePoint Site' -and (Test-MigrateFlag -Row $_) } |
        ForEach-Object {
            [PSCustomObject]@{
                'Source URL'               = $_.Url
                'Source object level'      = 'Site collection'
                'Destination URL'          = Convert-UrlToDestination -Url $_.Url -DestinationBaseUrl $destSpoUrl
                'Destination object level' = 'Site collection'
                'Method'                   = 'Merge'
            }
        })

    # --- OneDrive ---
    # Key-field guard excludes the '(No data collected)' placeholder row on empty sheets
    $oneDriveRows = @($oneDrives |
        Where-Object { $_.PSObject.Properties['OwnerUPN'] -and $_.OwnerUPN -and (Test-MigrateFlag -Row $_) } |
        ForEach-Object {
            [PSCustomObject]@{
                'Source user'      = $_.OwnerUPN
                'Destination user' = Convert-EmailToDestination -Address $_.OwnerUPN -DestinationDomain $destDomain
            }
        })

    # --- Teams Chat ---
    $teamsChatRows = @($adUsers |
        Where-Object { $_.PSObject.Properties['UserPrincipalName'] -and $_.UserPrincipalName -and (Test-MigrateFlag -Row $_) } |
        ForEach-Object {
            [PSCustomObject]@{
                'Source user'      = $_.UserPrincipalName
                'Destination user' = Convert-EmailToDestination -Address $_.UserPrincipalName -DestinationDomain $destDomain
            }
        })

    # --- Exchange - User Mailboxes first, then Shared Mailboxes ---
    $exchangeRows = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($mb in ($userMailboxes | Where-Object { $_.PSObject.Properties['PrimarySmtpAddress'] -and $_.PrimarySmtpAddress -and (Test-MigrateFlag -Row $_) })) {
        $exchangeRows.Add([PSCustomObject]@{
            'Source'           = $mb.PrimarySmtpAddress
            'Source type'      = 'User mailbox'
            'Destination'      = Convert-EmailToDestination -Address $mb.PrimarySmtpAddress -DestinationDomain $destDomain
            'Destination type' = 'User mailbox'
        })
    }
    foreach ($mb in ($sharedMailboxes | Where-Object { $_.PSObject.Properties['PrimarySmtpAddress'] -and $_.PrimarySmtpAddress -and (Test-MigrateFlag -Row $_) })) {
        $exchangeRows.Add([PSCustomObject]@{
            'Source'           = $mb.PrimarySmtpAddress
            'Source type'      = 'Shared mailbox'
            'Destination'      = Convert-EmailToDestination -Address $mb.PrimarySmtpAddress -DestinationDomain $destDomain
            'Destination type' = 'Shared mailbox'
        })
    }

    # --- Write files and summary ---
    Write-SectionHeader 'Mapping Files'
    Write-MappingFile -Path (Join-Path $outFolder $teamsFileName)      -Rows $teamRows               -Label 'Teams mappings'
    Write-MappingFile -Path (Join-Path $outFolder $m365GroupsFileName) -Rows $groupRows              -Label 'M365 Group mappings'
    Write-MappingFile -Path (Join-Path $outFolder $sharepointFileName) -Rows $spoRows                -Label 'SharePoint mappings'
    Write-MappingFile -Path (Join-Path $outFolder $onedriveFileName)   -Rows $oneDriveRows           -Label 'OneDrive mappings'
    Write-MappingFile -Path (Join-Path $outFolder $teamschatFileName)  -Rows $teamsChatRows          -Label 'Teams Chat mappings'
    Write-MappingFile -Path (Join-Path $outFolder $exchangeFileName)   -Rows $exchangeRows.ToArray() -Label 'Exchange mappings'

    Write-Host ''
    Write-Host ($PREFIX_OK + 'Mapping file generation complete') -ForegroundColor Green
}

# -----------------------------------------------------------------------
# Exports
# -----------------------------------------------------------------------

Export-ModuleMember -Function 'Invoke-MappingFileGeneration'
