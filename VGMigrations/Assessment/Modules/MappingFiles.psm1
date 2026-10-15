#Requires -Version 7.0

Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -DisableNameChecking -Force -Global
Import-Module ImportExcel -DisableNameChecking -ErrorAction Stop

# Fallback Fly Templates folder, used to pre-fill the prompt - override per-run if a
# tenant's templates live elsewhere.
$script:DefaultFlyTemplatesFolder = 'C:\Users\andyw\OneDrive - Volaris Group\GRP Data Security (Volaris Consolidated) - 3. Execution\M365 Migrations\Fly Templates'

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
    Shows a folder browser dialog for picking the folder the FLY output folder is created in.
#>
function Select-FlyOutputFolder {
    param([string]$InitialFolder)
    Add-Type -AssemblyName System.Windows.Forms
    $dialog             = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dialog.Description = 'Select the tenant folder to create the FLY folder in'
    if ($InitialFolder -and (Test-Path $InitialFolder)) { $dialog.SelectedPath = $InitialFolder }
    if ($dialog.ShowDialog() -eq 'OK') { return $dialog.SelectedPath }
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
    Copies an official Fly template into the output folder and writes rows into its 'Migration
    mappings' sheet, matching each row's properties to the template's own header row.
.DESCRIPTION
    Writing into a copy of the real template (rather than generating a fresh generic xlsx) keeps
    the template's column widths, conditional formatting, and any pre-seeded rows intact - notably
    the SharePoint template's default 'Site collection'/'Merge' rows, which this fills in place
    rather than pushing below a duplicate header. Skips (and does not create) the file when there
    are no source rows.
#>
function Write-FlyTemplateFile {
    param(
        [Parameter(Mandatory)][string]$TemplatePath,
        [Parameter(Mandatory)][string]$OutputPath,
        [AllowNull()][object[]]$Rows,
        [Parameter(Mandatory)][string]$Label
    )

    $count = if ($null -ne $Rows) { @($Rows).Count } else { 0 }

    if (-not (Test-Path $TemplatePath)) {
        Write-Host ($PREFIX_WARN + "$Label - template not found: $TemplatePath") -ForegroundColor Yellow
        return
    }
    if ($count -eq 0) {
        Write-Host ($PREFIX_WARN + "$Label - no source rows, file not written") -ForegroundColor Yellow
        Write-ProgressLine -Label $Label -Count 0
        return
    }

    Copy-Item -Path $TemplatePath -Destination $OutputPath -Force

    $pkg = Open-ExcelPackage -Path $OutputPath
    try {
        $ws = $pkg.Workbook.Worksheets['Migration mappings']

        $headers = [System.Collections.Generic.List[string]]::new()
        $col = 1
        $headerCell = $ws.Cells[1, $col]
        while ($headerCell.Value) {
            $headers.Add([string]$headerCell.Value)
            $col++
            $headerCell = $ws.Cells[1, $col]
        }

        $r = 2
        foreach ($row in @($Rows)) {
            for ($c = 0; $c -lt $headers.Count; $c++) {
                # $ws.Cells[$r, $c + 1] (unparenthesized) misparses as ($r, $c) + 1 - an array
                # concat, not the two-arg indexer call - and then silently returns a
                # System.Object[] instead of an ExcelRange. Parenthesize the column expression.
                $cell = $ws.Cells[$r, ($c + 1)]
                $cell.Value = $row.($headers[$c])
            }
            $r++
        }

        Close-ExcelPackage $pkg
    }
    catch {
        Close-ExcelPackage $pkg -NoSave
        throw
    }

    Write-ProgressLine -Label $Label -Count $count
}

# -----------------------------------------------------------------------
# Public functions
# -----------------------------------------------------------------------

<#
.SYNOPSIS
    Generates the seven AvePoint Fly import files from an existing assessment workbook, filling
    copies of the official Fly Templates.
.DESCRIPTION
    Every parameter is optional and, when omitted, is collected interactively (file/folder
    dialog or Read-Host) exactly as before - this is what Run-Assessment.ps1's console "Generate
    Mapping Files" mode still uses. Passing all five lets a non-interactive caller (the
    MigrationToolkit-Web Electron app's own file/folder pickers, via Generate-FlyMappingFiles.ps1)
    drive this with no prompts at all.

    Reads the source sheets with Import-Excel, keeps only rows flagged Migrate = Yes, rewrites
    each Source identity onto the destination domain/URL, and writes the exchange, user,
    m365groups, teams, teamschat, onedrive, and sharepoint mapping files into copies of their
    official templates under <OutputFolder>\FLY\. Reads only the workbook - no live services are
    queried.
.PARAMETER WorkbookPath
    Path to the assessment workbook. Prompted via file dialog when omitted.
.PARAMETER DestDomain
    Target tenant's destination email domain, e.g. newtenant.com. Prompted via Read-Host when omitted.
.PARAMETER DestSpoUrl
    Destination SPO base URL, e.g. https://newtenant.sharepoint.com. Prompted via Read-Host when omitted.
.PARAMETER FlyTemplatesFolder
    Folder containing the 7 official Fly Templates. Prompted via Read-Host (defaulting to the
    shared OneDrive location) when omitted.
.PARAMETER OutputFolder
    Folder to create the FLY\ output folder in. Prompted via folder dialog when omitted.
#>
function Invoke-MappingFileGeneration {
    [CmdletBinding()]
    param(
        [string]$WorkbookPath,
        [string]$DestDomain,
        [string]$DestSpoUrl,
        [string]$FlyTemplatesFolder,
        [string]$OutputFolder
    )

    Write-SectionHeader 'Mapping File Generation'

    # --- Inputs ---
    $workbookPath = $WorkbookPath
    if (-not $workbookPath) {
        Write-Host ($PREFIX_INFO + 'Select the assessment workbook...') -ForegroundColor DarkGray
        $workbookPath = Select-AssessmentWorkbook
        if (-not $workbookPath) {
            Write-Host ($PREFIX_SKIP + 'No workbook selected - mapping file generation cancelled') -ForegroundColor Yellow
            return
        }
    } elseif (-not (Test-Path $workbookPath)) {
        Write-Host ($PREFIX_FAIL + "Workbook not found: $workbookPath") -ForegroundColor Red
        return
    }
    Write-Host ($PREFIX_OK + "Workbook: $workbookPath") -ForegroundColor Green

    $destDomain = if ($DestDomain) { $DestDomain.Trim() } else { (Read-Host 'Target tenant domain (destination)   (e.g. newtenant.com)').Trim() }
    $destSpoUrl = if ($DestSpoUrl) { $DestSpoUrl.Trim() } else { (Read-Host 'Destination SPO base URL              (e.g. https://newtenant.sharepoint.com)').Trim() }

    if ($FlyTemplatesFolder) {
        $flyTemplatesFolder = $FlyTemplatesFolder.Trim()
    } else {
        $templatesInput     = (Read-Host "Fly Templates folder [$script:DefaultFlyTemplatesFolder]").Trim()
        $flyTemplatesFolder = if ($templatesInput) { $templatesInput } else { $script:DefaultFlyTemplatesFolder }
    }
    if (-not (Test-Path $flyTemplatesFolder)) {
        Write-Host ($PREFIX_FAIL + "Fly Templates folder not found: $flyTemplatesFolder") -ForegroundColor Red
        return
    }

    $selectedFolder = $OutputFolder
    if (-not $selectedFolder) {
        Write-Host ($PREFIX_INFO + 'Select the folder to create the FLY folder in...') -ForegroundColor DarkGray
        $selectedFolder = Select-FlyOutputFolder -InitialFolder (Split-Path $workbookPath -Parent)
        if (-not $selectedFolder) {
            Write-Host ($PREFIX_SKIP + 'No folder selected - mapping file generation cancelled') -ForegroundColor Yellow
            return
        }
    } elseif (-not (Test-Path $selectedFolder)) {
        Write-Host ($PREFIX_FAIL + "Output folder not found: $selectedFolder") -ForegroundColor Red
        return
    }

    # --- Output folder ---
    $outFolder = Join-Path $selectedFolder 'FLY'
    New-Item -ItemType Directory -Path $outFolder -Force | Out-Null
    Write-Host ($PREFIX_OK + "Output folder: $outFolder") -ForegroundColor Green

    # --- Mapping file names - sanitized VBU name from the Assessment Summary sheet ---
    $summarySheet = Import-WorkbookSheet -Path $workbookPath -SheetName 'Assessment Summary'
    $vbuNameValue = ($summarySheet | Where-Object { $_.Section -eq 'VBU Name' } | Select-Object -First 1 -ExpandProperty Value)
    $safeName     = ("$vbuNameValue" -replace '[^\w\-]', '')

    # Official template filenames, as shipped by AvePoint - must match exactly
    $templateFiles = @{
        Exchange   = 'Fly_Exchange_Online_Import_Mapping_Template.xlsx'
        User       = 'Fly_Import_User_Mapping_Template.xlsx'
        M365Groups = 'Fly_Microsoft_365_Groups_Import_Mapping_Template.xlsx'
        Teams      = 'Fly_Microsoft_Teams_Add_Mapping.xlsx'
        TeamsChat  = 'Fly_Microsoft_Teams_Chat_Import_Mapping_Template.xlsx'
        OneDrive   = 'Fly_OneDrive_Import_Mapping_Template.xlsx'
        SharePoint = 'Fly_SharePoint_Online_Import_Mapping_Template.xlsx'
    }

    $outputFiles = @{
        Exchange   = "$safeName-exchange-mapping.xlsx"
        User       = "$safeName-user-mapping.xlsx"
        M365Groups = "$safeName-m365groups-mapping.xlsx"
        Teams      = "$safeName-teams-mapping.xlsx"
        TeamsChat  = "$safeName-teamschat-mapping.xlsx"
        OneDrive   = "$safeName-onedrive-mapping.xlsx"
        SharePoint = "$safeName-sharepoint-mapping.xlsx"
    }

    # --- Read source sheets ---
    Write-Host ($PREFIX_INFO + 'Reading workbook sheets...') -ForegroundColor DarkGray
    $teams           = Import-WorkbookSheet -Path $workbookPath -SheetName 'Teams'
    $m365Groups      = Import-WorkbookSheet -Path $workbookPath -SheetName 'M365 Groups'
    $spoSites        = Import-WorkbookSheet -Path $workbookPath -SheetName 'SharePoint Sites'
    $oneDrives       = Import-WorkbookSheet -Path $workbookPath -SheetName 'OneDrives'
    $adUsers         = Import-WorkbookSheet -Path $workbookPath -SheetName 'AD Users'
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

    # --- User (identity mapping) - same in-scope AD user set as Teams Chat ---
    $userRows = @($adUsers |
        Where-Object { $_.PSObject.Properties['UserPrincipalName'] -and $_.UserPrincipalName -and (Test-MigrateFlag -Row $_) } |
        ForEach-Object {
            [PSCustomObject]@{
                'Source user/group'      = $_.UserPrincipalName
                'Destination user/group' = Convert-EmailToDestination -Address $_.UserPrincipalName -DestinationDomain $destDomain
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
    Write-FlyTemplateFile -TemplatePath (Join-Path $flyTemplatesFolder $templateFiles.Exchange)   -OutputPath (Join-Path $outFolder $outputFiles.Exchange)   -Rows $exchangeRows.ToArray() -Label 'Exchange mappings'
    Write-FlyTemplateFile -TemplatePath (Join-Path $flyTemplatesFolder $templateFiles.User)       -OutputPath (Join-Path $outFolder $outputFiles.User)       -Rows $userRows               -Label 'User mappings'
    Write-FlyTemplateFile -TemplatePath (Join-Path $flyTemplatesFolder $templateFiles.M365Groups) -OutputPath (Join-Path $outFolder $outputFiles.M365Groups) -Rows $groupRows              -Label 'M365 Group mappings'
    Write-FlyTemplateFile -TemplatePath (Join-Path $flyTemplatesFolder $templateFiles.Teams)      -OutputPath (Join-Path $outFolder $outputFiles.Teams)      -Rows $teamRows               -Label 'Teams mappings'
    Write-FlyTemplateFile -TemplatePath (Join-Path $flyTemplatesFolder $templateFiles.TeamsChat)  -OutputPath (Join-Path $outFolder $outputFiles.TeamsChat)  -Rows $teamsChatRows          -Label 'Teams Chat mappings'
    Write-FlyTemplateFile -TemplatePath (Join-Path $flyTemplatesFolder $templateFiles.OneDrive)   -OutputPath (Join-Path $outFolder $outputFiles.OneDrive)   -Rows $oneDriveRows           -Label 'OneDrive mappings'
    Write-FlyTemplateFile -TemplatePath (Join-Path $flyTemplatesFolder $templateFiles.SharePoint) -OutputPath (Join-Path $outFolder $outputFiles.SharePoint) -Rows $spoRows                -Label 'SharePoint mappings'

    Write-Host ''
    Write-Host ($PREFIX_OK + 'Mapping file generation complete') -ForegroundColor Green
}

# -----------------------------------------------------------------------
# Exports
# -----------------------------------------------------------------------

Export-ModuleMember -Function 'Invoke-MappingFileGeneration'
