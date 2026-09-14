#Requires -Version 7.0
<#
.SYNOPSIS
    Generate-FlyMappingFiles.ps1 — non-interactive wrapper around Assessment\Modules\
    MappingFiles.psm1's Invoke-MappingFileGeneration, for the MigrationToolkit-Web Electron app's
    AvePoint Steps > Mapping Files screen (its own file/folder dialogs supply every value up
    front, so this never prompts).

.DESCRIPTION
    Fills copies of the 7 official AvePoint Fly Templates (Exchange, User, M365 Groups, Teams,
    Teams Chat, OneDrive, SharePoint) from an assessment workbook's Migrate = Yes rows, and writes
    them to <OutputFolder>\FLY\. See Invoke-MappingFileGeneration for the full behaviour.

.EXAMPLE
    .\Generate-FlyMappingFiles.ps1 -WorkbookPath "C:\Discovery\Contoso-Assessment.xlsx" `
        -DestDomain "newtenant.com" -DestSpoUrl "https://newtenant.sharepoint.com" `
        -OutputFolder "C:\Discovery\Contoso"
#>

param(
    [Parameter(Mandatory)][string]$WorkbookPath,
    [Parameter(Mandatory)][string]$DestDomain,
    [Parameter(Mandatory)][string]$DestSpoUrl,
    # Defaulted (not Mandatory) so a non-interactive caller can never hit
    # Invoke-MappingFileGeneration's Read-Host fallback, which throws
    # ("PowerShell is in NonInteractive mode") when this process has no console.
    [string]$FlyTemplatesFolder = 'C:\Users\andyw\OneDrive - Volaris Group\GRP Data Security (Volaris Consolidated) - 3. Execution\M365 Migrations\Fly Templates',
    [Parameter(Mandatory)][string]$OutputFolder
)

$moduleRoot = Join-Path $PSScriptRoot 'Assessment\Modules'
Import-Module (Join-Path $moduleRoot 'Common.psm1')       -Force -DisableNameChecking -Global
Import-Module ImportExcel                                 -DisableNameChecking -ErrorAction Stop
Import-Module (Join-Path $moduleRoot 'MappingFiles.psm1')  -Force -DisableNameChecking -Global

$params = @{
    WorkbookPath = $WorkbookPath
    DestDomain   = $DestDomain
    DestSpoUrl   = $DestSpoUrl
    OutputFolder = $OutputFolder
}
if ($FlyTemplatesFolder) { $params.FlyTemplatesFolder = $FlyTemplatesFolder }

Invoke-MappingFileGeneration @params
