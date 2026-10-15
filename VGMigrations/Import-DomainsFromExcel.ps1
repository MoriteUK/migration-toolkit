#Requires -Version 7.0
<#
.SYNOPSIS
    Converts an Excel workbook into domains.json for the M365 Discovery Launcher.

.DESCRIPTION
    Reads a sheet from an xlsx file and maps two columns — one for the domain name
    and one for the VBU ID — into the JSON format expected by discovery-menu.ps1.

    Uses ImportExcel if available (auto-installs from PSGallery); falls back to
    the Excel COM object (requires Excel to be installed).

.PARAMETER XlsxPath
    Path to the xlsx file.  Prompted if omitted.

.PARAMETER DomainColumn
    Column header that contains domain names.  Default: 'Domain'

.PARAMETER VbuIdColumn
    Column header that contains VBU IDs.  Default: 'VBU ID'
    Use '' or omit if there is no VBU ID column — all entries get a blank vbuId.

.PARAMETER SheetName
    Worksheet name to read.  Defaults to the first sheet.

.PARAMETER OutputPath
    Where to write domains.json.  Defaults to the folder this script lives in.

.PARAMETER Append
    When set, merges new rows into an existing domains.json rather than replacing it.

.EXAMPLE
    .\Import-DomainsFromExcel.ps1 -XlsxPath "C:\Data\Clients.xlsx"

.EXAMPLE
    .\Import-DomainsFromExcel.ps1 -XlsxPath "C:\Data\Clients.xlsx" `
        -DomainColumn "Email Domain" -VbuIdColumn "BU Code" -SheetName "Sheet2"
#>
[CmdletBinding()]
param(
    [string]$XlsxPath    = '',
    [string]$DomainColumn = 'Domain',
    [string]$VbuIdColumn  = 'VBU ID',
    [string]$SheetName    = '',
    [string]$OutputPath   = '',
    [switch]$Append
)

$ErrorActionPreference = 'Stop'
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

function Write-Step  { param([string]$M) Write-Host "`n==> $M" -ForegroundColor Cyan   }
function Write-Ok    { param([string]$M) Write-Host "    [OK]   $M" -ForegroundColor Green  }
function Write-Warn  { param([string]$M) Write-Host "    [WARN] $M" -ForegroundColor Yellow }
function Write-Err   { param([string]$M) Write-Host "    [ERR]  $M" -ForegroundColor Red    }

# ── Resolve xlsx path ──────────────────────────────────────────────────────────
if (-not $XlsxPath) {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    $ofd = [System.Windows.Forms.OpenFileDialog]::new()
    $ofd.Title  = 'Select the Excel workbook containing domain data'
    $ofd.Filter = 'Excel workbooks (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All files (*.*)|*.*'
    if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Err 'No file selected — exiting.'; exit 1
    }
    $XlsxPath = $ofd.FileName
}

if (-not (Test-Path $XlsxPath)) { Write-Err "File not found: $XlsxPath"; exit 1 }

if (-not $OutputPath) { $OutputPath = Join-Path $scriptDir 'domains.json' }

Write-Step "Reading: $XlsxPath"

# ── Load ImportExcel (auto-install) or fall back to COM ───────────────────────
$rows = $null
$haveImportExcel = $false
try {
    if (-not (Get-Module -ListAvailable -Name ImportExcel -ErrorAction SilentlyContinue)) {
        Write-Warn "ImportExcel not installed — installing from PSGallery (CurrentUser)..."
        Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
    Import-Module ImportExcel -ErrorAction Stop
    $haveImportExcel = $true
} catch {
    Write-Warn "ImportExcel unavailable ($($_.Exception.Message)) — trying Excel COM..."
}

if ($haveImportExcel) {
    $params = @{ Path = $XlsxPath }
    if ($SheetName) { $params.WorksheetName = $SheetName }
    $rows = @(Import-Excel @params)
    Write-Ok "ImportExcel: read $($rows.Count) row(s)"
} else {
    # COM fallback
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    try {
        $wb = $excel.Workbooks.Open((Resolve-Path $XlsxPath).Path)
        $ws = if ($SheetName) { $wb.Sheets.Item($SheetName) } else { $wb.Sheets.Item(1) }
        $used = $ws.UsedRange
        $rc = $used.Rows.Count; $cc = $used.Columns.Count

        # First row = headers
        $headers = @{}
        for ($c = 1; $c -le $cc; $c++) {
            $h = [string]$used.Cells.Item(1,$c).Text
            $headers[$h] = $c
        }

        $rows = [System.Collections.Generic.List[object]]::new()
        for ($r = 2; $r -le $rc; $r++) {
            $ht = [ordered]@{}
            foreach ($h in $headers.Keys) {
                $ht[$h] = [string]$used.Cells.Item($r, $headers[$h]).Text
            }
            $rows.Add([PSCustomObject]$ht) | Out-Null
        }

        $wb.Close($false)
        Write-Ok "COM: read $($rows.Count) row(s)"
    } finally {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

# ── Show available columns if the expected ones are missing ───────────────────
if ($rows.Count -gt 0) {
    $available = $rows[0].PSObject.Properties.Name
    $missingDomain = $DomainColumn -notin $available
    $missingVbu    = $VbuIdColumn  -and ($VbuIdColumn -notin $available)

    if ($missingDomain -or $missingVbu) {
        Write-Host ''
        Write-Host '  Available columns in this sheet:' -ForegroundColor Yellow
        $available | ForEach-Object { Write-Host "    - $_" -ForegroundColor Gray }
        Write-Host ''
        if ($missingDomain) {
            Write-Err "Column '$DomainColumn' not found. Use -DomainColumn to specify the correct name."
        }
        if ($missingVbu) {
            Write-Warn "Column '$VbuIdColumn' not found. VBU ID will be left blank. Use -VbuIdColumn '' to suppress this warning."
        }
        if ($missingDomain) { exit 1 }
    }
}

# ── Build entries ──────────────────────────────────────────────────────────────
Write-Step "Mapping columns: Domain='$DomainColumn'  VBU ID='$(if($VbuIdColumn){"$VbuIdColumn"}else{"(none)"})'"

$seen    = @{}
$entries = [System.Collections.Generic.List[object]]::new()
$skipped = 0

foreach ($row in $rows) {
    $domain = ([string]($row.$DomainColumn)).Trim().ToLower().TrimStart('@')
    if ([string]::IsNullOrWhiteSpace($domain)) { $skipped++; continue }

    # Basic sanity: must look like a domain (contains a dot)
    if ($domain -notmatch '\.') {
        Write-Warn "Skipping '$domain' — doesn't look like a domain name"
        $skipped++; continue
    }

    if ($seen.ContainsKey($domain)) {
        Write-Warn "Duplicate domain '$domain' — keeping first occurrence"
        $skipped++; continue
    }
    $seen[$domain] = $true

    $vbu = ''
    if ($VbuIdColumn -and $row.PSObject.Properties[$VbuIdColumn]) {
        $vbu = ([string]($row.$VbuIdColumn)).Trim()
    }

    $entries.Add([ordered]@{ domain = $domain; vbuId = $vbu }) | Out-Null
}

Write-Ok "$($entries.Count) valid domain(s) mapped  ($skipped skipped)"

# ── Merge with existing file if -Append ───────────────────────────────────────
if ($Append -and (Test-Path $OutputPath)) {
    try {
        $existing = @(Get-Content $OutputPath -Raw -Encoding UTF8 | ConvertFrom-Json)
        $existingMap = @{}
        foreach ($e in $existing) { if ($e.domain) { $existingMap[$e.domain.ToLower()] = $e } }

        $newCount = 0
        foreach ($e in $entries) {
            if (-not $existingMap.ContainsKey($e.domain)) {
                $existingMap[$e.domain] = $e
                $newCount++
            }
        }
        $entries = @($existingMap.Values | Sort-Object { $_.domain })
        Write-Ok "Merged: $newCount new + $($existingMap.Count - $newCount) existing = $($entries.Count) total"
    } catch {
        Write-Warn "Could not merge with existing domains.json — overwriting: $($_.Exception.Message)"
    }
}

# ── Write output ──────────────────────────────────────────────────────────────
Write-Step "Writing: $OutputPath"
$entries | Sort-Object { $_.domain } | ConvertTo-Json -Depth 2 |
    Set-Content -Path $OutputPath -Encoding UTF8 -Force
Write-Ok "Done — $($entries.Count) domain(s) written to $(Split-Path $OutputPath -Leaf)"

Write-Host ''
Write-Host "  Open domains.json to verify, then launch the Discovery menu." -ForegroundColor White
Write-Host ''
