#Requires -Version 7.0
<#
.SYNOPSIS
    Debug script to show what's in the Tenant IDs.xlsx file
.DESCRIPTION
    Reads and displays all rows and columns from Tenant IDs.xlsx for debugging
#>

param(
    [string]$TenantIDsPath = "C:\Users\Andy White\Volaris Group\GRP Data Security (Volaris Consolidated) - M365 Migrations\Tenant IDs.xlsx"
)

Write-Host "=== Tenant IDs Debug Tool ===" -ForegroundColor Cyan
Write-Host "File: $TenantIDsPath" -ForegroundColor White
Write-Host ""

if (-not (Test-Path $TenantIDsPath)) {
    Write-Host "ERROR: File not found!" -ForegroundColor Red
    exit 1
}

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open($TenantIDsPath)
    $worksheet = $workbook.Worksheets.Item(1)
    $usedRange = $worksheet.UsedRange

    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count

    Write-Host "Total Rows: $rowCount" -ForegroundColor Green
    Write-Host "Total Columns: $colCount" -ForegroundColor Green
    Write-Host ""

    # Show headers
    Write-Host "=== Column Headers ===" -ForegroundColor Cyan
    for ($col = 1; $col -le $colCount; $col++) {
        $header = $usedRange.Cells.Item(1, $col).Text
        $colLetter = [char](64 + $col)
        Write-Host "Column $colLetter ($col): $header" -ForegroundColor Yellow
    }
    Write-Host ""

    # Show each data row
    Write-Host "=== Data Rows ===" -ForegroundColor Cyan
    for ($row = 2; $row -le $rowCount; $row++) {
        $colA = $usedRange.Cells.Item($row, 1).Text
        $colJ = $usedRange.Cells.Item($row, 10).Text

        # Show if row is empty
        if ([string]::IsNullOrWhiteSpace($colA)) {
            Write-Host "Row $row : [EMPTY ROW]" -ForegroundColor Gray
        } else {
            # Color code based on column J value (only skip if exactly "Yes")
            $color = if ($colJ.ToLower() -eq "yes") { "Red" } else { "Green" }
            Write-Host "Row $row : " -NoNewline
            Write-Host "$colA" -ForegroundColor White -NoNewline
            Write-Host " | Column J = " -NoNewline
            Write-Host "'$colJ'" -ForegroundColor $color -NoNewline

            if ($colJ.ToLower() -eq "yes") {
                Write-Host " [WILL SKIP]" -ForegroundColor Red
            } else {
                Write-Host " [WILL CHECK]" -ForegroundColor Green
            }
        }
    }

    Write-Host ""
    Write-Host "=== Summary ===" -ForegroundColor Cyan

    $totalRows = $rowCount - 1  # Exclude header
    $toCheck = 0
    $toSkip = 0
    $empty = 0

    for ($row = 2; $row -le $rowCount; $row++) {
        $colA = $usedRange.Cells.Item($row, 1).Text
        $colJ = $usedRange.Cells.Item($row, 10).Text

        if ([string]::IsNullOrWhiteSpace($colA)) {
            $empty++
        } elseif ($colJ -eq "YES") {
            $toSkip++
        } else {
            $toCheck++
        }
    }

    Write-Host "Total rows in file: $totalRows" -ForegroundColor White
    Write-Host "Rows to check (no YES in column J): $toCheck" -ForegroundColor Green
    Write-Host "Rows to skip (YES in column J): $toSkip" -ForegroundColor Red
    Write-Host "Empty rows: $empty" -ForegroundColor Gray

} catch {
    Write-Host "ERROR: $_" -ForegroundColor Red
} finally {
    if ($workbook) {
        $workbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "=== Debug Complete ===" -ForegroundColor Cyan
