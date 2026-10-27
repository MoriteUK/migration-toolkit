#Requires -Version 7.0
<#
.SYNOPSIS
    Compares two versions of an assessment Excel file and highlights all differences.

.DESCRIPTION
    GUI tool that compares two Excel workbooks (typically different versions of the same assessment)
    and highlights all changed cells in Turquoise (Accent 4) color for easy identification.

    Features:
    - Sheet-by-sheet comparison
    - Cell-by-cell difference detection
    - Highlights changes in Turquoise (Accent 4)
    - Saves comparison result as new file
    - Summary report of changes found

.NOTES
    Requires ImportExcel module
#>

[CmdletBinding()]
param()

# Load WinForms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ── Logging setup ─────────────────────────────────────────────────────────────
$script:LogDir = Join-Path $env:APPDATA 'FlyMigration\Logs'
if (-not (Test-Path $script:LogDir)) { New-Item -ItemType Directory -Path $script:LogDir -Force | Out-Null }
$script:LogFile = Join-Path $script:LogDir "Compare-Assessments-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

function Write-Log {
    param([string]$Msg, [string]$Level = 'INFO')
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff')  [$($Level.PadRight(5))]  $Msg"
    Add-Content -Path $script:LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
}

Write-Log "=== Compare-Assessments.ps1 started ==="

# ── Design System (matching lib.ps1) ─────────────────────────────────────────
$script:clrBg          = [System.Drawing.Color]::FromArgb(248, 249, 250)
$script:clrAccent      = [System.Drawing.Color]::FromArgb(38, 114, 236)
$script:clrAccentHover = [System.Drawing.Color]::FromArgb(30, 95, 200)
$script:clrSuccess     = [System.Drawing.Color]::FromArgb(0, 130, 70)
$script:clrDanger      = [System.Drawing.Color]::FromArgb(220, 53, 69)
$script:clrText        = [System.Drawing.Color]::FromArgb(33, 37, 41)
$script:clrMuted       = [System.Drawing.Color]::FromArgb(108, 117, 125)
$script:clrBorder      = [System.Drawing.Color]::FromArgb(222, 226, 230)

$script:FontTitle = New-Object System.Drawing.Font('Segoe UI Light', 18)
$script:FontBody  = New-Object System.Drawing.Font('Segoe UI', 10)
$script:FontBold  = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
$script:FontSmall = New-Object System.Drawing.Font('Segoe UI', 8.5)

# ── Main Form ─────────────────────────────────────────────────────────────────
function Show-ComparisonUI {
    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'Compare Assessment Versions - Migration Toolkit'
    $form.ClientSize      = [System.Drawing.Size]::new(900, 600)
    $form.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.BackColor       = $script:clrBg
    $form.Font            = $script:FontBody
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.MaximizeBox     = $false

    # Icon
    $icoPath = Join-Path $PSScriptRoot 'FlyMigration.ico'
    if (Test-Path $icoPath) { $form.Icon = [System.Drawing.Icon]::new($icoPath) }

    # ── Header ────────────────────────────────────────────────────────────────
    $header = New-Object System.Windows.Forms.Panel
    $header.Size      = [System.Drawing.Size]::new(900, 60)
    $header.Dock      = [System.Windows.Forms.DockStyle]::Top
    $header.BackColor = $script:clrAccent
    $form.Controls.Add($header)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text      = '  Compare Assessment Versions'
    $lblTitle.Font      = $script:FontTitle
    $lblTitle.ForeColor = [System.Drawing.Color]::White
    $lblTitle.Location  = [System.Drawing.Point]::new(16, 0)
    $lblTitle.Size      = [System.Drawing.Size]::new(860, 60)
    $lblTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $header.Controls.Add($lblTitle)

    $y = 80

    # ── File 1 (Baseline) ─────────────────────────────────────────────────────
    $lblFile1 = New-Object System.Windows.Forms.Label
    $lblFile1.Text     = 'Baseline File (Original):'
    $lblFile1.Font     = $script:FontBold
    $lblFile1.Location = [System.Drawing.Point]::new(20, $y)
    $lblFile1.AutoSize = $true
    $form.Controls.Add($lblFile1)
    $y += 26

    $txtFile1 = New-Object System.Windows.Forms.TextBox
    $txtFile1.Location = [System.Drawing.Point]::new(20, $y)
    $txtFile1.Size     = [System.Drawing.Size]::new(750, 24)
    $txtFile1.Font     = $script:FontBody
    $txtFile1.ReadOnly = $true
    $form.Controls.Add($txtFile1)

    $btnBrowse1 = New-Object System.Windows.Forms.Button
    $btnBrowse1.Text       = 'Browse...'
    $btnBrowse1.Location   = [System.Drawing.Point]::new(780, $y - 2)
    $btnBrowse1.Size       = [System.Drawing.Size]::new(100, 28)
    $btnBrowse1.BackColor  = $script:clrAccent
    $btnBrowse1.ForeColor  = [System.Drawing.Color]::White
    $btnBrowse1.FlatStyle  = [System.Windows.Forms.FlatStyle]::Flat
    $btnBrowse1.Font       = $script:FontBold
    $btnBrowse1.Cursor     = [System.Windows.Forms.Cursors]::Hand
    $btnBrowse1.Add_Click({
        $openFile = New-Object System.Windows.Forms.OpenFileDialog
        $openFile.Filter = 'Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*'
        $openFile.Title  = 'Select Baseline Assessment File'
        $openFile.InitialDirectory = 'C:\Users\Andy White\Volaris Group\GRP Data Security (Volaris Consolidated) - M365 Migrations\2. InProgress Migrations'
        if ($openFile.ShowDialog() -eq 'OK') {
            $txtFile1.Text = $openFile.FileName
            Write-Log "Baseline file selected: $($openFile.FileName)"
        }
    }.GetNewClosure())
    $form.Controls.Add($btnBrowse1)
    $y += 38

    # ── File 2 (Comparison) ───────────────────────────────────────────────────
    $lblFile2 = New-Object System.Windows.Forms.Label
    $lblFile2.Text     = 'Comparison File (Updated):'
    $lblFile2.Font     = $script:FontBold
    $lblFile2.Location = [System.Drawing.Point]::new(20, $y)
    $lblFile2.AutoSize = $true
    $form.Controls.Add($lblFile2)
    $y += 26

    $txtFile2 = New-Object System.Windows.Forms.TextBox
    $txtFile2.Location = [System.Drawing.Point]::new(20, $y)
    $txtFile2.Size     = [System.Drawing.Size]::new(750, 24)
    $txtFile2.Font     = $script:FontBody
    $txtFile2.ReadOnly = $true
    $form.Controls.Add($txtFile2)

    $btnBrowse2 = New-Object System.Windows.Forms.Button
    $btnBrowse2.Text       = 'Browse...'
    $btnBrowse2.Location   = [System.Drawing.Point]::new(780, $y - 2)
    $btnBrowse2.Size       = [System.Drawing.Size]::new(100, 28)
    $btnBrowse2.BackColor  = $script:clrAccent
    $btnBrowse2.ForeColor  = [System.Drawing.Color]::White
    $btnBrowse2.FlatStyle  = [System.Windows.Forms.FlatStyle]::Flat
    $btnBrowse2.Font       = $script:FontBold
    $btnBrowse2.Cursor     = [System.Windows.Forms.Cursors]::Hand
    $btnBrowse2.Add_Click({
        $openFile = New-Object System.Windows.Forms.OpenFileDialog
        $openFile.Filter = 'Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*'
        $openFile.Title  = 'Select Comparison Assessment File'
        $openFile.InitialDirectory = 'C:\Users\Andy White\Volaris Group\GRP Data Security (Volaris Consolidated) - M365 Migrations\2. InProgress Migrations'
        if ($openFile.ShowDialog() -eq 'OK') {
            $txtFile2.Text = $openFile.FileName
            Write-Log "Comparison file selected: $($openFile.FileName)"
        }
    }.GetNewClosure())
    $form.Controls.Add($btnBrowse2)
    $y += 38

    # ── Output File ───────────────────────────────────────────────────────────
    $lblOutput = New-Object System.Windows.Forms.Label
    $lblOutput.Text     = 'Output File (Comparison Result):'
    $lblOutput.Font     = $script:FontBold
    $lblOutput.Location = [System.Drawing.Point]::new(20, $y)
    $lblOutput.AutoSize = $true
    $form.Controls.Add($lblOutput)
    $y += 26

    $txtOutput = New-Object System.Windows.Forms.TextBox
    $txtOutput.Location = [System.Drawing.Point]::new(20, $y)
    $txtOutput.Size     = [System.Drawing.Size]::new(750, 24)
    $txtOutput.Font     = $script:FontBody
    $form.Controls.Add($txtOutput)

    $btnBrowseOut = New-Object System.Windows.Forms.Button
    $btnBrowseOut.Text       = 'Browse...'
    $btnBrowseOut.Location   = [System.Drawing.Point]::new(780, $y - 2)
    $btnBrowseOut.Size       = [System.Drawing.Size]::new(100, 28)
    $btnBrowseOut.BackColor  = $script:clrAccent
    $btnBrowseOut.ForeColor  = [System.Drawing.Color]::White
    $btnBrowseOut.FlatStyle  = [System.Windows.Forms.FlatStyle]::Flat
    $btnBrowseOut.Font       = $script:FontBold
    $btnBrowseOut.Cursor     = [System.Windows.Forms.Cursors]::Hand
    $btnBrowseOut.Add_Click({
        $saveFile = New-Object System.Windows.Forms.SaveFileDialog
        $saveFile.Filter = 'Excel Files (*.xlsx)|*.xlsx'
        $saveFile.Title  = 'Save Comparison Result'
        $saveFile.FileName = "Comparison-$(Get-Date -Format 'yyyyMMdd-HHmm').xlsx"
        if ($saveFile.ShowDialog() -eq 'OK') {
            $txtOutput.Text = $saveFile.FileName
            Write-Log "Output file set: $($saveFile.FileName)"
        }
    }.GetNewClosure())
    $form.Controls.Add($btnBrowseOut)
    $y += 48

    # ── Progress ──────────────────────────────────────────────────────────────
    $lblProgress = New-Object System.Windows.Forms.Label
    $lblProgress.Text     = 'Ready to compare'
    $lblProgress.Font     = $script:FontBold
    $lblProgress.Location = [System.Drawing.Point]::new(20, $y)
    $lblProgress.Size     = [System.Drawing.Size]::new(860, 20)
    $form.Controls.Add($lblProgress)
    $y += 26

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = [System.Drawing.Point]::new(20, $y)
    $progressBar.Size     = [System.Drawing.Size]::new(860, 24)
    $progressBar.Style    = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $form.Controls.Add($progressBar)
    $y += 34

    # ── Log Output ────────────────────────────────────────────────────────────
    $lblLog = New-Object System.Windows.Forms.Label
    $lblLog.Text     = 'Comparison Log:'
    $lblLog.Font     = $script:FontBold
    $lblLog.Location = [System.Drawing.Point]::new(20, $y)
    $lblLog.AutoSize = $true
    $form.Controls.Add($lblLog)
    $y += 22

    $txtLog = New-Object System.Windows.Forms.TextBox
    $txtLog.Location   = [System.Drawing.Point]::new(20, $y)
    $txtLog.Size       = [System.Drawing.Size]::new(860, 200)
    $txtLog.Multiline  = $true
    $txtLog.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $txtLog.Font       = New-Object System.Drawing.Font('Consolas', 9)
    $txtLog.ReadOnly   = $true
    $txtLog.BackColor  = [System.Drawing.Color]::White
    $form.Controls.Add($txtLog)
    $y += 210

    # ── Action Buttons ────────────────────────────────────────────────────────
    $btnCompare = New-Object System.Windows.Forms.Button
    $btnCompare.Text      = 'Compare Files'
    $btnCompare.Location  = [System.Drawing.Point]::new(20, $y)
    $btnCompare.Size      = [System.Drawing.Size]::new(180, 36)
    $btnCompare.BackColor = $script:clrSuccess
    $btnCompare.ForeColor = [System.Drawing.Color]::White
    $btnCompare.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCompare.Font      = $script:FontBold
    $btnCompare.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $btnCompare.Add_Click({
        Start-Comparison `
            -File1 $txtFile1.Text `
            -File2 $txtFile2.Text `
            -OutputFile $txtOutput.Text `
            -LogTextBox $txtLog `
            -ProgressBar $progressBar `
            -ProgressLabel $lblProgress `
            -CompareButton $btnCompare
    }.GetNewClosure())
    $form.Controls.Add($btnCompare)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text      = 'Close'
    $btnClose.Location  = [System.Drawing.Point]::new(790, $y)
    $btnClose.Size      = [System.Drawing.Size]::new(90, 36)
    $btnClose.BackColor = $script:clrDanger
    $btnClose.ForeColor = [System.Drawing.Color]::White
    $btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnClose.Font      = $script:FontBold
    $btnClose.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $btnClose.Add_Click({ $form.Close() }.GetNewClosure())
    $form.Controls.Add($btnClose)

    # ── Show Form ─────────────────────────────────────────────────────────────
    $form.Add_Shown({
        Log-ToUI -Message 'Excel Comparison Tool' -LogTextBox $txtLog
        Log-ToUI -Message "Log file: $script:LogFile" -LogTextBox $txtLog
        Log-ToUI -Message '' -LogTextBox $txtLog
        Log-ToUI -Message 'Instructions:' -LogTextBox $txtLog
        Log-ToUI -Message '1. Select baseline file (original version)' -LogTextBox $txtLog
        Log-ToUI -Message '2. Select comparison file (updated version)' -LogTextBox $txtLog
        Log-ToUI -Message '3. Specify output file location' -LogTextBox $txtLog
        Log-ToUI -Message '4. Click "Compare Files"' -LogTextBox $txtLog
        Log-ToUI -Message '' -LogTextBox $txtLog
        Log-ToUI -Message 'All differences will be highlighted in Turquoise (Accent 4)' -LogTextBox $txtLog
        Log-ToUI -Message '' -LogTextBox $txtLog
    }.GetNewClosure())

    [void]$form.ShowDialog()
}

# ── Helper Functions ──────────────────────────────────────────────────────────
function Log-ToUI {
    param(
        [string]$Message,
        [System.Windows.Forms.TextBox]$LogTextBox
    )

    $LogTextBox.AppendText("$Message`r`n")
    $LogTextBox.SelectionStart = $LogTextBox.Text.Length
    $LogTextBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Start-Comparison {
    param(
        [string]$File1,
        [string]$File2,
        [string]$OutputFile,
        [System.Windows.Forms.TextBox]$LogTextBox,
        [System.Windows.Forms.ProgressBar]$ProgressBar,
        [System.Windows.Forms.Label]$ProgressLabel,
        [System.Windows.Forms.Button]$CompareButton
    )

    # Validate inputs
    if ([string]::IsNullOrWhiteSpace($File1) -or -not (Test-Path $File1)) {
        [System.Windows.Forms.MessageBox]::Show('Please select a valid baseline file.', 'Invalid File', 'OK', 'Warning') | Out-Null
        return
    }

    if ([string]::IsNullOrWhiteSpace($File2) -or -not (Test-Path $File2)) {
        [System.Windows.Forms.MessageBox]::Show('Please select a valid comparison file.', 'Invalid File', 'OK', 'Warning') | Out-Null
        return
    }

    if ([string]::IsNullOrWhiteSpace($OutputFile)) {
        [System.Windows.Forms.MessageBox]::Show('Please specify an output file location.', 'Invalid Output', 'OK', 'Warning') | Out-Null
        return
    }

    # Check ImportExcel module
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Log-ToUI -Message '[ERROR] ImportExcel module not found!' -LogTextBox $LogTextBox
        Log-ToUI -Message '        Install with: Install-Module ImportExcel -Scope CurrentUser' -LogTextBox $LogTextBox
        [System.Windows.Forms.MessageBox]::Show(
            "ImportExcel module not found.`n`nInstall it with:`nInstall-Module ImportExcel -Scope CurrentUser",
            'Module Not Found',
            'OK',
            'Error'
        ) | Out-Null
        return
    }

    $CompareButton.Enabled = $false
    $ProgressLabel.Text = 'Starting comparison...'
    Log-ToUI -Message '' -LogTextBox $LogTextBox
    Log-ToUI -Message '=== Starting Comparison ===' -LogTextBox $LogTextBox
    Log-ToUI -Message "Baseline:   $File1" -LogTextBox $LogTextBox
    Log-ToUI -Message "Comparison: $File2" -LogTextBox $LogTextBox
    Log-ToUI -Message "Output:     $OutputFile" -LogTextBox $LogTextBox
    Log-ToUI -Message '' -LogTextBox $LogTextBox
    Write-Log "Comparison started: $File1 vs $File2"

    try {
        Import-Module ImportExcel -ErrorAction Stop

        # Copy File2 to output location (we'll mark up this file)
        $ProgressLabel.Text = 'Creating output file...'
        Copy-Item -Path $File2 -Destination $OutputFile -Force
        Log-ToUI -Message 'Output file created (copy of comparison file)' -LogTextBox $LogTextBox
        Write-Log "Output file created: $OutputFile"

        # Open both files
        $ProgressLabel.Text = 'Opening Excel files...'
        Log-ToUI -Message 'Opening baseline file...' -LogTextBox $LogTextBox
        $baseline = Import-Excel -Path $File1 -WorksheetName * -AsHashtable -ErrorAction Stop
        Log-ToUI -Message "  Loaded $($baseline.Keys.Count) worksheet(s)" -LogTextBox $LogTextBox

        Log-ToUI -Message 'Opening comparison file...' -LogTextBox $LogTextBox
        $comparison = Import-Excel -Path $File2 -WorksheetName * -AsHashtable -ErrorAction Stop
        Log-ToUI -Message "  Loaded $($comparison.Keys.Count) worksheet(s)" -LogTextBox $LogTextBox

        # Open output file for marking
        $excel = Open-ExcelPackage -Path $OutputFile

        $stats = @{
            TotalSheets = 0
            ComparedSheets = 0
            TotalCellsCompared = 0
            DifferentCells = 0
            SheetsWithDifferences = 0
        }

        # Turquoise, Accent 4 color
        $turquoise = [System.Drawing.Color]::FromArgb(0, 176, 240)

        # Compare each worksheet
        $allSheets = @($baseline.Keys) + @($comparison.Keys) | Select-Object -Unique
        $stats.TotalSheets = $allSheets.Count
        $ProgressBar.Maximum = $stats.TotalSheets
        $ProgressBar.Value = 0

        foreach ($sheetName in $allSheets) {
            $stats.TotalSheets++
            $ProgressBar.Value++
            $ProgressLabel.Text = "Comparing worksheet: $sheetName"
            [System.Windows.Forms.Application]::DoEvents()

            Log-ToUI -Message "Comparing sheet: $sheetName" -LogTextBox $LogTextBox
            Write-Log "Comparing sheet: $sheetName"

            # Check if sheet exists in both files
            if (-not $baseline.ContainsKey($sheetName)) {
                Log-ToUI -Message "  [NEW] Sheet exists only in comparison file" -LogTextBox $LogTextBox
                Write-Log "  Sheet only in File2: $sheetName"
                continue
            }

            if (-not $comparison.ContainsKey($sheetName)) {
                Log-ToUI -Message "  [REMOVED] Sheet exists only in baseline file" -LogTextBox $LogTextBox
                Write-Log "  Sheet only in File1: $sheetName"
                continue
            }

            $stats.ComparedSheets++

            # Get the worksheet from output file
            $ws = $excel.Workbook.Worksheets[$sheetName]
            if (-not $ws) {
                Log-ToUI -Message "  [SKIP] Worksheet not found in output file" -LogTextBox $LogTextBox
                continue
            }

            $baseData = $baseline[$sheetName]
            $compData = $comparison[$sheetName]

            $sheetDiffs = 0

            # Compare row counts
            $baseRowCount = if ($baseData) { @($baseData).Count } else { 0 }
            $compRowCount = if ($compData) { @($compData).Count } else { 0 }

            Log-ToUI -Message "  Rows: Baseline=$baseRowCount, Comparison=$compRowCount" -LogTextBox $LogTextBox

            # Get all column names from both datasets
            $allColumns = [System.Collections.Generic.HashSet[string]]::new()
            if ($baseData -and $baseData.Count -gt 0) {
                foreach ($row in @($baseData)) {
                    if ($row -is [hashtable] -or $row -is [System.Collections.Specialized.OrderedDictionary]) {
                        foreach ($key in $row.Keys) {
                            [void]$allColumns.Add($key)
                        }
                    } else {
                        foreach ($prop in $row.PSObject.Properties) {
                            [void]$allColumns.Add($prop.Name)
                        }
                    }
                }
            }

            if ($compData -and $compData.Count -gt 0) {
                foreach ($row in @($compData)) {
                    if ($row -is [hashtable] -or $row -is [System.Collections.Specialized.OrderedDictionary]) {
                        foreach ($key in $row.Keys) {
                            [void]$allColumns.Add($key)
                        }
                    } else {
                        foreach ($prop in $row.PSObject.Properties) {
                            [void]$allColumns.Add($prop.Name)
                        }
                    }
                }
            }

            # Compare cell by cell
            $maxRows = [Math]::Max($baseRowCount, $compRowCount)

            for ($rowIdx = 0; $rowIdx -lt $maxRows; $rowIdx++) {
                $excelRow = $rowIdx + 2  # +1 for 0-based, +1 for header

                $baseRow = if ($rowIdx -lt $baseRowCount) { @($baseData)[$rowIdx] } else { $null }
                $compRow = if ($rowIdx -lt $compRowCount) { @($compData)[$rowIdx] } else { $null }

                $colIdx = 1
                foreach ($colName in $allColumns) {
                    $stats.TotalCellsCompared++

                    $baseValue = if ($baseRow) {
                        if ($baseRow -is [hashtable] -or $baseRow -is [System.Collections.Specialized.OrderedDictionary]) {
                            if ($baseRow.ContainsKey($colName)) { $baseRow[$colName] } else { $null }
                        } else {
                            $baseRow.$colName
                        }
                    } else { $null }

                    $compValue = if ($compRow) {
                        if ($compRow -is [hashtable] -or $compRow -is [System.Collections.Specialized.OrderedDictionary]) {
                            if ($compRow.ContainsKey($colName)) { $compRow[$colName] } else { $null }
                        } else {
                            $compRow.$colName
                        }
                    } else { $null }

                    # Convert to strings for comparison
                    $baseStr = if ($null -eq $baseValue) { '' } else { [string]$baseValue }
                    $compStr = if ($null -eq $compValue) { '' } else { [string]$compValue }

                    # Compare values
                    if ($baseStr -ne $compStr) {
                        $stats.DifferentCells++
                        $sheetDiffs++

                        # Highlight the cell in turquoise
                        try {
                            $cell = $ws.Cells[$excelRow, $colIdx]
                            $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                            $cell.Style.Fill.BackgroundColor.SetColor($turquoise)
                        }
                        catch {
                            Write-Log "  Error highlighting cell [$excelRow,$colIdx]: $($_.Exception.Message)" 'WARN'
                        }
                    }

                    $colIdx++
                }
            }

            if ($sheetDiffs -gt 0) {
                $stats.SheetsWithDifferences++
                Log-ToUI -Message "  [CHANGES] $sheetDiffs cell(s) different" -LogTextBox $LogTextBox
                Write-Log "  Sheet $sheetName: $sheetDiffs differences"
            } else {
                Log-ToUI -Message "  [IDENTICAL] No differences found" -LogTextBox $LogTextBox
                Write-Log "  Sheet $sheetName: identical"
            }
        }

        # Save the marked-up file
        $ProgressLabel.Text = 'Saving comparison result...'
        Log-ToUI -Message '' -LogTextBox $LogTextBox
        Log-ToUI -Message 'Saving marked-up file...' -LogTextBox $LogTextBox
        Close-ExcelPackage $excel -SaveAs $OutputFile
        Write-Log "Comparison file saved: $OutputFile"

        # Summary
        Log-ToUI -Message '' -LogTextBox $LogTextBox
        Log-ToUI -Message '=== Comparison Complete ===' -LogTextBox $LogTextBox
        Log-ToUI -Message "Total Worksheets:      $($stats.TotalSheets)" -LogTextBox $LogTextBox
        Log-ToUI -Message "Compared Worksheets:   $($stats.ComparedSheets)" -LogTextBox $LogTextBox
        Log-ToUI -Message "Sheets with Changes:   $($stats.SheetsWithDifferences)" -LogTextBox $LogTextBox
        Log-ToUI -Message "Total Cells Compared:  $($stats.TotalCellsCompared)" -LogTextBox $LogTextBox
        Log-ToUI -Message "Different Cells:       $($stats.DifferentCells)" -LogTextBox $LogTextBox
        Log-ToUI -Message '' -LogTextBox $LogTextBox
        Log-ToUI -Message "Output saved: $OutputFile" -LogTextBox $LogTextBox
        Log-ToUI -Message '' -LogTextBox $LogTextBox
        Log-ToUI -Message 'All differences are highlighted in Turquoise (Accent 4)' -LogTextBox $LogTextBox

        $ProgressLabel.Text = 'Comparison complete'
        Write-Log "Comparison complete: $($stats.DifferentCells) differences found"

        [System.Windows.Forms.MessageBox]::Show(
            "Comparison complete!`n`nSheets Compared: $($stats.ComparedSheets)`nCells Different: $($stats.DifferentCells)`n`nOutput: $OutputFile",
            'Comparison Complete',
            'OK',
            'Information'
        ) | Out-Null

        # Ask to open the file
        $openResult = [System.Windows.Forms.MessageBox]::Show(
            'Would you like to open the comparison file now?',
            'Open File',
            'YesNo',
            'Question'
        )

        if ($openResult -eq 'Yes') {
            Start-Process $OutputFile
        }
    }
    catch {
        Log-ToUI -Message "[ERROR] $($_.Exception.Message)" -LogTextBox $LogTextBox
        $ProgressLabel.Text = 'Error occurred'
        Write-Log "Comparison failed: $($_.Exception.Message)" 'ERROR'
        [System.Windows.Forms.MessageBox]::Show(
            "An error occurred:`n`n$($_.Exception.Message)",
            'Error',
            'OK',
            'Error'
        ) | Out-Null
    }
    finally {
        $CompareButton.Enabled = $true
    }
}

# ── Entry Point ───────────────────────────────────────────────────────────────
try {
    Show-ComparisonUI
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    [System.Windows.Forms.MessageBox]::Show(
        "Fatal error:`n`n$($_.Exception.Message)`n`nLog: $script:LogFile",
        'Fatal Error',
        'OK',
        'Error'
    ) | Out-Null
}

Write-Log "=== Compare-Assessments.ps1 exiting ==="
