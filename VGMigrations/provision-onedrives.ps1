#Requires -Version 7.0
<#
.SYNOPSIS
    provision-onedrives.ps1 — Pre-provision OneDrive for Business sites from a
    mapping file (CSV or Excel).

.DESCRIPTION
    GUI wrapper that reads a list of destination UPNs from a mapping file and
    calls Request-SPOPersonalSite in batches of 200 via a background runspace.
    Requires Microsoft.Online.SharePoint.PowerShell (PnP or classic SPO module).

.NOTES
    Dependency : lib.ps1 (colours, fonts, helpers), settings.ps1 (Show-SettingsDialog)
    Log file   : logs\provision-onedrives-<timestamp>.log

    Change log
    ----------
    2026-05-27  Added settings.ps1 load + gear icon so the settings dialog is
                reachable from within this screen.
                Replaced the tenant-prefix ComboBox with a plain TextBox that
                auto-fills from the SharePointAdminUrl stored in shared config
                (Settings > Customer tab).  The resolved-URL label was removed.
                Improved startup logging and error trapping throughout.
#>

$libPath      = Join-Path $PSScriptRoot 'lib.ps1'
$settingsPath = Join-Path $PSScriptRoot 'settings.ps1'

# ── File logging — set up BEFORE lib load so any lib error is captured ────────
$_logDir = Join-Path $PSScriptRoot 'logs'
if (-not (Test-Path $_logDir)) { New-Item -ItemType Directory -Path $_logDir -Force | Out-Null }
$script:LogFile = Join-Path $_logDir "provision-onedrives-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

function _RawLog {
    param([string]$Msg)
    "$(Get-Date -Format 'HH:mm:ss.fff') $Msg" | Add-Content -Path $script:LogFile -Encoding UTF8 -ErrorAction SilentlyContinue
}

_RawLog "=== provision-onedrives.ps1 started  PID=$PID  PSVersion=$($PSVersionTable.PSVersion) ==="
_RawLog "PSScriptRoot : $PSScriptRoot"
_RawLog "lib.ps1      : $libPath  exists=$(Test-Path $libPath)"
_RawLog "settings.ps1 : $settingsPath  exists=$(Test-Path $settingsPath)"

if (Test-Path $libPath) {
    try {
        . $libPath
        _RawLog "lib.ps1 loaded OK"
    } catch {
        _RawLog "lib.ps1 LOAD ERROR: $($_.Exception.Message)"
        _RawLog "Stack: $($_.ScriptStackTrace)"
    }
} else {
    _RawLog "lib.ps1 NOT FOUND — colours and helpers will be missing"
}

if (Test-Path $settingsPath) {
    try {
        . $settingsPath
        _RawLog "settings.ps1 loaded OK"
    } catch {
        _RawLog "settings.ps1 LOAD ERROR: $($_.Exception.Message)"
    }
} else {
    _RawLog "settings.ps1 NOT FOUND — settings dialog will be unavailable"
}

# ── Helper: read Excel file via COM (UI thread only) ─────────────────────────
function Read-ExcelFile {
    param([string]$Path, [string]$SheetName)
    $excel = $null; $book = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false
        $abs = (Resolve-Path $Path).Path
        $book = $excel.Workbooks.Open($abs, [Type]::Missing, $true)

        $vis = @()
        for ($s = 1; $s -le $book.Worksheets.Count; $s++) {
            $sh = $book.Worksheets.Item($s)
            if ($sh.Visible -eq -1) { $vis += $sh.Name }
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($sh)
        }
        Write-Log "Sheets found: $($vis -join ', ')"

        $chosen = $null
        if ($SheetName) {
            if ($vis -contains $SheetName) { $chosen = $SheetName }
            else { throw "Sheet '$SheetName' not found. Available: $($vis -join ', ')" }
        } elseif ($vis.Count -eq 1) {
            $chosen = $vis[0]
            Write-Log "Single sheet - using '$chosen'."
        } else {
            $chosen = Show-SheetPicker $vis
            if (-not $chosen) { throw 'No sheet selected.' }
        }
        Write-Log "Reading sheet '$chosen'..."

        $sheet = $book.Worksheets.Item($chosen)
        $data = $sheet.UsedRange.Value2
        $rc   = $sheet.UsedRange.Rows.Count
        $cc   = $sheet.UsedRange.Columns.Count
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)

        if ($rc -lt 2) { throw "Sheet '$chosen' contains no data rows." }

        $headers = 1..$cc | ForEach-Object {
            if ($null -ne $data[1, $_]) { [string]$data[1, $_] } else { "Column$_" }
        }
        $list = [System.Collections.Generic.List[object]]::new()
        for ($r = 2; $r -le $rc; $r++) {
            $obj = [ordered]@{}
            for ($c = 1; $c -le $cc; $c++) { $obj[$headers[$c - 1]] = $data[$r, $c] }
            $list.Add([PSCustomObject]$obj) | Out-Null
        }
        Write-Log "Excel: $($list.Count) data row(s) loaded."
        return @($list)
    } catch {
        Write-Log "Excel read failed: $($_.Exception.Message)" 'ERROR'
        throw
    } finally {
        if ($book)  { try { $book.Close($false) }  catch {}; [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
        if ($excel) { try { $excel.Quit() }          catch {}; [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
    }
}

# ── Helper: sheet picker dialog ───────────────────────────────────────────────
function Show-SheetPicker {
    param([string[]]$Sheets)
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Select Sheet'; $dlg.ClientSize = [System.Drawing.Size]::new(320, 140)
    $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false
    $dlg.BackColor = $clrBg

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = 'Select the sheet to read:'
    $lbl.Location = [System.Drawing.Point]::new(16, 16); $lbl.AutoSize = $true
    $lbl.ForeColor = $clrText; $dlg.Controls.Add($lbl)

    $cmb = New-Object System.Windows.Forms.ComboBox
    $cmb.Location = [System.Drawing.Point]::new(16, 40); $cmb.Size = [System.Drawing.Size]::new(284, 24)
    $cmb.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $Sheets | ForEach-Object { [void]$cmb.Items.Add($_) }
    if ($cmb.Items.Count -gt 0) { $cmb.SelectedIndex = 0 }
    $dlg.Controls.Add($cmb)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'OK'; $btnOk.Location = [System.Drawing.Point]::new(116, 84)
    $btnOk.Size = [System.Drawing.Size]::new(80, 28); $btnOk.Font = $FontBold
    $btnOk.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnOk.FlatAppearance.BorderSize = 0
    $btnOk.BackColor = $clrAccent; $btnOk.ForeColor = [System.Drawing.Color]::White
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOk); $dlg.AcceptButton = $btnOk

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $cmb.SelectedItem }
    return $null
}

# ── Helper: read mapping file and return valid UPN list ───────────────────────
function Read-MappingFile {
    param([string]$MappingFile, [string]$ColumnName, [string]$SheetName)
    $ext = [System.IO.Path]::GetExtension($MappingFile).ToLowerInvariant()
    $rows = $null

    if ($ext -eq '.csv') {
        $rows = @(Import-Csv -Path $MappingFile -Encoding UTF8)
        Write-Log "CSV loaded: $($rows.Count) row(s)."
    } elseif ($ext -in '.xlsx', '.xlsm', '.xls') {
        $rows = Read-ExcelFile $MappingFile $SheetName
    } else {
        throw "Unsupported file type '$ext'. Use .csv, .xlsx, .xlsm, or .xls."
    }

    if (-not $rows -or $rows.Count -eq 0) { throw 'No rows found in file.' }

    $cols = @($rows[0].PSObject.Properties.Name)
    Write-Log "Columns: $($cols -join ', ')"

    $candidates = @('Destination', 'DestinationUPN', 'DestinationUserUPN',
                    'DestinationUserPrincipalName', 'TargetUPN', 'Target')
    $col = $null
    if ($ColumnName) {
        if ($cols -contains $ColumnName) { $col = $ColumnName }
        else { throw "Column '$ColumnName' not found. Available: $($cols -join ', ')" }
    } else {
        foreach ($c in $candidates) { if ($cols -contains $c) { $col = $c; break } }
        if (-not $col) {
            throw "Could not auto-detect UPN column. Tried: $($candidates -join ', '). Use the Column override field."
        }
    }
    Write-Log "Using column: '$col'"

    $upns = @(
        $rows |
        ForEach-Object { $_.$col } |
        ForEach-Object { if ($_) { $_.ToString().Trim() } } |
        Where-Object   { $_ -match '^[^@]+@[^@]+\.[^@]+$' } |
        Select-Object  -Unique
    )
    $dropped = $rows.Count - $upns.Count
    if ($dropped -gt 0) { Write-Log "$dropped row(s) had blank or invalid UPNs and were skipped." 'WARN' }
    return $upns
}

# ── Main GUI ──────────────────────────────────────────────────────────────────
function Show-ProvisionUI {
    _RawLog "Show-ProvisionUI entered"

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'Provision OneDrives'
    $form.ClientSize      = [System.Drawing.Size]::new(640, 620)
    $form.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.BackColor       = $clrBg
    $form.Font            = $FontBody
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.MaximizeBox     = $false
    $_ico = Join-Path $PSScriptRoot 'FlyMigration.ico'
    if (Test-Path $_ico) { $form.Icon = [System.Drawing.Icon]::new($_ico) }

    # ── Header ────────────────────────────────────────────────────────────────
    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Size = [System.Drawing.Size]::new(640, 56); $hdr.Dock = [System.Windows.Forms.DockStyle]::Top
    $hdr.BackColor = $clrAccent; $form.Controls.Add($hdr)
    $_hdrX = Add-HeaderLogo $hdr 36
    $hdrLbl = New-Object System.Windows.Forms.Label
    $hdrLbl.Text = '  Provision OneDrives'; $hdrLbl.Font = $FontTitle
    $hdrLbl.ForeColor = [System.Drawing.Color]::White
    $hdrLbl.Location  = [System.Drawing.Point]::new($_hdrX, 0)
    $hdrLbl.Size      = [System.Drawing.Size]::new(540, 56)
    $hdrLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdr.Controls.Add($hdrLbl)

    $btnGear = New-Object System.Windows.Forms.Button
    $btnGear.BackColor = $clrAccent; $btnGear.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnGear.FlatAppearance.BorderSize = 0
    $btnGear.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(0, 78, 152)
    $btnGear.Size = [System.Drawing.Size]::new(38, 38); $btnGear.Location = [System.Drawing.Point]::new(594, 9)
    $btnGear.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnGear.Add_Click({
        _RawLog "Settings button clicked"
        if (Get-Command Show-SettingsDialog -ErrorAction SilentlyContinue) { Show-SettingsDialog }
        else { [System.Windows.Forms.MessageBox]::Show('Settings are not available.','Settings','OK','Information') | Out-Null }
    })
    if ($script:GearBitmap) { $btnGear.Image = $script:GearBitmap; $btnGear.ImageAlign = [System.Drawing.ContentAlignment]::MiddleCenter }
    else { $btnGear.Text = [char]0x2699; $btnGear.Font = New-Object System.Drawing.Font('Segoe UI', 16); $btnGear.ForeColor = [System.Drawing.Color]::White }
    $hdr.Controls.Add($btnGear)

    # ── Footer ────────────────────────────────────────────────────────────────
    $footer = New-Object System.Windows.Forms.Panel
    $footer.Height = 46; $footer.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $footer.BackColor = [System.Drawing.Color]::FromArgb(20, 24, 38)
    $form.Controls.Add($footer)
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'; $btnClose.Size = [System.Drawing.Size]::new(90, 30)
    $btnClose.Location = [System.Drawing.Point]::new(534, 8)
    $btnClose.BackColor = [System.Drawing.Color]::FromArgb(200, 55, 55)
    $btnClose.ForeColor = [System.Drawing.Color]::White; $btnClose.Font = $FontBold
    $btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnClose.FlatAppearance.BorderSize = 0
    $btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnClose.Add_Click({ $form.Close() }.GetNewClosure())
    $footer.Controls.Add($btnClose)
    $footer.Add_SizeChanged({ $btnClose.Left = $footer.Width - 100 }.GetNewClosure())

    # ── Input card ────────────────────────────────────────────────────────────
    $card = New-Object System.Windows.Forms.Panel
    $card.Location  = [System.Drawing.Point]::new(12, 66)
    $card.Size      = [System.Drawing.Size]::new(616, 206)
    $card.BackColor = $clrPanel
    $form.Controls.Add($card)

    $lx = 16; $ex = 130; $y = 14

    # Mapping file
    $lbFile = New-Object System.Windows.Forms.Label
    $lbFile.Text = 'Mapping file:'; $lbFile.Font = $FontBold; $lbFile.ForeColor = $clrText
    $lbFile.Location = [System.Drawing.Point]::new($lx, $y + 5); $lbFile.AutoSize = $true
    $card.Controls.Add($lbFile)
    $txtFile = New-Object System.Windows.Forms.TextBox
    $txtFile.Location = [System.Drawing.Point]::new($ex, $y); $txtFile.Size = [System.Drawing.Size]::new(370, 24)
    $txtFile.Font = $FontBody; $txtFile.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $txtFile.ReadOnly = $true; $txtFile.BackColor = [System.Drawing.Color]::FromArgb(245, 246, 250)
    $card.Controls.Add($txtFile)
    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = 'Browse...'; $btnBrowse.Location = [System.Drawing.Point]::new($ex + 376, $y - 2)
    $btnBrowse.Size = [System.Drawing.Size]::new(88, 28); $btnBrowse.Font = $FontBold
    $btnBrowse.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnBrowse.FlatAppearance.BorderSize = 0
    $btnBrowse.BackColor = [System.Drawing.Color]::FromArgb(225, 228, 238); $btnBrowse.ForeColor = $clrText
    $btnBrowse.Cursor = [System.Windows.Forms.Cursors]::Hand
    $card.Controls.Add($btnBrowse)
    $y += 38

    # Admin URL — pre-filled from Settings > Customer > SharePoint Admin URL
    $lbAdmin = New-Object System.Windows.Forms.Label
    $lbAdmin.Text = 'SPO Admin URL:'; $lbAdmin.Font = $FontBold; $lbAdmin.ForeColor = $clrText
    $lbAdmin.Location = [System.Drawing.Point]::new($lx, $y + 5); $lbAdmin.AutoSize = $true
    $card.Controls.Add($lbAdmin)
    $txtAdmin = New-Object System.Windows.Forms.TextBox
    $txtAdmin.Location = [System.Drawing.Point]::new($ex, $y); $txtAdmin.Size = [System.Drawing.Size]::new(464, 24)
    $txtAdmin.Font = $FontBody; $txtAdmin.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    try { $txtAdmin.Text = (Read-SharedConfig).SharePointAdminUrl } catch {}
    $card.Controls.Add($txtAdmin)
    $y += 38

    # Column + Sheet on same row
    $lbCol = New-Object System.Windows.Forms.Label
    $lbCol.Text = 'Column (opt):'; $lbCol.ForeColor = $clrText
    $lbCol.Location = [System.Drawing.Point]::new($lx, $y + 5); $lbCol.AutoSize = $true
    $card.Controls.Add($lbCol)
    $txtCol = New-Object System.Windows.Forms.TextBox
    $txtCol.Location = [System.Drawing.Point]::new($ex, $y); $txtCol.Size = [System.Drawing.Size]::new(180, 24)
    $txtCol.Font = $FontBody; $txtCol.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    try { $txtCol.PlaceholderText = 'auto-detect' } catch {}
    $card.Controls.Add($txtCol)

    $lbSheet = New-Object System.Windows.Forms.Label
    $lbSheet.Text = 'Sheet (opt):'; $lbSheet.ForeColor = $clrText
    $lbSheet.Location = [System.Drawing.Point]::new(328, $y + 5); $lbSheet.AutoSize = $true
    $card.Controls.Add($lbSheet)
    $txtSheet = New-Object System.Windows.Forms.TextBox
    $txtSheet.Location = [System.Drawing.Point]::new(410, $y); $txtSheet.Size = [System.Drawing.Size]::new(180, 24)
    $txtSheet.Font = $FontBody; $txtSheet.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    try { $txtSheet.PlaceholderText = 'auto (Excel only)' } catch {}
    $card.Controls.Add($txtSheet)
    $y += 38

    # WhatIf + Run button
    $chkWhatIf = New-Object System.Windows.Forms.CheckBox
    $chkWhatIf.Text = 'WhatIf  (preview only - no requests submitted)'
    $chkWhatIf.Location = [System.Drawing.Point]::new($lx, $y + 5)
    $chkWhatIf.AutoSize = $true; $chkWhatIf.ForeColor = $clrText
    $card.Controls.Add($chkWhatIf)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = 'Run Provisioning'; $btnRun.Location = [System.Drawing.Point]::new(448, $y)
    $btnRun.Size = [System.Drawing.Size]::new(152, 32); $btnRun.Font = $FontBold
    $btnRun.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnRun.FlatAppearance.BorderSize = 0
    $btnRun.BackColor = $clrGreen; $btnRun.ForeColor = [System.Drawing.Color]::White
    $btnRun.Cursor = [System.Windows.Forms.Cursors]::Hand
    $card.Controls.Add($btnRun)
    $y += 40

    # Status label
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = 'Select a mapping file, then click Run Provisioning.'
    $lblStatus.ForeColor = $clrMuted; $lblStatus.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $lblStatus.Location = [System.Drawing.Point]::new($lx, $y)
    $lblStatus.Size = [System.Drawing.Size]::new(590, 18)
    $card.Controls.Add($lblStatus)

    # ── Progress bar ──────────────────────────────────────────────────────────
    $progress = New-Object System.Windows.Forms.ProgressBar
    $progress.Location = [System.Drawing.Point]::new(12, $card.Bottom + 8)
    $progress.Size     = [System.Drawing.Size]::new(616, 8)
    $progress.Style    = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $form.Controls.Add($progress)

    # ── Log RTB ───────────────────────────────────────────────────────────────
    $script:rtbLog = New-Object System.Windows.Forms.RichTextBox
    $script:rtbLog.Location    = [System.Drawing.Point]::new(12, $progress.Bottom + 8)
    $script:rtbLog.Size        = [System.Drawing.Size]::new(616, $form.ClientSize.Height - $progress.Bottom - 8 - 46 - 8)
    $script:rtbLog.BackColor   = $clrLogBg
    $script:rtbLog.ForeColor   = [System.Drawing.Color]::FromArgb(200, 210, 230)
    $script:rtbLog.Font        = $FontMono
    $script:rtbLog.ReadOnly    = $true
    $script:rtbLog.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $script:rtbLog.ScrollBars  = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    $form.Controls.Add($script:rtbLog)

    # ── Runspace state ────────────────────────────────────────────────────────
    $script:provTimer     = $null
    $script:provRunspace  = $null
    $script:provPS        = $null
    $script:provHandle    = $null
    $script:tickBusy      = $false

    # ── Browse handler ────────────────────────────────────────────────────────
    $btnBrowse.Add_Click({
        _RawLog "Browse dialog opened"
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Title  = 'Select mapping file'
        $ofd.Filter = 'Excel and CSV files (*.xlsx;*.xlsm;*.xls;*.csv)|*.xlsx;*.xlsm;*.xls;*.csv|All files (*.*)|*.*'
        $ofd.InitialDirectory = $PSScriptRoot
        if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            _RawLog "Mapping file selected: $($ofd.FileName)"
            $txtFile.Text = $ofd.FileName
        } else {
            _RawLog "Browse dialog cancelled"
        }
    }.GetNewClosure())

    # ── Run handler ───────────────────────────────────────────────────────────
    $btnRun.Add_Click({
        Write-Log "Run clicked — File='$($txtFile.Text)'  AdminUrl='$($txtAdmin.Text.Trim())'  WhatIf=$($chkWhatIf.Checked)"
        if (-not $txtFile.Text) {
            Write-Log 'Validation failed: no mapping file selected' 'WARN'
            [System.Windows.Forms.MessageBox]::Show('Please select a mapping file.', 'Missing Input', 'OK', 'Warning') | Out-Null; return
        }
        $adminUrl = $txtAdmin.Text.Trim().TrimEnd('/')
        if (-not $adminUrl) {
            Write-Log 'Validation failed: no SPO Admin URL set' 'WARN'
            [System.Windows.Forms.MessageBox]::Show('Please enter the SPO Admin URL.  Set it in Settings > Customer > SharePoint Admin URL.', 'Missing Input', 'OK', 'Warning') | Out-Null; return
        }
        Write-Log "Validation passed — AdminUrl=$adminUrl"

        $btnRun.Enabled = $false; $btnBrowse.Enabled = $false; $btnClose.Enabled = $false
        $progress.Value = 0
        _RawLog "Buttons locked — reading mapping file"
        $lblStatus.Text = 'Reading mapping file...'
        [System.Windows.Forms.Application]::DoEvents()

        Write-Log '=== Provision OneDrives started ==='
        Write-Log "File     : $($txtFile.Text)"
        Write-Log "Admin URL: $adminUrl"

        # Read file on UI thread (Excel COM requires STA / main thread)
        $upns = $null
        try {
            $upns = Read-MappingFile $txtFile.Text $txtCol.Text.Trim() $txtSheet.Text.Trim()
        } catch {
            Write-Log "File read failed: $($_.Exception.Message)" 'ERROR'
            $lblStatus.Text = 'File read failed - check log'; $btnRun.Enabled = $true; $btnBrowse.Enabled = $true; $btnClose.Enabled = $true; return
        }

        if ($upns.Count -eq 0) {
            Write-Log 'No valid UPNs found after filtering.' 'ERROR'
            $lblStatus.Text = 'No valid UPNs found'; $btnRun.Enabled = $true; $btnBrowse.Enabled = $true; $btnClose.Enabled = $true; return
        }

        Write-Log "$($upns.Count) unique UPN(s) ready."
        $first5 = $upns | Select-Object -First 5
        $first5 | ForEach-Object { Write-Log "  $_" }
        if ($upns.Count -gt 5) { Write-Log "  ... and $($upns.Count - 5) more" }

        # WhatIf short-circuit
        if ($chkWhatIf.Checked) {
            Write-Log "WhatIf: would submit $($upns.Count) UPN(s) to $adminUrl - no changes made." 'WARN'
            Write-Log '=== WhatIf complete ===' 'OK'
            $progress.Value = 100
            $lblStatus.Text = "WhatIf: $($upns.Count) UPN(s) would be submitted"
            $btnRun.Enabled = $true; $btnBrowse.Enabled = $true; $btnClose.Enabled = $true; return
        }

        # Set up shared state for background runspace
        $rs = [hashtable]::Synchronized(@{
            Done         = $false
            FatalError   = $null
            Submitted    = 0
            Failed       = 0
            BatchesDone  = 0
            TotalBatches = [math]::Ceiling($upns.Count / 200)
            LogQueue     = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())
        })

        $logFilePath = $script:LogFile

        # Build runspace
        $script:provRunspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $script:provRunspace.ApartmentState = 'STA'
        $script:provRunspace.ThreadOptions  = 'ReuseThread'
        $script:provRunspace.Open()

        $script:provPS = [System.Management.Automation.PowerShell]::Create()
        $script:provPS.Runspace = $script:provRunspace

        [void]$script:provPS.AddScript({
            param($adminUrl, $upns, $rs, $logFilePath)

            function QLog {
                param([string]$Msg, [string]$Level = 'INFO')
                $rs.LogQueue.Enqueue("[$Level] $Msg")
                try { "$(Get-Date -Format 'HH:mm:ss') [$Level] $Msg" | Add-Content $logFilePath -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
            }

            try {
                QLog 'Checking for Microsoft.Online.SharePoint.PowerShell module...'
                $mod = Get-Module -ListAvailable -Name 'Microsoft.Online.SharePoint.PowerShell' -ErrorAction SilentlyContinue
                if (-not $mod) {
                    throw 'Microsoft.Online.SharePoint.PowerShell is not installed. Run: Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser'
                }
                Import-Module 'Microsoft.Online.SharePoint.PowerShell' -DisableNameChecking -ErrorAction Stop
                QLog "SPO module $($mod.Version) loaded." 'OK'

                QLog "Connecting to $adminUrl - sign in when the browser window opens..."
                Connect-SPOService -Url $adminUrl -ErrorAction Stop
                QLog 'Connected to SharePoint Online.' 'OK'

                $batchSize = 200
                $submitted = 0; $failed = 0
                for ($i = 0; $i -lt $upns.Count; $i += $batchSize) {
                    $end   = [math]::Min($i + $batchSize, $upns.Count)
                    $chunk = $upns[$i..($end - 1)]
                    $batch = [math]::Floor($i / $batchSize) + 1
                    QLog "Batch $batch / $($rs.TotalBatches): submitting $($chunk.Count) UPN(s)..."
                    try {
                        Request-SPOPersonalSite -UserEmails $chunk -ErrorAction Stop
                        $submitted      += $chunk.Count
                        $rs.Submitted    = $submitted
                        $rs.BatchesDone  = $batch
                        QLog "Batch $batch accepted." 'OK'
                    } catch {
                        $failed       += $chunk.Count
                        $rs.Failed     = $failed
                        QLog "Batch $batch failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'
                    }
                }

                QLog "Submission complete. Submitted: $submitted  Failed: $failed"
                if ($failed -gt 0) { QLog "$failed UPN(s) were in failed batches - check log for details." 'WARN' }
            } catch {
                $rs.FatalError = $_.Exception.Message
                QLog "Fatal error: $($_.Exception.Message)" 'ERROR'
            } finally {
                $rs.Done = $true
            }
        })
        [void]$script:provPS.AddParameters(@{
            adminUrl    = $adminUrl
            upns        = $upns
            rs          = $rs
            logFilePath = $logFilePath
        })

        $script:provHandle = $script:provPS.BeginInvoke()
        $lblStatus.Text = 'Connecting to SPO - sign in when the browser opens...'

        # Poll timer
        $script:provTimer = New-Object System.Windows.Forms.Timer
        $script:provTimer.Interval = 300
        $rs2 = $rs
        $script:provTimer.Add_Tick({
            if ($script:tickBusy) { return }
            $script:tickBusy = $true
            try {
                while ($rs2.LogQueue.Count -gt 0) {
                    $raw   = $rs2.LogQueue.Dequeue()
                    $level = if ($raw -match '^\[(\w+)\]') { $matches[1] } else { 'INFO' }
                    $text  = $raw -replace '^\[\w+\] ', ''
                    Write-Log $text $level
                }
                if ($rs2.TotalBatches -gt 0 -and $rs2.BatchesDone -gt 0) {
                    $pct = [int]([math]::Min(($rs2.BatchesDone / $rs2.TotalBatches) * 99, 99))
                    $progress.Value = $pct
                    $lblStatus.Text = "Batch $($rs2.BatchesDone) / $($rs2.TotalBatches) - $($rs2.Submitted) submitted"
                }
                if ($rs2.Done) {
                    $script:provTimer.Stop()
                    $progress.Value = 100
                    if ($rs2.FatalError) {
                        $lblStatus.Text = "Failed - $($rs2.FatalError)"
                        Write-Log '=== Provisioning failed ===' 'ERROR'
                    } else {
                        $lblStatus.Text = "Done - $($rs2.Submitted) submitted, $($rs2.Failed) failed"
                        Write-Log '=== Provisioning complete ===' 'OK'
                    }
                    $btnRun.Enabled = $true; $btnBrowse.Enabled = $true; $btnClose.Enabled = $true
                    try { $script:provRunspace.Close(); $script:provRunspace.Dispose() } catch {}
                }
            } finally {
                $script:tickBusy = $false
            }
        }.GetNewClosure())
        $script:provTimer.Start()
    }.GetNewClosure())

    # ── Cleanup on close ──────────────────────────────────────────────────────
    $form.Add_FormClosing({
        param($s, $e)
        _RawLog "FormClosing event  timerRunning=$($script:provTimer -and $script:provTimer.Enabled)"
        if ($script:provTimer -and $script:provTimer.Enabled) {
            $result = [System.Windows.Forms.MessageBox]::Show(
                'Provisioning is still running. Close anyway?', 'In Progress', 'YesNo', 'Warning')
            if ($result -eq [System.Windows.Forms.DialogResult]::No) {
                _RawLog "FormClosing cancelled — user chose to keep running"
                $e.Cancel = $true; return
            }
            _RawLog "FormClosing confirmed by user despite running timer"
        }
        _RawLog "FormClosing: cleaning up timer/PS/runspace"
        if ($script:provTimer)    { try { $script:provTimer.Stop(); $script:provTimer.Dispose() }    catch { _RawLog "Timer dispose error: $_" } }
        if ($script:provPS)       { try { $script:provPS.Stop();    $script:provPS.Dispose() }        catch { _RawLog "PS dispose error: $_" } }
        if ($script:provRunspace) { try { $script:provRunspace.Close(); $script:provRunspace.Dispose() } catch { _RawLog "Runspace dispose error: $_" } }
        _RawLog "FormClosing cleanup complete"
    }.GetNewClosure())

    $form.Add_Shown({
        $form.BringToFront(); $form.Activate()
    }.GetNewClosure())

    _RawLog "Calling Application.Run"
    [System.Windows.Forms.Application]::Run($form)
    _RawLog "Application.Run returned (form closed)"
}

_RawLog "=== Calling Show-ProvisionUI ==="
Write-Log '=== provision-onedrives.ps1 starting ==='
try {
    Show-ProvisionUI
    _RawLog "Show-ProvisionUI returned normally"
} catch {
    _RawLog "Show-ProvisionUI THREW: $($_.Exception.GetType().FullName): $($_.Exception.Message)"
    _RawLog "Stack: $($_.ScriptStackTrace)"
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    [System.Windows.Forms.MessageBox]::Show(
        "Failed to start:`n$($_.Exception.Message)", 'Launch Error', 'OK', 'Error') | Out-Null
}
_RawLog "Script exiting"
