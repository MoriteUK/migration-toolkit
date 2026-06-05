#Requires -Version 7.0
<#
.SYNOPSIS
    Hide-AddressBook.ps1 — Sets HiddenFromAddressListsEnabled = $true on Exchange
    Online recipients discovered by Search-M365Domain.ps1.
    Reads the same Discovery CSV files as remove-domain.ps1.

.NOTES
    Requires : ExchangeOnlineManagement
    CSVs used: 02_Mailboxes, 03_DistributionGroups, 04_MailContacts,
               05_SharedMailboxes, 06_M365Groups
#>

$script:RootDir = $PSScriptRoot

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ── File logging ──────────────────────────────────────────────────────────────
$_logDir = Join-Path $script:RootDir 'logs'
if (-not (Test-Path $_logDir)) { New-Item -ItemType Directory -Path $_logDir -Force | Out-Null }
$script:LogFile = Join-Path $_logDir "hide-addressbook-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

function _RawLog {
    param([string]$Msg)
    "$(Get-Date -Format 'HH:mm:ss.fff') $Msg" | Add-Content -Path $script:LogFile -Encoding UTF8 -ErrorAction SilentlyContinue
}

_RawLog "=== Hide-AddressBook.ps1 started  PID=$PID  PSVersion=$($PSVersionTable.PSVersion) ==="
_RawLog "PSScriptRoot : $script:RootDir"

$libPath      = Join-Path $script:RootDir 'lib.ps1'
$settingsPath = Join-Path $script:RootDir 'settings.ps1'
$_libLoaded   = $false

if (Test-Path $libPath) {
    try { . $libPath; $_libLoaded = $true; _RawLog 'lib.ps1 loaded OK' }
    catch { _RawLog "lib.ps1 LOAD ERROR: $($_.Exception.Message)" }
}
if (Test-Path $settingsPath) {
    try { . $settingsPath; _RawLog 'settings.ps1 loaded OK' }
    catch { _RawLog "settings.ps1 LOAD ERROR: $($_.Exception.Message)" }
}

if (-not $_libLoaded) {
    $clrBg     = [System.Drawing.Color]::FromArgb(240, 242, 247)
    $clrPanel  = [System.Drawing.Color]::White
    $clrAccent = [System.Drawing.Color]::FromArgb(0, 100, 180)
    $clrText   = [System.Drawing.Color]::FromArgb(28, 28, 32)
    $clrMuted  = [System.Drawing.Color]::FromArgb(100, 108, 120)
    $clrBorder = [System.Drawing.Color]::FromArgb(210, 215, 228)
    $clrLogBg  = [System.Drawing.Color]::FromArgb(26, 27, 38)
    $clrRed    = [System.Drawing.Color]::FromArgb(195, 30, 30)
    $FontBody  = New-Object System.Drawing.Font('Segoe UI', 9)
    $FontBold  = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
    $FontCap   = New-Object System.Drawing.Font('Segoe UI Semibold', 7.5)
    $FontMono  = New-Object System.Drawing.Font('Consolas', 8.5)
    $FontTitle = New-Object System.Drawing.Font('Segoe UI Semibold', 14)
    $script:GearBitmap = $null
    function Add-HeaderLogo { param($Header, [int]$LogoH = 34); return 8 }
    function Write-Log {
        param([string]$Msg, [string]$Level = 'INFO')
        $ts = Get-Date -Format 'HH:mm:ss'
        if ($script:LogFile) {
            try { "$ts [$Level] $Msg" | Add-Content -Path $script:LogFile -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
        }
        if ($script:rtbLog -and -not $script:rtbLog.IsDisposed) {
            $script:rtbLog.SelectionStart = $script:rtbLog.TextLength; $script:rtbLog.SelectionLength = 0
            $script:rtbLog.SelectionColor = [System.Drawing.Color]::FromArgb(80, 95, 120)
            $script:rtbLog.AppendText("$ts ")
            $lc = switch ($Level) {
                'OK'    { [System.Drawing.Color]::FromArgb(65, 195, 110)  }
                'WARN'  { [System.Drawing.Color]::FromArgb(220, 165, 45)  }
                'ERROR' { [System.Drawing.Color]::FromArgb(225, 80, 80)   }
                default { [System.Drawing.Color]::FromArgb(120, 155, 220) }
            }
            $script:rtbLog.SelectionColor = $lc; $script:rtbLog.AppendText("[$Level] ")
            $script:rtbLog.SelectionColor = [System.Drawing.Color]::FromArgb(205, 212, 230)
            $script:rtbLog.AppendText("$Msg`n"); $script:rtbLog.ScrollToCaret()
        }
    }
}

# ── Section definitions — only the object types that support HiddenFromAddressListsEnabled ──
$script:SectionDefs = @(
    [pscustomobject]@{ CsvName='02_Mailboxes.csv';          Label='Mailboxes';             KeyField='UserPrincipalName' }
    [pscustomobject]@{ CsvName='03_DistributionGroups.csv'; Label='Distribution Groups';   KeyField='PrimarySmtpAddress' }
    [pscustomobject]@{ CsvName='04_MailContacts.csv';       Label='Mail Contacts';         KeyField='ExternalEmailAddress' }
    [pscustomobject]@{ CsvName='05_SharedMailboxes.csv';    Label='Shared Mailboxes';      KeyField='PrimarySmtpAddress' }
    [pscustomobject]@{ CsvName='06_M365Groups.csv';         Label='M365 Groups (+ Teams)'; KeyField='PrimarySmtpAddress' }
)

function Show-HideAddressBookUI {
    $rootDir = $script:RootDir
    $secDefs = $script:SectionDefs

    # ── Form ──────────────────────────────────────────────────────────────────
    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'Hide Recipients from Address Book'
    $form.ClientSize      = [System.Drawing.Size]::new(680, 1080)
    $form.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.BackColor       = $clrBg
    $form.Font            = $FontBody
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.MaximizeBox     = $false
    $_ico = Join-Path $rootDir 'FlyMigration.ico'
    if (Test-Path $_ico) { $form.Icon = [System.Drawing.Icon]::new($_ico) }

    # ── Header ────────────────────────────────────────────────────────────────
    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Size = [System.Drawing.Size]::new(680, 56); $hdr.Dock = [System.Windows.Forms.DockStyle]::Top
    $hdr.BackColor = $clrAccent; $form.Controls.Add($hdr)
    $_hdrX = Add-HeaderLogo $hdr 36
    $hdrLbl = New-Object System.Windows.Forms.Label
    $hdrLbl.Text = '  Hide Recipients from Address Book'; $hdrLbl.Font = $FontTitle
    $hdrLbl.ForeColor = [System.Drawing.Color]::White
    $hdrLbl.Location  = [System.Drawing.Point]::new($_hdrX, 0)
    $hdrLbl.Size      = [System.Drawing.Size]::new(580, 56)
    $hdrLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdr.Controls.Add($hdrLbl)

    $btnGear = New-Object System.Windows.Forms.Button
    $btnGear.BackColor = $clrAccent; $btnGear.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnGear.FlatAppearance.BorderSize = 0
    $btnGear.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(0, 78, 152)
    $btnGear.Size = [System.Drawing.Size]::new(38, 38); $btnGear.Location = [System.Drawing.Point]::new(634, 9)
    $btnGear.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnGear.Add_Click({
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
    $btnClose.Location = [System.Drawing.Point]::new(574, 8)
    $btnClose.BackColor = [System.Drawing.Color]::FromArgb(200, 55, 55)
    $btnClose.ForeColor = [System.Drawing.Color]::White; $btnClose.Font = $FontBold
    $btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnClose.FlatAppearance.BorderSize = 0
    $btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnClose.Add_Click({ $form.Close() }.GetNewClosure())
    $footer.Controls.Add($btnClose)
    $footer.Add_SizeChanged({ $btnClose.Left = $footer.Width - 100 }.GetNewClosure())

    # ── Card panel ────────────────────────────────────────────────────────────
    $card = New-Object System.Windows.Forms.Panel
    $card.Location  = [System.Drawing.Point]::new(12, 66)
    $card.Size      = [System.Drawing.Size]::new(656, 720)
    $card.BackColor = $clrPanel
    $form.Controls.Add($card)

    $lx = 14; $y = 12

    # Discovery folder
    $lblFolderCap = New-Object System.Windows.Forms.Label
    $lblFolderCap.Text = 'DISCOVERY FOLDER'; $lblFolderCap.Font = $FontCap; $lblFolderCap.ForeColor = $clrMuted
    $lblFolderCap.Location = [System.Drawing.Point]::new($lx, $y); $lblFolderCap.AutoSize = $true
    $card.Controls.Add($lblFolderCap); $y += 18

    $tbFolder = New-Object System.Windows.Forms.TextBox
    $tbFolder.Location    = [System.Drawing.Point]::new($lx, $y)
    $tbFolder.Size        = [System.Drawing.Size]::new(516, 24)
    $tbFolder.Font        = $FontBody
    $tbFolder.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $card.Controls.Add($tbFolder)

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = 'Browse...'; $btnBrowse.Location = [System.Drawing.Point]::new(538, $y - 2)
    $btnBrowse.Size = [System.Drawing.Size]::new(104, 28); $btnBrowse.Font = $FontBold
    $btnBrowse.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnBrowse.FlatAppearance.BorderSize = 0
    $btnBrowse.BackColor = [System.Drawing.Color]::FromArgb(225, 228, 238); $btnBrowse.ForeColor = $clrText
    $btnBrowse.Cursor = [System.Windows.Forms.Cursors]::Hand
    $card.Controls.Add($btnBrowse); $y += 32

    $btnScan = New-Object System.Windows.Forms.Button
    $btnScan.Text = 'Scan Folder'; $btnScan.Location = [System.Drawing.Point]::new($lx, $y)
    $btnScan.Size = [System.Drawing.Size]::new(120, 28); $btnScan.Font = $FontBold
    $btnScan.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnScan.FlatAppearance.BorderSize = 0
    $btnScan.BackColor = [System.Drawing.Color]::FromArgb(225, 228, 238); $btnScan.ForeColor = $clrText
    $btnScan.Cursor = [System.Windows.Forms.Cursors]::Hand
    $card.Controls.Add($btnScan)

    $lblScanStatus = New-Object System.Windows.Forms.Label
    $lblScanStatus.Text = 'Enter or browse to a Discovery folder, then click Scan.'
    $lblScanStatus.ForeColor = $clrMuted; $lblScanStatus.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $lblScanStatus.Location = [System.Drawing.Point]::new(142, $y + 8); $lblScanStatus.AutoSize = $true
    $card.Controls.Add($lblScanStatus); $y += 38

    $sep1 = New-Object System.Windows.Forms.Label
    $sep1.Location = [System.Drawing.Point]::new($lx, $y); $sep1.Size = [System.Drawing.Size]::new(628, 1)
    $sep1.BackColor = $clrBorder; $card.Controls.Add($sep1); $y += 10

    $lblSecCap = New-Object System.Windows.Forms.Label
    $lblSecCap.Text = 'SECTIONS TO PROCESS'; $lblSecCap.Font = $FontCap; $lblSecCap.ForeColor = $clrMuted
    $lblSecCap.Location = [System.Drawing.Point]::new($lx, $y); $lblSecCap.AutoSize = $true
    $card.Controls.Add($lblSecCap); $y += 18

    $clbSections = New-Object System.Windows.Forms.CheckedListBox
    $clbSections.Location     = [System.Drawing.Point]::new($lx, $y)
    $clbSections.Size         = [System.Drawing.Size]::new(628, 148)
    $clbSections.Font         = $FontBody
    $clbSections.BackColor    = [System.Drawing.Color]::FromArgb(250, 251, 253)
    $clbSections.ForeColor    = $clrText
    $clbSections.BorderStyle  = [System.Windows.Forms.BorderStyle]::FixedSingle
    $clbSections.CheckOnClick = $true
    $card.Controls.Add($clbSections)
    foreach ($sec in $secDefs) {
        [void]$clbSections.Items.Add(("{0,-38} (not scanned)" -f $sec.Label), $false)
    }
    $y += 153

    $btnSelAll = New-Object System.Windows.Forms.Button
    $btnSelAll.Text = 'Select All'; $btnSelAll.Location = [System.Drawing.Point]::new($lx, $y)
    $btnSelAll.Size = [System.Drawing.Size]::new(100, 26); $btnSelAll.Font = $FontBold
    $btnSelAll.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnSelAll.FlatAppearance.BorderSize = 0
    $btnSelAll.BackColor = [System.Drawing.Color]::FromArgb(225, 228, 238); $btnSelAll.ForeColor = $clrText
    $btnSelAll.Cursor = [System.Windows.Forms.Cursors]::Hand; $card.Controls.Add($btnSelAll)

    $btnDeselAll = New-Object System.Windows.Forms.Button
    $btnDeselAll.Text = 'Deselect All'; $btnDeselAll.Location = [System.Drawing.Point]::new($lx + 106, $y)
    $btnDeselAll.Size = [System.Drawing.Size]::new(100, 26); $btnDeselAll.Font = $FontBold
    $btnDeselAll.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnDeselAll.FlatAppearance.BorderSize = 0
    $btnDeselAll.BackColor = [System.Drawing.Color]::FromArgb(225, 228, 238); $btnDeselAll.ForeColor = $clrText
    $btnDeselAll.Cursor = [System.Windows.Forms.Cursors]::Hand; $card.Controls.Add($btnDeselAll)

    $lblSelStatus = New-Object System.Windows.Forms.Label
    $lblSelStatus.Text = '0 section(s) selected'; $lblSelStatus.ForeColor = $clrMuted
    $lblSelStatus.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $lblSelStatus.Location = [System.Drawing.Point]::new(280, $y + 7); $lblSelStatus.AutoSize = $true
    $card.Controls.Add($lblSelStatus); $y += 34

    $sep2 = New-Object System.Windows.Forms.Label
    $sep2.Location = [System.Drawing.Point]::new($lx, $y); $sep2.Size = [System.Drawing.Size]::new(628, 1)
    $sep2.BackColor = $clrBorder; $card.Controls.Add($sep2); $y += 10

    # ── Account List ──────────────────────────────────────────────────────────
    $lblAccountsCap = New-Object System.Windows.Forms.Label
    $lblAccountsCap.Text = 'ACCOUNTS TO BE HIDDEN'; $lblAccountsCap.Font = $FontCap; $lblAccountsCap.ForeColor = $clrMuted
    $lblAccountsCap.Location = [System.Drawing.Point]::new($lx, $y); $lblAccountsCap.AutoSize = $true
    $card.Controls.Add($lblAccountsCap); $y += 18

    $lvAccounts = New-Object System.Windows.Forms.ListView
    $lvAccounts.Location = [System.Drawing.Point]::new($lx, $y)
    $lvAccounts.Size = [System.Drawing.Size]::new(628, 200)
    $lvAccounts.View = [System.Windows.Forms.View]::Details
    $lvAccounts.FullRowSelect = $true
    $lvAccounts.GridLines = $true
    $lvAccounts.Font = $FontBody
    $lvAccounts.BackColor = [System.Drawing.Color]::FromArgb(250, 251, 253)
    $lvAccounts.ForeColor = $clrText
    $lvAccounts.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    [void]$lvAccounts.Columns.Add('Identity', 320)
    [void]$lvAccounts.Columns.Add('Type', 160)
    [void]$lvAccounts.Columns.Add('CSV File', 140)
    $card.Controls.Add($lvAccounts); $y += 205

    $lblAccountsCount = New-Object System.Windows.Forms.Label
    $lblAccountsCount.Text = '0 accounts listed'; $lblAccountsCount.ForeColor = $clrMuted
    $lblAccountsCount.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $lblAccountsCount.Location = [System.Drawing.Point]::new($lx, $y); $lblAccountsCount.AutoSize = $true
    $card.Controls.Add($lblAccountsCount); $y += 24

    $sep3 = New-Object System.Windows.Forms.Label
    $sep3.Location = [System.Drawing.Point]::new($lx, $y); $sep3.Size = [System.Drawing.Size]::new(628, 1)
    $sep3.BackColor = $clrBorder; $card.Controls.Add($sep3); $y += 12

    $chkWhatIf = New-Object System.Windows.Forms.CheckBox
    $chkWhatIf.Text = 'WhatIf  (list items only - no changes will be made)'
    $chkWhatIf.Location = [System.Drawing.Point]::new($lx, $y + 6)
    $chkWhatIf.AutoSize = $true; $chkWhatIf.ForeColor = $clrText
    $card.Controls.Add($chkWhatIf)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = 'Connect and Hide'; $btnRun.Location = [System.Drawing.Point]::new(470, $y)
    $btnRun.Size = [System.Drawing.Size]::new(172, 34); $btnRun.Font = $FontBold
    $btnRun.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnRun.FlatAppearance.BorderSize = 0
    $btnRun.BackColor = $clrRed; $btnRun.ForeColor = [System.Drawing.Color]::White
    $btnRun.Cursor = [System.Windows.Forms.Cursors]::Hand
    $card.Controls.Add($btnRun)

    # ── Progress bar ──────────────────────────────────────────────────────────
    $progress = New-Object System.Windows.Forms.ProgressBar
    $progress.Location = [System.Drawing.Point]::new(12, $card.Bottom + 8)
    $progress.Size     = [System.Drawing.Size]::new(656, 8)
    $progress.Style    = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $form.Controls.Add($progress)

    # ── Log RTB ───────────────────────────────────────────────────────────────
    $script:rtbLog = New-Object System.Windows.Forms.RichTextBox
    $script:rtbLog.Location    = [System.Drawing.Point]::new(12, $progress.Bottom + 8)
    $script:rtbLog.Size        = [System.Drawing.Size]::new(656, $form.ClientSize.Height - $progress.Bottom - 8 - 46 - 8)
    $script:rtbLog.BackColor   = $clrLogBg
    $script:rtbLog.ForeColor   = [System.Drawing.Color]::FromArgb(200, 210, 230)
    $script:rtbLog.Font        = $FontMono
    $script:rtbLog.ReadOnly    = $true
    $script:rtbLog.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $script:rtbLog.ScrollBars  = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    $form.Controls.Add($script:rtbLog)

    # ── State ─────────────────────────────────────────────────────────────────
    $script:rdTimer    = $null
    $script:rdRunspace = $null
    $script:rdPS       = $null
    $script:RowCounts  = @(0, 0, 0, 0, 0)
    $RowCounts = $script:RowCounts

    # ── Selection status and account list ─────────────────────────────────────
    $script:AllAccounts = [System.Collections.Generic.List[pscustomobject]]::new()

    $updateSelStatus = {
        $cnt = $clbSections.CheckedItems.Count
        $selItems = 0
        for ($i = 0; $i -lt $clbSections.Items.Count; $i++) {
            if ($clbSections.GetItemChecked($i)) { $selItems += $RowCounts[$i] }
        }
        $lblSelStatus.Text = "$cnt section(s) selected  ($selItems items)"

        # Update account list based on selected sections
        $lvAccounts.Items.Clear()
        $accountCount = 0
        for ($i = 0; $i -lt $clbSections.Items.Count; $i++) {
            if ($clbSections.GetItemChecked($i)) {
                $csvName = $secDefsLocal[$i].CsvName
                $typeLabel = $secDefsLocal[$i].Label

                # Get accounts from the stored list
                $accountsForSection = $script:AllAccounts | Where-Object { $_.CsvName -eq $csvName }
                foreach ($acc in $accountsForSection) {
                    $lvi = New-Object System.Windows.Forms.ListViewItem($acc.Identity)
                    [void]$lvi.SubItems.Add($typeLabel)
                    [void]$lvi.SubItems.Add($csvName)
                    [void]$lvAccounts.Items.Add($lvi)
                    $accountCount++
                }
            }
        }
        $lblAccountsCount.Text = "$accountCount account(s) listed"
    }.GetNewClosure()

    # ── Browse ────────────────────────────────────────────────────────────────
    $btnBrowse.Add_Click({
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description = 'Select the Discovery folder (or its parent output folder)'
        $fbd.ShowNewFolderButton = $false
        if ($tbFolder.Text -and (Test-Path $tbFolder.Text)) { $fbd.SelectedPath = $tbFolder.Text }
        if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Write-Log "Folder selected: $($fbd.SelectedPath)"
            $tbFolder.Text = $fbd.SelectedPath
        }
    }.GetNewClosure())

    # ── Scan ──────────────────────────────────────────────────────────────────
    $secDefsLocal = $secDefs
    $btnScan.Add_Click({
        $raw = $tbFolder.Text.Trim().Trim('"')
        if (-not $raw) {
            [System.Windows.Forms.MessageBox]::Show('Please enter or browse to a Discovery folder.','Missing Folder','OK','Warning') | Out-Null; return
        }
        if (-not (Test-Path $raw)) {
            [System.Windows.Forms.MessageBox]::Show("Path not found:`n$raw",'Not Found','OK','Warning') | Out-Null; return
        }
        $discFolder = $raw
        $candidate  = Join-Path $raw 'Discovery'
        if ((Split-Path $raw -Leaf) -ne 'Discovery' -and (Test-Path $candidate)) {
            $discFolder = $candidate; $tbFolder.Text = $discFolder
            Write-Log "Auto-resolved to Discovery subfolder: $discFolder"
        }
        Write-Log "Scanning: $discFolder"
        $lblScanStatus.Text = 'Scanning...'; [System.Windows.Forms.Application]::DoEvents()
        $script:AllAccounts.Clear()
        $found = 0
        for ($i = 0; $i -lt $secDefsLocal.Count; $i++) {
            $sec  = $secDefsLocal[$i]
            $path = Join-Path $discFolder $sec.CsvName
            $cnt  = 0
            if (Test-Path $path) {
                try {
                    $rows = @(Import-Csv -Path $path -Encoding UTF8)
                    $cnt = $rows.Count
                    # Store account identities for display
                    foreach ($row in $rows) {
                        $identity = $row.($sec.KeyField)
                        if ($identity) {
                            $script:AllAccounts.Add([pscustomobject]@{
                                CsvName  = $sec.CsvName
                                Identity = $identity
                            })
                        }
                    }
                }
                catch { Write-Log "Could not read $($sec.CsvName): $_" 'WARN' }
                $found++
                Write-Log "  $($sec.CsvName): $cnt row(s)"
            } else {
                Write-Log "  $($sec.CsvName): not found"
            }
            $RowCounts[$i] = $cnt
            $clbSections.Items[$i] = if ($cnt -gt 0) {
                "{0,-38} ({1} items)" -f $sec.Label, $cnt
            } elseif (Test-Path $path) {
                "{0,-38} (empty)" -f $sec.Label
            } else {
                "{0,-38} (CSV not found)" -f $sec.Label
            }
            $clbSections.SetItemChecked($i, ($cnt -gt 0))
        }
        $lblScanStatus.Text = "Scan complete — $found of $($secDefsLocal.Count) CSV(s) found"
        & $updateSelStatus
        Write-Log "Scan complete: $found CSV(s) found"
    }.GetNewClosure())

    # ── Select All / Deselect All ─────────────────────────────────────────────
    $btnSelAll.Add_Click({
        for ($i = 0; $i -lt $clbSections.Items.Count; $i++) { $clbSections.SetItemChecked($i, $true) }
        & $updateSelStatus
    }.GetNewClosure())
    $btnDeselAll.Add_Click({
        for ($i = 0; $i -lt $clbSections.Items.Count; $i++) { $clbSections.SetItemChecked($i, $false) }
        & $updateSelStatus
    }.GetNewClosure())
    $clbSections.Add_ItemCheck({ & $updateSelStatus }.GetNewClosure())

    # ── Run ───────────────────────────────────────────────────────────────────
    $secDefsRun = $secDefs
    $btnRun.Add_Click({
        $raw = $tbFolder.Text.Trim().Trim('"')
        if (-not $raw -or -not (Test-Path $raw)) {
            [System.Windows.Forms.MessageBox]::Show('Please scan a valid Discovery folder first.','No Folder','OK','Warning') | Out-Null; return
        }
        $discFolder = $raw
        $candidate  = Join-Path $raw 'Discovery'
        if ((Split-Path $raw -Leaf) -ne 'Discovery' -and (Test-Path $candidate)) { $discFolder = $candidate }

        $selectedCsvs = [System.Collections.Generic.List[string]]::new()
        $totalItems   = 0
        for ($i = 0; $i -lt $clbSections.Items.Count; $i++) {
            if ($clbSections.GetItemChecked($i)) {
                $selectedCsvs.Add($secDefsRun[$i].CsvName)
                $totalItems += $RowCounts[$i]
                Write-Log "  Selected: $($secDefsRun[$i].CsvName)  ($($RowCounts[$i]) item(s))"
            }
        }
        if ($selectedCsvs.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('No sections selected. Scan first and tick at least one section.','Nothing Selected','OK','Warning') | Out-Null; return
        }

        $whatIf = $chkWhatIf.Checked

        if (-not $whatIf) {
            $dlgC = New-Object System.Windows.Forms.Form
            $dlgC.Text = 'Confirm Hide'; $dlgC.ClientSize = [System.Drawing.Size]::new(480, 170)
            $dlgC.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
            $dlgC.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
            $dlgC.MaximizeBox = $false; $dlgC.MinimizeBox = $false; $dlgC.BackColor = $clrBg

            $lblCMsg = New-Object System.Windows.Forms.Label
            $lblCMsg.Text = "Hide $totalItems recipient(s) across $($selectedCsvs.Count) section(s) from the address book?`nThis CANNOT be undone. Type YES to confirm:"
            $lblCMsg.Location = [System.Drawing.Point]::new(16, 16); $lblCMsg.Size = [System.Drawing.Size]::new(448, 48)
            $lblCMsg.ForeColor = $clrText; $dlgC.Controls.Add($lblCMsg)

            $tbC = New-Object System.Windows.Forms.TextBox
            $tbC.Location = [System.Drawing.Point]::new(16, 70); $tbC.Size = [System.Drawing.Size]::new(448, 24)
            $tbC.Font = $FontBody; $tbC.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
            $dlgC.Controls.Add($tbC)

            $btnCP = New-Object System.Windows.Forms.Button
            $btnCP.Text = 'Proceed'; $btnCP.Location = [System.Drawing.Point]::new(280, 106)
            $btnCP.Size = [System.Drawing.Size]::new(90, 30); $btnCP.Font = $FontBold
            $btnCP.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnCP.FlatAppearance.BorderSize = 0
            $btnCP.BackColor = $clrRed; $btnCP.ForeColor = [System.Drawing.Color]::White
            $btnCP.DialogResult = [System.Windows.Forms.DialogResult]::OK; $dlgC.Controls.Add($btnCP)

            $btnCC = New-Object System.Windows.Forms.Button
            $btnCC.Text = 'Cancel'; $btnCC.Location = [System.Drawing.Point]::new(378, 106)
            $btnCC.Size = [System.Drawing.Size]::new(90, 30); $btnCC.Font = $FontBold
            $btnCC.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnCC.FlatAppearance.BorderSize = 0
            $btnCC.BackColor = [System.Drawing.Color]::FromArgb(225, 228, 238); $btnCC.ForeColor = $clrText
            $btnCC.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $dlgC.Controls.Add($btnCC)
            $dlgC.AcceptButton = $btnCP; $dlgC.CancelButton = $btnCC

            if ($dlgC.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK -or $tbC.Text -ne 'YES') {
                Write-Log 'Cancelled by user.' 'WARN'; return
            }
            Write-Log 'User typed YES — proceeding with live changes.'
        }

        $btnRun.Enabled = $false; $btnBrowse.Enabled = $false; $btnScan.Enabled = $false; $btnClose.Enabled = $false
        $progress.Value = 0
        Write-Log '=== Hide from Address Book started ==='
        Write-Log "Discovery folder : $discFolder"
        if ($whatIf) { Write-Log 'WhatIf mode — no changes will be made.' 'WARN' }

        $rs = [hashtable]::Synchronized(@{
            Done       = $false
            FatalError = $null
            Progress   = 0
            Total      = $totalItems
            LogQueue   = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())
        })

        $logFilePath    = $script:LogFile
        $selectedCsvArr = @($selectedCsvs)
        $totalItems2    = $totalItems
        $secDefsArr     = @($secDefsRun)

        $script:rdRunspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $script:rdRunspace.ApartmentState = 'MTA'
        $script:rdRunspace.ThreadOptions  = 'ReuseThread'
        $script:rdRunspace.Open()
        $script:rdPS = [System.Management.Automation.PowerShell]::Create()
        $script:rdPS.Runspace = $script:rdRunspace

        [void]$script:rdPS.AddScript({
            param($discoveryFolder, $selectedCsvs, $secDefs, $whatIf, $rs, $logFilePath)

            function QLog {
                param([string]$Msg, [string]$Level = 'INFO')
                $rs.LogQueue.Enqueue("[$Level] $Msg")
                try { "$(Get-Date -Format 'HH:mm:ss') [$Level] $Msg" | Add-Content $logFilePath -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
            }

            $exoConnected = $false

            function Ensure-ExoConnected {
                if ($exoConnected) { return $true }
                try {
                    QLog 'Connecting to Exchange Online — sign in when the browser opens...'
                    $cmds = @('Set-Mailbox','Set-MailContact','Set-DistributionGroup','Set-UnifiedGroup')
                    Connect-ExchangeOnline -ShowBanner:$false -CommandName $cmds -ErrorAction Stop
                    $exoConnected = $true
                    QLog 'Exchange Online connected.' 'OK'
                    return $true
                } catch {
                    QLog "Failed to connect to Exchange Online: $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'
                    return $false
                }
            }

            function Read-SectionCsv {
                param([string]$CsvName)
                $path = Join-Path $discoveryFolder $CsvName
                if (-not (Test-Path $path)) { QLog "CSV not found: $CsvName" 'WARN'; return $null }
                try {
                    $rows = @(Import-Csv -Path $path -Encoding UTF8)
                    QLog "Loaded $($rows.Count) row(s) from $CsvName"
                    return $rows
                } catch {
                    QLog "Failed to read ${CsvName}: $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'
                    return $null
                }
            }

            # Check if ExchangeOnlineManagement is available
            $mod = Get-Module -ListAvailable -Name 'ExchangeOnlineManagement' -ErrorAction SilentlyContinue
            if (-not $mod) {
                QLog 'ExchangeOnlineManagement module is not installed.' 'ERROR'
                QLog 'Install it with: Install-Module ExchangeOnlineManagement -Scope CurrentUser' 'ERROR'
                $rs.FatalError = 'ExchangeOnlineManagement module not installed'
                $rs.Done = $true
                return
            }

            # Import the module
            try {
                Import-Module 'ExchangeOnlineManagement' -ErrorAction Stop
                QLog 'ExchangeOnlineManagement module loaded successfully.' 'OK'
            } catch {
                QLog "Failed to import ExchangeOnlineManagement: $($_.Exception.Message)" 'ERROR'
                $rs.FatalError = "Module import failed: $($_.Exception.Message)"
                $rs.Done = $true
                return
            }

            # Build a lookup from CsvName → section def so we know the KeyField
            $secLookup = @{}
            foreach ($s in $secDefs) { $secLookup[$s.CsvName] = $s }

            try {
                foreach ($csvName in $selectedCsvs) {
                    QLog "--- $csvName ---"
                    $sec  = $secLookup[$csvName]
                    $rows = Read-SectionCsv $csvName
                    if (-not $rows) { continue }

                    $connected = $whatIf -or (Ensure-ExoConnected)
                    if (-not $connected) { $rs.Progress += $rows.Count; continue }

                    $ok = 0; $fail = 0

                    foreach ($r in $rows) {
                        $identity = $r.($sec.KeyField)

                        if ($whatIf) {
                            QLog "WhatIf: would hide $identity  [$($sec.Label)]" 'WARN'
                            $ok++
                        } else {
                            try {
                                switch ($csvName) {
                                    '02_Mailboxes.csv'          { Set-Mailbox          -Identity $identity -HiddenFromAddressListsEnabled $true -Confirm:$false -ErrorAction Stop }
                                    '03_DistributionGroups.csv' { Set-DistributionGroup -Identity $identity -HiddenFromAddressListsEnabled $true -Confirm:$false -ErrorAction Stop }
                                    '04_MailContacts.csv'       { Set-MailContact       -Identity $identity -HiddenFromAddressListsEnabled $true -Confirm:$false -ErrorAction Stop }
                                    '05_SharedMailboxes.csv'    { Set-Mailbox          -Identity $identity -HiddenFromAddressListsEnabled $true -Confirm:$false -ErrorAction Stop }
                                    '06_M365Groups.csv'         { Set-UnifiedGroup     -Identity $identity -HiddenFromAddressListsEnabled $true -Confirm:$false -ErrorAction Stop }
                                }
                                QLog "HIDDEN  $identity  [$($sec.Label)]" 'OK'
                                $ok++
                            } catch {
                                QLog "FAILED  ${identity}: $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'
                                $fail++
                            }
                        }
                        $rs.Progress++
                    }
                    QLog "$($sec.Label) — hidden: $ok  |  failed: $fail" $(if ($fail -eq 0) {'OK'} else {'WARN'})
                }

                QLog '=== All selected sections processed ===' 'OK'
            } catch {
                $rs.FatalError = $_.Exception.Message
                QLog "Fatal error: $($_.Exception.Message)" 'ERROR'
            } finally {
                $rs.Done = $true
            }
        })

        [void]$script:rdPS.AddParameters(@{
            discoveryFolder = $discFolder
            selectedCsvs    = $selectedCsvArr
            secDefs         = $secDefsArr
            whatIf          = $whatIf
            rs              = $rs
            logFilePath     = $logFilePath
        })

        $script:rdPS.BeginInvoke() | Out-Null

        $script:rdTimer = New-Object System.Windows.Forms.Timer
        $script:rdTimer.Interval = 300
        $rs2         = $rs
        $rdTimer2    = $script:rdTimer
        $rdRunspace2 = $script:rdRunspace

        $script:rdTimer.Add_Tick({
            try {
                while ($rs2.LogQueue.Count -gt 0) {
                    $raw   = $rs2.LogQueue.Dequeue()
                    $level = if ($raw -match '^\[(\w+)\]') { $matches[1] } else { 'INFO' }
                    $text  = $raw -replace '^\[\w+\] ', ''
                    Write-Log $text $level
                }
                if ($totalItems2 -gt 0 -and $rs2.Progress -gt 0) {
                    $pct = [int]([math]::Min(($rs2.Progress / $totalItems2) * 99, 99))
                    try { $progress.Value = $pct } catch { }
                }
                if ($rs2.Done) {
                    $rdTimer2.Stop()
                    try { $progress.Value = 100 } catch { }
                    Write-Log '=== Completed ===' $(if ($rs2.FatalError) {'ERROR'} else {'OK'})
                    try { $rdRunspace2.Close(); $rdRunspace2.Dispose() } catch {}
                }
            } catch {
                _RawLog "Tick error: $_"
            } finally {
                if ($rs2.Done) {
                    $btnRun.Enabled = $true; $btnBrowse.Enabled = $true
                    $btnScan.Enabled = $true; $btnClose.Enabled = $true
                }
            }
        }.GetNewClosure())
        $script:rdTimer.Start()
    }.GetNewClosure())

    # ── Form events ───────────────────────────────────────────────────────────
    $form.Add_FormClosing({
        param($s, $e)
        if ($script:rdTimer -and $script:rdTimer.Enabled) {
            $r = [System.Windows.Forms.MessageBox]::Show('Operation still running. Close anyway?','In Progress','YesNo','Warning')
            if ($r -eq [System.Windows.Forms.DialogResult]::No) { $e.Cancel = $true; return }
        }
        if ($script:rdTimer)    { try { $script:rdTimer.Stop();    $script:rdTimer.Dispose()    } catch {} }
        if ($script:rdPS)       { try { $script:rdPS.Stop();       $script:rdPS.Dispose()       } catch {} }
        if ($script:rdRunspace) { try { $script:rdRunspace.Close(); $script:rdRunspace.Dispose() } catch {} }
    }.GetNewClosure())

    $form.Add_Shown({
        $form.BringToFront(); $form.Activate()
        Write-Log '=== Hide Recipients from Address Book ready ==='
        Write-Log 'Browse to a Discovery folder, click Scan, select sections, then click Connect and Hide.'
    }.GetNewClosure())

    [System.Windows.Forms.Application]::Run($form)
}

try {
    Show-HideAddressBookUI
} catch {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    [System.Windows.Forms.MessageBox]::Show("Failed to start: $($_.Exception.Message)", 'Launch Error', 'OK', 'Error') | Out-Null
}
