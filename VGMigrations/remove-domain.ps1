#Requires -Version 7.0
<#
.SYNOPSIS
    remove-domain.ps1 — Remove M365 domain objects from Discovery CSVs (GUI).

.DESCRIPTION
    Reads Discovery CSV files produced by search-domain.ps1 and removes the
    corresponding M365 objects (mailboxes, distribution groups, mail contacts,
    shared mailboxes, M365 Groups/Teams, app registrations, enterprise apps,
    Entra devices, and proxy addresses) via Exchange Online and/or Microsoft Graph.

    The user scans a Discovery folder to list available sections, selects which to
    process, and must type YES to confirm before any live changes are made.
    A WhatIf mode lists what would be removed without making any changes.

.NOTES
    Dependency : lib.ps1 (colours, fonts, Write-Log), settings.ps1 (Show-SettingsDialog)
    Requires   : ExchangeOnlineManagement, Microsoft.Graph (various sub-modules)
    Log file   : logs\remove-domain-<timestamp>.log

    Change log
    ----------
    2026-05-27  Added settings.ps1 load + gear icon so the settings dialog is
                reachable from within this screen.
                Startup logging now captures lib/settings load results to file.
                Fallback Write-Log (used when lib.ps1 is unavailable) now also
                writes to the log file so all output is captured regardless of
                whether lib.ps1 loaded successfully.
                Error message from lib.ps1 load failure is now captured and logged.
#>

param(
    [string]$DiscoveryFolder = '',
    [string]$Sections        = 'all',  # comma-separated CSV names or 'all'
    [switch]$WhatIf
)

$script:RootDir = $PSScriptRoot

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ── File logging — initialised before lib load so any load error is captured ──
$_logDir = Join-Path $script:RootDir 'logs'
if (-not (Test-Path $_logDir)) { New-Item -ItemType Directory -Path $_logDir -Force | Out-Null }
$script:LogFile = Join-Path $_logDir "remove-domain-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

function _RawLog {
    param([string]$Msg)
    "$(Get-Date -Format 'HH:mm:ss.fff') $Msg" | Add-Content -Path $script:LogFile -Encoding UTF8 -ErrorAction SilentlyContinue
}

_RawLog "=== remove-domain.ps1 started  PID=$PID  PSVersion=$($PSVersionTable.PSVersion) ==="
_RawLog "PSScriptRoot : $script:RootDir"

$libPath      = Join-Path $script:RootDir 'lib.ps1'
$settingsPath = Join-Path $script:RootDir 'settings.ps1'
$_libLoaded   = $false

_RawLog "lib.ps1      : $libPath  exists=$(Test-Path $libPath)"

if (Test-Path $libPath) {
    try { . $libPath; $_libLoaded = $true; _RawLog "lib.ps1 loaded OK" }
    catch { _RawLog "lib.ps1 LOAD ERROR: $($_.Exception.Message)"; _RawLog "Stack: $($_.ScriptStackTrace)" }
} else {
    _RawLog "lib.ps1 NOT FOUND — colours and helpers will be missing"
}

_RawLog "settings.ps1 : $settingsPath  exists=$(Test-Path $settingsPath)"
if (Test-Path $settingsPath) {
    try { . $settingsPath; _RawLog "settings.ps1 loaded OK" }
    catch { _RawLog "settings.ps1 LOAD ERROR: $($_.Exception.Message)" }
} else {
    _RawLog "settings.ps1 NOT FOUND — settings dialog will be unavailable"
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
    $AnchorTL  = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
    $AnchorTR  = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
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
                'OK'    { [System.Drawing.Color]::FromArgb(65, 195, 110) }
                'WARN'  { [System.Drawing.Color]::FromArgb(220, 165, 45) }
                'ERROR' { [System.Drawing.Color]::FromArgb(225, 80, 80) }
                default { [System.Drawing.Color]::FromArgb(120, 155, 220) }
            }
            $script:rtbLog.SelectionColor = $lc; $script:rtbLog.AppendText("[$Level] ")
            $script:rtbLog.SelectionColor = [System.Drawing.Color]::FromArgb(205, 212, 230)
            $script:rtbLog.AppendText("$Msg`n"); $script:rtbLog.ScrollToCaret()
        }
    }
}

# Section definitions
$script:SectionDefs = @(
    [pscustomobject]@{ CsvName='01_AcceptedDomains.csv';    Label='Accepted Domains';       NeedsEXO=$true;  NeedsGraph=$false }
    [pscustomobject]@{ CsvName='02_Mailboxes.csv';          Label='Mailboxes';              NeedsEXO=$true;  NeedsGraph=$false }
    [pscustomobject]@{ CsvName='03_DistributionGroups.csv'; Label='Distribution Groups';    NeedsEXO=$true;  NeedsGraph=$false }
    [pscustomobject]@{ CsvName='04_MailContacts.csv';       Label='Mail Contacts';          NeedsEXO=$true;  NeedsGraph=$false }
    [pscustomobject]@{ CsvName='05_SharedMailboxes.csv';    Label='Shared Mailboxes';       NeedsEXO=$true;  NeedsGraph=$false }
    [pscustomobject]@{ CsvName='06_M365Groups.csv';         Label='M365 Groups (+ Teams)';  NeedsEXO=$true;  NeedsGraph=$false }
    [pscustomobject]@{ CsvName='07_AppRegistrations.csv';   Label='App Registrations';      NeedsEXO=$false; NeedsGraph=$true  }
    [pscustomobject]@{ CsvName='08_EnterpriseApps.csv';     Label='Enterprise Apps';        NeedsEXO=$false; NeedsGraph=$true  }
    [pscustomobject]@{ CsvName='11_Devices.csv';            Label='Devices';                NeedsEXO=$false; NeedsGraph=$true  }
    [pscustomobject]@{ CsvName='12_ProxyAddresses.csv';     Label='Proxy Addresses';        NeedsEXO=$true;  NeedsGraph=$false }
)

function Show-RemoveDomainUI {
    $rootDir     = $script:RootDir
    $secDefs     = $script:SectionDefs

    # ── Form ──────────────────────────────────────────────────────────────────
    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'Remove M365 Domain Objects'
    $form.ClientSize      = [System.Drawing.Size]::new(680, 820)
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
    $hdrLbl.Text = '  Remove M365 Domain Objects'; $hdrLbl.Font = $FontTitle
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
    $card.Size      = [System.Drawing.Size]::new(656, 456)
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

    # Separator
    $sep1 = New-Object System.Windows.Forms.Label
    $sep1.Location = [System.Drawing.Point]::new($lx, $y); $sep1.Size = [System.Drawing.Size]::new(628, 1)
    $sep1.BackColor = $clrBorder; $card.Controls.Add($sep1); $y += 10

    # Sections header
    $lblSecCap = New-Object System.Windows.Forms.Label
    $lblSecCap.Text = 'SECTIONS TO PROCESS'; $lblSecCap.Font = $FontCap; $lblSecCap.ForeColor = $clrMuted
    $lblSecCap.Location = [System.Drawing.Point]::new($lx, $y); $lblSecCap.AutoSize = $true
    $card.Controls.Add($lblSecCap); $y += 18

    # CheckedListBox
    $clbSections = New-Object System.Windows.Forms.CheckedListBox
    $clbSections.Location     = [System.Drawing.Point]::new($lx, $y)
    $clbSections.Size         = [System.Drawing.Size]::new(628, 210)
    $clbSections.Font         = $FontBody
    $clbSections.BackColor    = [System.Drawing.Color]::FromArgb(250, 251, 253)
    $clbSections.ForeColor    = $clrText
    $clbSections.BorderStyle  = [System.Windows.Forms.BorderStyle]::FixedSingle
    $clbSections.CheckOnClick = $true
    $card.Controls.Add($clbSections)
    foreach ($sec in $secDefs) {
        [void]$clbSections.Items.Add(("{0,-38} (not scanned)" -f $sec.Label), $false)
    }
    $y += 215

    # Select All / Deselect All
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

    # Separator
    $sep2 = New-Object System.Windows.Forms.Label
    $sep2.Location = [System.Drawing.Point]::new($lx, $y); $sep2.Size = [System.Drawing.Size]::new(628, 1)
    $sep2.BackColor = $clrBorder; $card.Controls.Add($sep2); $y += 10

    # On-Premises Active Directory
    $lblADCap = New-Object System.Windows.Forms.Label
    $lblADCap.Text = 'ON-PREMISES ACTIVE DIRECTORY'; $lblADCap.Font = $FontCap; $lblADCap.ForeColor = $clrMuted
    $lblADCap.Location = [System.Drawing.Point]::new($lx, $y); $lblADCap.AutoSize = $true
    $card.Controls.Add($lblADCap); $y += 18

    $chkUpdateAD = New-Object System.Windows.Forms.CheckBox
    $chkUpdateAD.Text = 'Update on-prem AD users (UPN, proxyAddresses, mail → ourvolaris.onmicrosoft.com)'
    $chkUpdateAD.Location = [System.Drawing.Point]::new($lx, $y)
    $chkUpdateAD.Size = [System.Drawing.Size]::new(628, 20)
    $chkUpdateAD.ForeColor = $clrText
    $card.Controls.Add($chkUpdateAD); $y += 28

    # Separator
    $sep3 = New-Object System.Windows.Forms.Label
    $sep3.Location = [System.Drawing.Point]::new($lx, $y); $sep3.Size = [System.Drawing.Size]::new(628, 1)
    $sep3.BackColor = $clrBorder; $card.Controls.Add($sep3); $y += 12

    # WhatIf + Run
    $chkWhatIf = New-Object System.Windows.Forms.CheckBox
    $chkWhatIf.Text = 'WhatIf  (list items only - no changes will be made)'
    $chkWhatIf.Location = [System.Drawing.Point]::new($lx, $y + 6)
    $chkWhatIf.AutoSize = $true; $chkWhatIf.ForeColor = $clrText
    $card.Controls.Add($chkWhatIf)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = 'Connect and Remove'; $btnRun.Location = [System.Drawing.Point]::new(470, $y)
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
    $script:tickBusy   = $false
    $script:RowCounts  = @(0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
    $RowCounts = $script:RowCounts  # local ref so closures don't use $script: scope

    # ── Selection status helper ───────────────────────────────────────────────
    $updateSelStatus = {
        $cnt = $clbSections.CheckedItems.Count
        $selItems = 0
        for ($i = 0; $i -lt $clbSections.Items.Count; $i++) {
            if ($clbSections.GetItemChecked($i)) { $selItems += $RowCounts[$i] }
        }
        $lblSelStatus.Text = "$cnt section(s) selected  ($selItems items)"
    }.GetNewClosure()

    # ── Browse ────────────────────────────────────────────────────────────────
    $btnBrowse.Add_Click({
        Write-Log "Browse folder dialog opened"
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description = 'Select the Discovery folder (or its parent output folder)'
        $fbd.ShowNewFolderButton = $false
        if ($tbFolder.Text -and (Test-Path $tbFolder.Text)) { $fbd.SelectedPath = $tbFolder.Text }
        if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Write-Log "Folder selected: $($fbd.SelectedPath)"
            $tbFolder.Text = $fbd.SelectedPath
        } else {
            Write-Log "Browse folder dialog cancelled"
        }
    }.GetNewClosure())

    # ── Scan ──────────────────────────────────────────────────────────────────
    $secDefsLocal = $secDefs
    $btnScan.Add_Click({
        $raw = $tbFolder.Text.Trim().Trim('"')
        Write-Log "Scan clicked — folder='$raw'"
        if (-not $raw) {
            Write-Log 'Scan validation failed: no folder entered' 'WARN'
            [System.Windows.Forms.MessageBox]::Show('Please enter or browse to a Discovery folder.', 'Missing Folder', 'OK', 'Warning') | Out-Null; return
        }
        if (-not (Test-Path $raw)) {
            Write-Log "Scan validation failed: path not found '$raw'" 'WARN'
            [System.Windows.Forms.MessageBox]::Show("Path not found:`n$raw", 'Not Found', 'OK', 'Warning') | Out-Null; return
        }
        $discFolder = $raw
        $candidate  = Join-Path $raw 'Discovery'
        if ((Split-Path $raw -Leaf) -ne 'Discovery' -and (Test-Path $candidate)) {
            $discFolder = $candidate; $tbFolder.Text = $discFolder
            Write-Log "Auto-resolved to Discovery subfolder: $discFolder"
        }
        Write-Log "Scanning folder: $discFolder"
        $lblScanStatus.Text = 'Scanning...'
        [System.Windows.Forms.Application]::DoEvents()
        $found = 0
        for ($i = 0; $i -lt $secDefsLocal.Count; $i++) {
            $sec  = $secDefsLocal[$i]
            $path = Join-Path $discFolder $sec.CsvName
            $cnt  = 0
            if (Test-Path $path) {
                try { $cnt = @(Import-Csv -Path $path -Encoding UTF8).Count }
                catch { Write-Log "Could not read $($sec.CsvName): $_" 'WARN'; $cnt = 0 }
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
        $lblScanStatus.Text = "Scan complete - $found of $($secDefsLocal.Count) CSV(s) found"
        & $updateSelStatus
        Write-Log "Scan complete: $found of $($secDefsLocal.Count) CSV(s) found in $discFolder"
    }.GetNewClosure())

    # ── Select All / Deselect All ─────────────────────────────────────────────
    $btnSelAll.Add_Click({
        Write-Log "Select All clicked"
        for ($i = 0; $i -lt $clbSections.Items.Count; $i++) { $clbSections.SetItemChecked($i, $true) }
        & $updateSelStatus
    }.GetNewClosure())
    $btnDeselAll.Add_Click({
        Write-Log "Deselect All clicked"
        for ($i = 0; $i -lt $clbSections.Items.Count; $i++) { $clbSections.SetItemChecked($i, $false) }
        & $updateSelStatus
    }.GetNewClosure())
    $clbSections.Add_ItemCheck({ & $updateSelStatus }.GetNewClosure())

    # ── Run ───────────────────────────────────────────────────────────────────
    $secDefsRun = $secDefs
    $btnRun.Add_Click({
        $raw = $tbFolder.Text.Trim().Trim('"')
        Write-Log "Run clicked — folder='$raw'  WhatIf=$($chkWhatIf.Checked)"
        if (-not $raw -or -not (Test-Path $raw)) {
            Write-Log 'Validation failed: folder not set or not found' 'WARN'
            [System.Windows.Forms.MessageBox]::Show('Please scan a valid Discovery folder first.', 'No Folder', 'OK', 'Warning') | Out-Null; return
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
            Write-Log 'Validation failed: no sections selected' 'WARN'
            [System.Windows.Forms.MessageBox]::Show('No sections selected. Scan first and tick at least one section.', 'Nothing Selected', 'OK', 'Warning') | Out-Null; return
        }
        Write-Log "$($selectedCsvs.Count) section(s) selected  ($totalItems total item(s))  folder=$discFolder"

        $whatIf = $chkWhatIf.Checked

        if (-not $whatIf) {
            # Confirmation dialog
            $dlgC = New-Object System.Windows.Forms.Form
            $dlgC.Text = 'Confirm Removal'; $dlgC.ClientSize = [System.Drawing.Size]::new(480, 170)
            $dlgC.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
            $dlgC.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
            $dlgC.MaximizeBox = $false; $dlgC.MinimizeBox = $false; $dlgC.BackColor = $clrBg

            $lblCMsg = New-Object System.Windows.Forms.Label
            $lblCMsg.Text = "Permanently remove $totalItems item(s) across $($selectedCsvs.Count) section(s)?`nThis CANNOT be undone. Type YES to confirm:"
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

            Write-Log "Showing removal confirmation — $totalItems item(s) across $($selectedCsvs.Count) section(s)"
            $dlgResult = $dlgC.ShowDialog()
            if ($dlgResult -ne [System.Windows.Forms.DialogResult]::OK -or $tbC.Text -ne 'YES') {
                Write-Log 'Removal cancelled by user at confirmation prompt' 'WARN'; return
            }
            Write-Log 'User typed YES — proceeding with live removal'
        }

        $btnRun.Enabled = $false; $btnBrowse.Enabled = $false; $btnScan.Enabled = $false; $btnClose.Enabled = $false
        $progress.Value = 0
        _RawLog "Buttons locked — starting runspace"
        Write-Log '=== Remove M365 Domain Objects started ==='
        Write-Log "Discovery folder : $discFolder"
        if ($whatIf) { Write-Log 'WhatIf mode - no changes will be made' 'WARN' }

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

        $script:rdRunspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $script:rdRunspace.ApartmentState = 'MTA'
        $script:rdRunspace.ThreadOptions  = 'ReuseThread'
        $script:rdRunspace.Open()
        $script:rdPS = [System.Management.Automation.PowerShell]::Create()
        $script:rdPS.Runspace = $script:rdRunspace

        [void]$script:rdPS.AddScript({
            param($discoveryFolder, $selectedCsvs, $whatIf, $updateAD, $rs, $logFilePath)

            function QLog {
                param([string]$Msg, [string]$Level = 'INFO')
                $rs.LogQueue.Enqueue("[$Level] $Msg")
                try { "$(Get-Date -Format 'HH:mm:ss') [$Level] $Msg" | Add-Content $logFilePath -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
            }

            $connState = @{ EXO = $false; Graph = $false }

            function Ensure-ExoConnected {
                if ($connState.EXO) { return $true }
                try {
                    QLog 'Connecting to Exchange Online - sign in when the browser opens...'
                    $cmds = @('Remove-AcceptedDomain','Remove-DistributionGroup','Remove-MailContact',
                              'Remove-Mailbox','Remove-UnifiedGroup','Set-Mailbox','Set-DistributionGroup',
                              'Set-UnifiedGroup','Get-Mailbox','Get-DistributionGroup','Get-UnifiedGroup',
                              'Get-MailContact','Get-Recipient')
                    Connect-ExchangeOnline -ShowBanner:$false -CommandName $cmds -ErrorAction Stop
                    $connState.EXO = $true
                    QLog 'Exchange Online connected.' 'OK'
                    return $true
                } catch {
                    QLog "Failed to connect to Exchange Online: $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'
                    return $false
                }
            }

            function Ensure-GraphConnected {
                if ($connState.Graph) { return $true }
                try {
                    QLog 'Connecting to Microsoft Graph - sign in when the browser opens...'
                    $scopes = @('Application.ReadWrite.All','Device.ReadWrite.All',
                                'Directory.ReadWrite.All','DeviceManagementManagedDevices.ReadWrite.All')
                    try { Set-MgGraphOption -DisableLoginByWAM $true -ErrorAction SilentlyContinue } catch {}
                    Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
                    $connState.Graph = $true
                    QLog 'Microsoft Graph connected.' 'OK'
                    return $true
                } catch {
                    QLog "Failed to connect to Microsoft Graph: $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'
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

            # Module check
            $exoCsvs   = @('01_AcceptedDomains.csv','02_Mailboxes.csv','03_DistributionGroups.csv',
                           '04_MailContacts.csv','05_SharedMailboxes.csv','06_M365Groups.csv','12_ProxyAddresses.csv')
            $graphCsvs = @('07_AppRegistrations.csv','08_EnterpriseApps.csv','11_Devices.csv')
            $needsEXO   = $selectedCsvs | Where-Object { $exoCsvs -contains $_ }
            $needsGraph = $selectedCsvs | Where-Object { $graphCsvs -contains $_ }

            if ($needsEXO) {
                $mod = Get-Module -ListAvailable -Name 'ExchangeOnlineManagement' -ErrorAction SilentlyContinue
                if (-not $mod) { QLog 'ExchangeOnlineManagement not installed. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser' 'WARN' }
                else { Import-Module 'ExchangeOnlineManagement' -ErrorAction SilentlyContinue }
            }
            if ($needsGraph) {
                foreach ($m in @('Microsoft.Graph.Authentication','Microsoft.Graph.Identity.DirectoryManagement','Microsoft.Graph.Applications')) {
                    $mod = Get-Module -ListAvailable -Name $m -ErrorAction SilentlyContinue
                    if (-not $mod) { QLog "Module not installed: $m" 'WARN' }
                    else { Import-Module $m -ErrorAction SilentlyContinue }
                }
            }

            try {
                foreach ($csvName in $selectedCsvs) {
                    QLog "--- $csvName ---"

                    if ($csvName -eq '01_AcceptedDomains.csv') {
                        $rows = Read-SectionCsv $csvName
                        if ($rows) {
                            $removable = @($rows | Where-Object { $_.IsDefault -ne 'True' })
                            $skipped   = $rows.Count - $removable.Count
                            if ($skipped -gt 0) { QLog "Skipping $skipped default domain(s)." 'WARN' }
                            if ($removable.Count -gt 0) {
                                $ok = 0; $fail = 0
                                $connected = $whatIf -or (Ensure-ExoConnected)
                                if ($connected) {
                                    foreach ($r in $removable) {
                                        $name = $r.DomainName
                                        if ($whatIf) { QLog "WhatIf: would remove accepted domain: $name" 'WARN'; $ok++ }
                                        else {
                                            try { Remove-AcceptedDomain -Identity $name -Confirm:$false -ErrorAction Stop; QLog "REMOVED accepted domain: $name" 'OK'; $ok++ }
                                            catch { QLog "Failed '${name}': $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'; $fail++ }
                                        }
                                        $rs.Progress++
                                    }
                                } else { $rs.Progress += $removable.Count }
                            } else { QLog 'No non-default accepted domains to remove.' 'WARN' }
                            $rs.Progress += $skipped
                            QLog "Accepted domains: done"
                        }
                    }

                    elseif ($csvName -eq '02_Mailboxes.csv') {
                        $rows = Read-SectionCsv $csvName
                        if ($rows) {
                            $ok = 0; $fail = 0
                            $connected = $whatIf -or (Ensure-ExoConnected)
                            if ($connected) {
                                foreach ($r in $rows) {
                                    $upn = $r.UserPrincipalName
                                    if ($whatIf) { QLog "WhatIf: would remove mailbox: $upn" 'WARN'; $ok++ }
                                    else {
                                        try { Remove-Mailbox -Identity $upn -Confirm:$false -ErrorAction Stop; QLog "REMOVED mailbox: $upn" 'OK'; $ok++ }
                                        catch { QLog "Failed '${upn}': $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'; $fail++ }
                                    }
                                    $rs.Progress++
                                }
                            } else { $rs.Progress += $rows.Count }
                            QLog "Mailboxes: removed $ok  failed $fail"
                        }
                    }

                    elseif ($csvName -eq '03_DistributionGroups.csv') {
                        $rows = Read-SectionCsv $csvName
                        if ($rows) {
                            $ok = 0; $fail = 0
                            $connected = $whatIf -or (Ensure-ExoConnected)
                            if ($connected) {
                                foreach ($r in $rows) {
                                    $addr = $r.PrimarySmtpAddress
                                    if ($whatIf) { QLog "WhatIf: would remove DG: $addr  [$($r.DisplayName)]" 'WARN'; $ok++ }
                                    else {
                                        try { Remove-DistributionGroup -Identity $addr -Confirm:$false -ErrorAction Stop; QLog "REMOVED DG: $addr  [$($r.DisplayName)]" 'OK'; $ok++ }
                                        catch { QLog "Failed '${addr}': $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'; $fail++ }
                                    }
                                    $rs.Progress++
                                }
                            } else { $rs.Progress += $rows.Count }
                            QLog "Distribution groups: removed $ok  failed $fail"
                        }
                    }

                    elseif ($csvName -eq '04_MailContacts.csv') {
                        $rows = Read-SectionCsv $csvName
                        if ($rows) {
                            $ok = 0; $fail = 0
                            $connected = $whatIf -or (Ensure-ExoConnected)
                            if ($connected) {
                                foreach ($r in $rows) {
                                    $addr = $r.ExternalEmailAddress; $dn = $r.DisplayName
                                    if ($whatIf) { QLog "WhatIf: would remove mail contact: $addr  [$dn]" 'WARN'; $ok++ }
                                    else {
                                        try { Remove-MailContact -Identity $addr -Confirm:$false -ErrorAction Stop; QLog "REMOVED mail contact: $addr  [$dn]" 'OK'; $ok++ }
                                        catch { QLog "Failed '${addr}': $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'; $fail++ }
                                    }
                                    $rs.Progress++
                                }
                            } else { $rs.Progress += $rows.Count }
                            QLog "Mail contacts: removed $ok  failed $fail"
                        }
                    }

                    elseif ($csvName -eq '05_SharedMailboxes.csv') {
                        $rows = Read-SectionCsv $csvName
                        if ($rows) {
                            $ok = 0; $fail = 0
                            $connected = $whatIf -or (Ensure-ExoConnected)
                            if ($connected) {
                                foreach ($r in $rows) {
                                    $addr = $r.PrimarySmtpAddress
                                    if ($whatIf) { QLog "WhatIf: would remove shared mailbox: $addr  [$($r.DisplayName)]" 'WARN'; $ok++ }
                                    else {
                                        try { Remove-Mailbox -Identity $addr -Confirm:$false -ErrorAction Stop; QLog "REMOVED shared mailbox: $addr  [$($r.DisplayName)]" 'OK'; $ok++ }
                                        catch { QLog "Failed '${addr}': $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'; $fail++ }
                                    }
                                    $rs.Progress++
                                }
                            } else { $rs.Progress += $rows.Count }
                            QLog "Shared mailboxes: removed $ok  failed $fail"
                        }
                    }

                    elseif ($csvName -eq '06_M365Groups.csv') {
                        $rows = Read-SectionCsv $csvName
                        if ($rows) {
                            $ok = 0; $fail = 0
                            $connected = $whatIf -or (Ensure-ExoConnected)
                            if ($connected) {
                                foreach ($r in $rows) {
                                    $addr = $r.PrimarySmtpAddress
                                    if ($whatIf) { QLog "WhatIf: would remove M365 Group: $addr  [$($r.DisplayName)]" 'WARN'; $ok++ }
                                    else {
                                        try { Remove-UnifiedGroup -Identity $addr -Confirm:$false -ErrorAction Stop; QLog "REMOVED M365 Group: $addr  [$($r.DisplayName)]" 'OK'; $ok++ }
                                        catch { QLog "Failed '${addr}': $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'; $fail++ }
                                    }
                                    $rs.Progress++
                                }
                            } else { $rs.Progress += $rows.Count }
                            QLog "M365 Groups: removed $ok  failed $fail"
                        }
                    }

                    elseif ($csvName -eq '07_AppRegistrations.csv') {
                        $rows = Read-SectionCsv $csvName
                        if ($rows) {
                            $ok = 0; $fail = 0
                            $connected = $whatIf -or (Ensure-GraphConnected)
                            if ($connected) {
                                foreach ($r in $rows) {
                                    $appId = $r.AppId; $dn = $r.DisplayName
                                    if ($whatIf) { QLog "WhatIf: would remove app registration: $dn  [AppId: $appId]" 'WARN'; $ok++ }
                                    else {
                                        try { Remove-MgApplication -ApplicationId $appId -Confirm:$false -ErrorAction Stop; QLog "REMOVED app registration: $dn  [AppId: $appId]" 'OK'; $ok++ }
                                        catch { QLog "Failed '${dn}' (${appId}): $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'; $fail++ }
                                    }
                                    $rs.Progress++
                                }
                            } else { $rs.Progress += $rows.Count }
                            QLog "App registrations: removed $ok  failed $fail"
                        }
                    }

                    elseif ($csvName -eq '08_EnterpriseApps.csv') {
                        $rows = Read-SectionCsv $csvName
                        if ($rows) {
                            $ok = 0; $fail = 0
                            $connected = $whatIf -or (Ensure-GraphConnected)
                            if ($connected) {
                                foreach ($r in $rows) {
                                    $spId = $r.ObjectId; $dn = $r.DisplayName
                                    if ($whatIf) { QLog "WhatIf: would remove enterprise app: $dn  [ObjectId: $spId]" 'WARN'; $ok++ }
                                    else {
                                        try { Remove-MgServicePrincipal -ServicePrincipalId $spId -Confirm:$false -ErrorAction Stop; QLog "REMOVED enterprise app: $dn  [ObjectId: $spId]" 'OK'; $ok++ }
                                        catch { QLog "Failed '${dn}' (${spId}): $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'; $fail++ }
                                    }
                                    $rs.Progress++
                                }
                            } else { $rs.Progress += $rows.Count }
                            QLog "Enterprise apps: removed $ok  failed $fail"
                        }
                    }

                    elseif ($csvName -eq '11_Devices.csv') {
                        $rows = Read-SectionCsv $csvName
                        if ($rows) {
                            $ok = 0; $fail = 0
                            $connected = $whatIf -or (Ensure-GraphConnected)
                            if ($connected) {
                                foreach ($r in $rows) {
                                    $objId = $r.DeviceObjectId; $name = $r.DeviceName; $owner = $r.OwnerUPN
                                    if ($whatIf) { QLog "WhatIf: would remove device: $name  [ObjectId: $objId  Owner: $owner]" 'WARN'; $ok++ }
                                    else {
                                        try { Remove-MgDevice -DeviceId $objId -Confirm:$false -ErrorAction Stop; QLog "REMOVED device: $name  [ObjectId: $objId  Owner: $owner]" 'OK'; $ok++ }
                                        catch { QLog "Failed '${name}' (${objId}): $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'; $fail++ }
                                    }
                                    $rs.Progress++
                                }
                            } else { $rs.Progress += $rows.Count }
                            QLog "Devices: removed $ok  failed $fail"
                        }
                    }

                    elseif ($csvName -eq '12_ProxyAddresses.csv') {
                        $rows = Read-SectionCsv $csvName
                        if ($rows) {
                            $primaries    = @($rows | Where-Object { $_.IsPrimary -eq 'True' })
                            $nonPrimaries = @($rows | Where-Object { $_.IsPrimary -ne 'True' })
                            if ($primaries.Count -gt 0) {
                                QLog "$($primaries.Count) primary SMTP address(es) will be SKIPPED - reclassify manually." 'WARN'
                            }
                            if ($nonPrimaries.Count -gt 0) {
                                $ok = 0; $fail = 0; $skip = 0
                                $connected = $whatIf -or (Ensure-ExoConnected)
                                if ($connected) {
                                    $byRecipient = $nonPrimaries | Group-Object -Property PrimarySmtpAddress
                                    foreach ($grp in $byRecipient) {
                                        $primaryAddr     = $grp.Name
                                        $addressesToDrop = @($grp.Group | ForEach-Object { "$($_.AddressType):$($_.ProxyAddress)" })
                                        if ($whatIf) {
                                            foreach ($a in $addressesToDrop) { QLog "WhatIf: would remove proxy $a from $primaryAddr" 'WARN' }
                                            $ok += $addressesToDrop.Count; $rs.Progress += $addressesToDrop.Count
                                        } else {
                                            try {
                                                $recip = Get-Recipient -Identity $primaryAddr -ErrorAction Stop
                                                $type  = $recip.RecipientTypeDetails
                                                $currentProxies = @($recip.EmailAddresses | ForEach-Object { $_.ToString() })
                                                $newProxies = @($currentProxies | Where-Object {
                                                    $a = $_; -not ($addressesToDrop | Where-Object { $_ -ieq $a })
                                                })
                                                $removed = $currentProxies.Count - $newProxies.Count
                                                if ($removed -eq 0) {
                                                    QLog "No matching addresses found on ${primaryAddr} - skipping." 'WARN'
                                                    $skip++; $rs.Progress += $addressesToDrop.Count
                                                } else {
                                                    switch -Wildcard ($type) {
                                                        'UserMailbox'        { Set-Mailbox           -Identity $primaryAddr -EmailAddresses $newProxies -ErrorAction Stop }
                                                        'SharedMailbox'      { Set-Mailbox           -Identity $primaryAddr -EmailAddresses $newProxies -ErrorAction Stop }
                                                        'RoomMailbox'        { Set-Mailbox           -Identity $primaryAddr -EmailAddresses $newProxies -ErrorAction Stop }
                                                        'EquipmentMailbox'   { Set-Mailbox           -Identity $primaryAddr -EmailAddresses $newProxies -ErrorAction Stop }
                                                        'MailUniversalDistributionGroup' { Set-DistributionGroup -Identity $primaryAddr -EmailAddresses $newProxies -ErrorAction Stop }
                                                        'GroupMailbox'       { Set-UnifiedGroup      -Identity $primaryAddr -EmailAddresses $newProxies -ErrorAction Stop }
                                                        default {
                                                            QLog "Unknown recipient type '${type}' for ${primaryAddr} - skipping." 'WARN'
                                                            $skip++; $rs.Progress += $addressesToDrop.Count
                                                        }
                                                    }
                                                    foreach ($dropped in $addressesToDrop) { QLog "REMOVED proxy: $dropped  from: $primaryAddr" 'OK' }
                                                    $ok += $removed; $rs.Progress += $addressesToDrop.Count
                                                }
                                            } catch {
                                                QLog "Failed to update proxies on '${primaryAddr}': $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'
                                                $fail++; $rs.Progress += $addressesToDrop.Count
                                            }
                                        }
                                    }
                                } else { $rs.Progress += $nonPrimaries.Count }
                                QLog "Proxy addresses: removed $ok  failed $fail  skipped $skip"
                            } else {
                                QLog 'No non-primary proxy addresses to remove.' 'WARN'
                            }
                            $rs.Progress += $primaries.Count
                        }
                    }
                }

                # ── On-Premises Active Directory Update ──────────────────────────────
                if ($updateAD) {
                    QLog '--- On-Premises Active Directory ---'

                    # Import ActiveDirectory module
                    try {
                        Import-Module ActiveDirectory -ErrorAction Stop
                        QLog 'ActiveDirectory module loaded.' 'OK'
                    } catch {
                        QLog "ActiveDirectory module not available: $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'
                        QLog 'Skipping on-prem AD updates.' 'WARN'
                        $updateAD = $false
                    }

                    if ($updateAD) {
                        # Read 02_Mailboxes.csv to get the list of users with the domain to remove
                        $mailboxRows = Read-SectionCsv '02_Mailboxes.csv'
                        if ($mailboxRows) {
                            QLog "Processing $($mailboxRows.Count) user(s) for on-prem AD updates..."

                            # Extract the domain to remove from the first user's UPN or primary SMTP
                            $domainToRemove = $null
                            foreach ($row in $mailboxRows) {
                                if ($row.UserPrincipalName -and $row.UserPrincipalName -match '@(.+)$') {
                                    $domainToRemove = $matches[1]
                                    break
                                }
                            }

                            if (-not $domainToRemove) {
                                QLog 'Could not determine domain to remove from mailbox data.' 'WARN'
                            } else {
                                QLog "Domain to replace: $domainToRemove  →  ourvolaris.onmicrosoft.com"

                                $ok = 0; $fail = 0; $skip = 0

                                foreach ($row in $mailboxRows) {
                                    $upn = $row.UserPrincipalName
                                    if (-not $upn) { $skip++; continue }

                                    try {
                                        # Find AD user by UPN
                                        $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$upn'" -Properties mail,proxyAddresses -ErrorAction Stop

                                        if (-not $adUser) {
                                            QLog "AD user not found: $upn" 'WARN'
                                            $skip++
                                            continue
                                        }

                                        $changes = @{}

                                        # 1. Update UPN if it contains the domain to remove
                                        if ($adUser.UserPrincipalName -like "*@$domainToRemove") {
                                            $newUPN = $adUser.UserPrincipalName -replace [regex]::Escape("@$domainToRemove"), '@ourvolaris.onmicrosoft.com'
                                            $changes['UserPrincipalName'] = $newUPN
                                        }

                                        # 2. Update mail attribute if it contains the domain to remove
                                        if ($adUser.mail -and $adUser.mail -like "*@$domainToRemove") {
                                            $newMail = $adUser.mail -replace [regex]::Escape("@$domainToRemove"), '@ourvolaris.onmicrosoft.com'
                                            $changes['mail'] = $newMail
                                        }

                                        # 3. Update proxyAddresses - replace all references to the domain
                                        if ($adUser.proxyAddresses) {
                                            $newProxies = @()
                                            $proxyChanged = $false

                                            foreach ($proxy in $adUser.proxyAddresses) {
                                                if ($proxy -match "^(smtp|SMTP):(.+)@$([regex]::Escape($domainToRemove))$") {
                                                    $prefix = $matches[1]
                                                    $localPart = $matches[2]
                                                    $newProxies += "${prefix}:${localPart}@ourvolaris.onmicrosoft.com"
                                                    $proxyChanged = $true
                                                } else {
                                                    $newProxies += $proxy
                                                }
                                            }

                                            if ($proxyChanged) {
                                                $changes['proxyAddresses'] = $newProxies
                                            }
                                        }

                                        if ($changes.Count -eq 0) {
                                            QLog "No AD changes needed for $upn" 'WARN'
                                            $skip++
                                        } else {
                                            if ($whatIf) {
                                                QLog "WhatIf: would update AD user $upn with:" 'WARN'
                                                foreach ($k in $changes.Keys) {
                                                    if ($k -eq 'proxyAddresses') {
                                                        QLog "  proxyAddresses: $($changes[$k].Count) address(es)" 'WARN'
                                                    } else {
                                                        QLog "  ${k}: $($changes[$k])" 'WARN'
                                                    }
                                                }
                                                $ok++
                                            } else {
                                                Set-ADUser -Identity $adUser.DistinguishedName @changes -ErrorAction Stop
                                                QLog "UPDATED AD user: $upn" 'OK'
                                                foreach ($k in $changes.Keys) {
                                                    if ($k -eq 'proxyAddresses') {
                                                        QLog "  proxyAddresses: $($changes[$k].Count) address(es) updated" 'OK'
                                                    } else {
                                                        QLog "  ${k} → $($changes[$k])" 'OK'
                                                    }
                                                }
                                                $ok++
                                            }
                                        }
                                    } catch {
                                        QLog "Failed to update AD user ${upn}: $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'
                                        $fail++
                                    }
                                }

                                QLog "On-prem AD updates: $ok updated  |  $fail failed  |  $skip skipped" $(if ($fail -eq 0) {'OK'} else {'WARN'})
                            }
                        } else {
                            QLog '02_Mailboxes.csv not found - skipping AD updates.' 'WARN'
                        }
                    }
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
            whatIf          = $whatIf
            updateAD        = $chkUpdateAD.Checked
            rs              = $rs
            logFilePath     = $logFilePath
        })

        $script:rdPS.BeginInvoke() | Out-Null

        $script:rdTimer = New-Object System.Windows.Forms.Timer
        $script:rdTimer.Interval = 300
        $rs2         = $rs
        $rdTimer2    = $script:rdTimer    # local ref — $script: doesn't resolve reliably inside closures
        $rdRunspace2 = $script:rdRunspace # local ref for same reason
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
                    $progress.Value = $pct
                }
                if ($rs2.Done) {
                    $rdTimer2.Stop()
                    $progress.Value = 100
                    $status = if ($rs2.FatalError) { 'ERROR' } else { 'OK' }
                    Write-Log '=== Completed ===' $status
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
        _RawLog "FormClosing event  timerRunning=$($script:rdTimer -and $script:rdTimer.Enabled)"
        if ($script:rdTimer -and $script:rdTimer.Enabled) {
            $r = [System.Windows.Forms.MessageBox]::Show('Removal is still running. Close anyway?', 'In Progress', 'YesNo', 'Warning')
            if ($r -eq [System.Windows.Forms.DialogResult]::No) {
                _RawLog "FormClosing cancelled — user chose to keep running"
                $e.Cancel = $true; return
            }
            _RawLog "FormClosing confirmed by user despite running timer"
        }
        _RawLog "FormClosing: cleaning up timer/PS/runspace"
        if ($script:rdTimer)    { try { $script:rdTimer.Stop(); $script:rdTimer.Dispose() }       catch { _RawLog "Timer dispose error: $_" } }
        if ($script:rdPS)       { try { $script:rdPS.Stop();    $script:rdPS.Dispose() }           catch { _RawLog "PS dispose error: $_" } }
        if ($script:rdRunspace) { try { $script:rdRunspace.Close(); $script:rdRunspace.Dispose() } catch { _RawLog "Runspace dispose error: $_" } }
        _RawLog "FormClosing cleanup complete"
    }.GetNewClosure())

    $form.Add_Shown({
        $form.BringToFront(); $form.Activate()
        Write-Log '=== Remove M365 Domain Objects ready ==='
        Write-Log 'Browse to a Discovery folder, click Scan, select sections, then click Connect and Remove.'
    }.GetNewClosure())

    [System.Windows.Forms.Application]::Run($form)
}

function Invoke-RemoveDomainHeadless {
    $discFolder = $DiscoveryFolder.Trim().Trim('"')
    $candidate  = Join-Path $discFolder 'Discovery'
    if ((Split-Path $discFolder -Leaf) -ne 'Discovery' -and (Test-Path $candidate)) { $discFolder = $candidate }
    if (-not (Test-Path $discFolder)) { Write-Host "ERROR: Discovery folder not found: $discFolder"; exit 1 }

    $allSections = $script:SectionDefs
    $toProcess   = if ($Sections -eq 'all') {
        $allSections
    } else {
        $filter = @($Sections -split ',') | ForEach-Object { $_.Trim() }
        $allSections | Where-Object { $_.CsvName -in $filter }
    }

    Write-Host "=== Remove M365 Domain Objects$(if ($WhatIf) { ' [WhatIf]' }) ==="
    Write-Host "Discovery folder : $discFolder"
    Write-Host "Sections         : $($toProcess.Label -join ', ')"
    Write-Host ''

    $exoConnected   = $false
    $graphConnected = $false

    foreach ($sec in $toProcess) {
        $csvPath = Join-Path $discFolder $sec.CsvName
        if (-not (Test-Path $csvPath)) { Write-Host "Skipped (not found): $($sec.CsvName)"; continue }
        $rows = @(Import-Csv -Path $csvPath -Encoding UTF8)
        Write-Host "--- $($sec.Label): $($rows.Count) item(s) ---"
        if ($rows.Count -eq 0) { continue }

        if ($sec.NeedsEXO -and -not $exoConnected -and -not $WhatIf) {
            Write-Host 'Connecting to Exchange Online — sign in when the browser opens...'
            $cmds = @('Remove-AcceptedDomain','Remove-DistributionGroup','Remove-MailContact',
                      'Remove-Mailbox','Remove-UnifiedGroup','Set-Mailbox','Set-DistributionGroup',
                      'Set-UnifiedGroup','Get-Recipient')
            try {
                Connect-ExchangeOnline -ShowBanner:$false -CommandName $cmds -ErrorAction Stop
                $exoConnected = $true; Write-Host 'Exchange Online connected.'
            } catch {
                Write-Host "ERROR: EXO connect failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
                continue
            }
        }
        if ($sec.NeedsGraph -and -not $graphConnected -and -not $WhatIf) {
            Write-Host 'Connecting to Microsoft Graph — sign in when prompted...'
            $scopes = @('Application.ReadWrite.All','Device.ReadWrite.All','Directory.ReadWrite.All')
            try {
                Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
                $graphConnected = $true; Write-Host 'Microsoft Graph connected.'
            } catch {
                Write-Host "ERROR: Graph connect failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
                continue
            }
        }

        $ok = 0; $fail = 0; $skip = 0

        switch ($sec.CsvName) {
            '01_AcceptedDomains.csv' {
                $removable = @($rows | Where-Object { $_.IsDefault -ne 'True' })
                $skipCount = $rows.Count - $removable.Count
                if ($skipCount -gt 0) { Write-Host "  Skipping $skipCount default domain(s)"; $skip += $skipCount }
                foreach ($r in $removable) {
                    $name = $r.DomainName
                    if ($WhatIf) { Write-Host "  WhatIf : $name"; $ok++ }
                    else {
                        try { Remove-AcceptedDomain -Identity $name -Confirm:$false -ErrorAction Stop; Write-Host "  Removed: $name"; $ok++ }
                        catch { Write-Host "  FAILED : $name — $($_.Exception.Message.Split([Environment]::NewLine)[0])"; $fail++ }
                    }
                }
            }
            { $_ -in '02_Mailboxes.csv','05_SharedMailboxes.csv' } {
                foreach ($r in $rows) {
                    $id = if ($r.UserPrincipalName) { $r.UserPrincipalName } else { $r.PrimarySmtpAddress }
                    if ($WhatIf) { Write-Host "  WhatIf : $id"; $ok++ }
                    else {
                        try { Remove-Mailbox -Identity $id -Confirm:$false -ErrorAction Stop; Write-Host "  Removed: $id"; $ok++ }
                        catch { Write-Host "  FAILED : $id — $($_.Exception.Message.Split([Environment]::NewLine)[0])"; $fail++ }
                    }
                }
            }
            '03_DistributionGroups.csv' {
                foreach ($r in $rows) {
                    $addr = $r.PrimarySmtpAddress
                    if ($WhatIf) { Write-Host "  WhatIf : $addr"; $ok++ }
                    else {
                        try { Remove-DistributionGroup -Identity $addr -Confirm:$false -ErrorAction Stop; Write-Host "  Removed: $addr"; $ok++ }
                        catch { Write-Host "  FAILED : $addr — $($_.Exception.Message.Split([Environment]::NewLine)[0])"; $fail++ }
                    }
                }
            }
            '04_MailContacts.csv' {
                foreach ($r in $rows) {
                    $addr = $r.ExternalEmailAddress; $dn = $r.DisplayName
                    if ($WhatIf) { Write-Host "  WhatIf : $addr  [$dn]"; $ok++ }
                    else {
                        try { Remove-MailContact -Identity $addr -Confirm:$false -ErrorAction Stop; Write-Host "  Removed: $addr  [$dn]"; $ok++ }
                        catch { Write-Host "  FAILED : $addr — $($_.Exception.Message.Split([Environment]::NewLine)[0])"; $fail++ }
                    }
                }
            }
            '06_M365Groups.csv' {
                foreach ($r in $rows) {
                    $addr = $r.PrimarySmtpAddress; $dn = $r.DisplayName
                    if ($WhatIf) { Write-Host "  WhatIf : $addr  [$dn]"; $ok++ }
                    else {
                        try { Remove-UnifiedGroup -Identity $addr -Confirm:$false -ErrorAction Stop; Write-Host "  Removed: $addr  [$dn]"; $ok++ }
                        catch { Write-Host "  FAILED : $addr — $($_.Exception.Message.Split([Environment]::NewLine)[0])"; $fail++ }
                    }
                }
            }
            '07_AppRegistrations.csv' {
                foreach ($r in $rows) {
                    $appId = $r.AppId; $dn = $r.DisplayName
                    if ($WhatIf) { Write-Host "  WhatIf : $dn  [$appId]"; $ok++ }
                    else {
                        try { Remove-MgApplication -ApplicationId $appId -Confirm:$false -ErrorAction Stop; Write-Host "  Removed: $dn  [$appId]"; $ok++ }
                        catch { Write-Host "  FAILED : $dn — $($_.Exception.Message.Split([Environment]::NewLine)[0])"; $fail++ }
                    }
                }
            }
            '08_EnterpriseApps.csv' {
                foreach ($r in $rows) {
                    $spId = $r.ObjectId; $dn = $r.DisplayName
                    if ($WhatIf) { Write-Host "  WhatIf : $dn  [$spId]"; $ok++ }
                    else {
                        try { Remove-MgServicePrincipal -ServicePrincipalId $spId -Confirm:$false -ErrorAction Stop; Write-Host "  Removed: $dn  [$spId]"; $ok++ }
                        catch { Write-Host "  FAILED : $dn — $($_.Exception.Message.Split([Environment]::NewLine)[0])"; $fail++ }
                    }
                }
            }
            '11_Devices.csv' {
                foreach ($r in $rows) {
                    $objId = $r.DeviceObjectId; $name = $r.DeviceName
                    if ($WhatIf) { Write-Host "  WhatIf : $name  [$objId]"; $ok++ }
                    else {
                        try { Remove-MgDevice -DeviceId $objId -Confirm:$false -ErrorAction Stop; Write-Host "  Removed: $name  [$objId]"; $ok++ }
                        catch { Write-Host "  FAILED : $name — $($_.Exception.Message.Split([Environment]::NewLine)[0])"; $fail++ }
                    }
                }
            }
            '12_ProxyAddresses.csv' {
                $primaries    = @($rows | Where-Object { $_.IsPrimary -eq 'True' })
                $nonPrimaries = @($rows | Where-Object { $_.IsPrimary -ne 'True' })
                if ($primaries.Count -gt 0) {
                    Write-Host "  Skipping $($primaries.Count) primary SMTP address(es) — cannot remove primary addresses"
                    $skip += $primaries.Count
                }
                $byRecipient = $nonPrimaries | Group-Object -Property PrimarySmtpAddress
                foreach ($grp in $byRecipient) {
                    $primaryAddr     = $grp.Name
                    $addressesToDrop = @($grp.Group | ForEach-Object { "$($_.AddressType):$($_.ProxyAddress)" })
                    if ($WhatIf) {
                        foreach ($a in $addressesToDrop) { Write-Host "  WhatIf : remove proxy $a from $primaryAddr" }
                        $ok += $addressesToDrop.Count
                    } else {
                        try {
                            $recip = Get-Recipient -Identity $primaryAddr -ErrorAction Stop
                            $currentProxies = @($recip.EmailAddresses | ForEach-Object { $_.ToString() })
                            $newProxies = @($currentProxies | Where-Object { $a = $_; -not ($addressesToDrop | Where-Object { $_ -ieq $a }) })
                            $removed = $currentProxies.Count - $newProxies.Count
                            if ($removed -eq 0) { Write-Host "  No matching proxies on $primaryAddr"; $skip++ }
                            else {
                                switch ($recip.RecipientTypeDetails) {
                                    { $_ -in 'UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox' } {
                                        Set-Mailbox -Identity $primaryAddr -EmailAddresses $newProxies -ErrorAction Stop
                                    }
                                    'MailUniversalDistributionGroup' {
                                        Set-DistributionGroup -Identity $primaryAddr -EmailAddresses $newProxies -ErrorAction Stop
                                    }
                                    'GroupMailbox' {
                                        Set-UnifiedGroup -Identity $primaryAddr -EmailAddresses $newProxies -ErrorAction Stop
                                    }
                                    default {
                                        Write-Host "  SKIPPED $primaryAddr — unhandled type: $($recip.RecipientTypeDetails)"
                                        $skip++
                                    }
                                }
                                if ($skip -eq 0 -or $recip.RecipientTypeDetails -in 'UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox','MailUniversalDistributionGroup','GroupMailbox') {
                                    foreach ($dropped in $addressesToDrop) { Write-Host "  Removed proxy: $dropped  from: $primaryAddr" }
                                    $ok += $removed
                                }
                            }
                        } catch {
                            Write-Host "  FAILED : $primaryAddr — $($_.Exception.Message.Split([Environment]::NewLine)[0])"; $fail++
                        }
                    }
                }
            }
        }

        Write-Host "  $($sec.Label): $ok processed  |  $fail failed  |  $skip skipped"
        Write-Host ''
    }

    if ($exoConnected) { try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {} }
    if ($graphConnected) { try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {} }
    Write-Host '=== Done ==='
}

if ($DiscoveryFolder) {
    Invoke-RemoveDomainHeadless
} else {
    try {
        Show-RemoveDomainUI
    } catch {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        [System.Windows.Forms.MessageBox]::Show("Failed to start: $($_.Exception.Message)", 'Launch Error', 'OK', 'Error') | Out-Null
    }
}
