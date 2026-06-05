#Requires -Version 5.1
<#
.SYNOPSIS
    Update-OnPremUPN.ps1 — Changes UPN and manages email aliases for on-premise AD users from CSV files.

.DESCRIPTION
    This script processes users from CSV files and updates their:
    - UserPrincipalName (UPN)
    - Email address (mail attribute)
    - Proxy addresses (email aliases)

    Designed for on-premise Active Directory environments with the ActiveDirectory module.

.PARAMETER CSVFolder
    Folder containing CSV files with user information.

.PARAMETER SourceDomain
    The current/source domain to search for (e.g., "olddomain.com").

.PARAMETER TargetDomain
    The new/target domain to change to (e.g., "newdomain.com").

.PARAMETER WhatIf
    Preview changes without applying them.

.EXAMPLE
    .\Update-OnPremUPN.ps1 -CSVFolder "C:\Users\Discovery" -SourceDomain "old.com" -TargetDomain "new.com"

.NOTES
    Requires: ActiveDirectory PowerShell module
    CSV Format: Must contain UserPrincipalName column
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$CSVFolder,

    [Parameter(Mandatory=$false)]
    [string]$SourceDomain,

    [Parameter(Mandatory=$false)]
    [string]$TargetDomain,

    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'
$script:RootDir = $PSScriptRoot

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ── Logging ───────────────────────────────────────────────────────────────────
$_logDir = Join-Path $script:RootDir 'logs'
if (-not (Test-Path $_logDir)) { New-Item -ItemType Directory -Path $_logDir -Force | Out-Null }
$script:LogFile = Join-Path $_logDir "update-onprem-upn-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

function Write-Log {
    param([string]$Msg, [string]$Level = 'INFO')
    $ts = Get-Date -Format 'HH:mm:ss'
    "$ts [$Level] $Msg" | Add-Content -Path $script:LogFile -Encoding UTF8 -ErrorAction SilentlyContinue

    if ($script:rtbLog -and -not $script:rtbLog.IsDisposed) {
        $script:rtbLog.SelectionStart = $script:rtbLog.TextLength
        $script:rtbLog.SelectionLength = 0
        $script:rtbLog.SelectionColor = [System.Drawing.Color]::FromArgb(80, 95, 120)
        $script:rtbLog.AppendText("$ts ")

        $lc = switch ($Level) {
            'OK'    { [System.Drawing.Color]::FromArgb(65, 195, 110)  }
            'WARN'  { [System.Drawing.Color]::FromArgb(220, 165, 45)  }
            'ERROR' { [System.Drawing.Color]::FromArgb(225, 80, 80)   }
            default { [System.Drawing.Color]::FromArgb(120, 155, 220) }
        }
        $script:rtbLog.SelectionColor = $lc
        $script:rtbLog.AppendText("[$Level] ")
        $script:rtbLog.SelectionColor = [System.Drawing.Color]::FromArgb(205, 212, 230)
        $script:rtbLog.AppendText("$Msg`n")
        $script:rtbLog.ScrollToCaret()
    }
}

Write-Log "=== Update-OnPremUPN.ps1 started ==="

# ── Check ActiveDirectory module ──────────────────────────────────────────────
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Write-Host "ERROR: ActiveDirectory module not found." -ForegroundColor Red
    Write-Host "This script requires the Active Directory PowerShell module." -ForegroundColor Yellow
    Write-Host "Install RSAT (Remote Server Administration Tools) or run from a Domain Controller." -ForegroundColor Yellow
    exit 1
}

Import-Module ActiveDirectory -ErrorAction Stop

# ── UI Theme ──────────────────────────────────────────────────────────────────
$libPath = Join-Path $script:RootDir 'lib.ps1'
if (Test-Path $libPath) {
    try { . $libPath }
    catch { Write-Log "lib.ps1 load error: $_" 'WARN' }
}

if (-not (Get-Variable -Name 'clrBg' -Scope Script -ErrorAction SilentlyContinue)) {
    $clrBg     = [System.Drawing.Color]::FromArgb(240, 242, 247)
    $clrPanel  = [System.Drawing.Color]::White
    $clrAccent = [System.Drawing.Color]::FromArgb(0, 100, 180)
    $clrText   = [System.Drawing.Color]::FromArgb(28, 28, 32)
    $clrMuted  = [System.Drawing.Color]::FromArgb(100, 108, 120)
    $clrBorder = [System.Drawing.Color]::FromArgb(210, 215, 228)
    $clrLogBg  = [System.Drawing.Color]::FromArgb(26, 27, 38)
    $FontBody  = New-Object System.Drawing.Font('Segoe UI', 9)
    $FontBold  = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
    $FontCap   = New-Object System.Drawing.Font('Segoe UI Semibold', 7.5)
    $FontMono  = New-Object System.Drawing.Font('Consolas', 8.5)
    $FontTitle = New-Object System.Drawing.Font('Segoe UI Semibold', 14)
}

# ── Main UI ───────────────────────────────────────────────────────────────────
function Show-UpdateOnPremUPNUI {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Update On-Premise UPN & Aliases'
    $form.ClientSize = [System.Drawing.Size]::new(900, 900)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.BackColor = $clrBg
    $form.Font = $FontBody
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.MaximizeBox = $false

    # ── Header ────────────────────────────────────────────────────────────────
    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Size = [System.Drawing.Size]::new(900, 56)
    $hdr.Dock = [System.Windows.Forms.DockStyle]::Top
    $hdr.BackColor = $clrAccent
    $form.Controls.Add($hdr)

    $hdrLbl = New-Object System.Windows.Forms.Label
    $hdrLbl.Text = '  Update On-Premise UPN & Email Aliases'
    $hdrLbl.Font = $FontTitle
    $hdrLbl.ForeColor = [System.Drawing.Color]::White
    $hdrLbl.Location = [System.Drawing.Point]::new(12, 0)
    $hdrLbl.Size = [System.Drawing.Size]::new(800, 56)
    $hdrLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdr.Controls.Add($hdrLbl)

    # ── Footer ────────────────────────────────────────────────────────────────
    $footer = New-Object System.Windows.Forms.Panel
    $footer.Height = 46
    $footer.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $footer.BackColor = [System.Drawing.Color]::FromArgb(20, 24, 38)
    $form.Controls.Add($footer)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'
    $btnClose.Size = [System.Drawing.Size]::new(90, 30)
    $btnClose.Location = [System.Drawing.Point]::new(794, 8)
    $btnClose.BackColor = [System.Drawing.Color]::FromArgb(200, 55, 55)
    $btnClose.ForeColor = [System.Drawing.Color]::White
    $btnClose.Font = $FontBold
    $btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnClose.FlatAppearance.BorderSize = 0
    $btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnClose.Add_Click({ $form.Close() })
    $footer.Controls.Add($btnClose)

    # ── Main Panel ────────────────────────────────────────────────────────────
    $card = New-Object System.Windows.Forms.Panel
    $card.Location = [System.Drawing.Point]::new(12, 66)
    $card.Size = [System.Drawing.Size]::new(876, 600)
    $card.BackColor = $clrPanel
    $form.Controls.Add($card)

    $lx = 14
    $y = 12

    # CSV Folder
    $lblFolderCap = New-Object System.Windows.Forms.Label
    $lblFolderCap.Text = 'CSV FOLDER'
    $lblFolderCap.Font = $FontCap
    $lblFolderCap.ForeColor = $clrMuted
    $lblFolderCap.Location = [System.Drawing.Point]::new($lx, $y)
    $lblFolderCap.AutoSize = $true
    $card.Controls.Add($lblFolderCap)
    $y += 18

    $tbFolder = New-Object System.Windows.Forms.TextBox
    $tbFolder.Location = [System.Drawing.Point]::new($lx, $y)
    $tbFolder.Size = [System.Drawing.Size]::new(730, 24)
    $tbFolder.Font = $FontBody
    $tbFolder.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $card.Controls.Add($tbFolder)

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = 'Browse...'
    $btnBrowse.Location = [System.Drawing.Point]::new(752, $y - 2)
    $btnBrowse.Size = [System.Drawing.Size]::new(104, 28)
    $btnBrowse.Font = $FontBold
    $btnBrowse.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnBrowse.FlatAppearance.BorderSize = 0
    $btnBrowse.BackColor = [System.Drawing.Color]::FromArgb(225, 228, 238)
    $btnBrowse.ForeColor = $clrText
    $btnBrowse.Cursor = [System.Windows.Forms.Cursors]::Hand
    $card.Controls.Add($btnBrowse)
    $y += 32

    $btnScan = New-Object System.Windows.Forms.Button
    $btnScan.Text = 'Load CSV Files'
    $btnScan.Location = [System.Drawing.Point]::new($lx, $y)
    $btnScan.Size = [System.Drawing.Size]::new(130, 28)
    $btnScan.Font = $FontBold
    $btnScan.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnScan.FlatAppearance.BorderSize = 0
    $btnScan.BackColor = [System.Drawing.Color]::FromArgb(225, 228, 238)
    $btnScan.ForeColor = $clrText
    $btnScan.Cursor = [System.Windows.Forms.Cursors]::Hand
    $card.Controls.Add($btnScan)

    $lblScanStatus = New-Object System.Windows.Forms.Label
    $lblScanStatus.Text = 'Select a folder containing CSV files with UserPrincipalName column'
    $lblScanStatus.ForeColor = $clrMuted
    $lblScanStatus.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $lblScanStatus.Location = [System.Drawing.Point]::new(152, $y + 8)
    $lblScanStatus.AutoSize = $true
    $card.Controls.Add($lblScanStatus)
    $y += 38

    $sep1 = New-Object System.Windows.Forms.Label
    $sep1.Location = [System.Drawing.Point]::new($lx, $y)
    $sep1.Size = [System.Drawing.Size]::new(848, 1)
    $sep1.BackColor = $clrBorder
    $card.Controls.Add($sep1)
    $y += 10

    # Domain Configuration
    $lblDomainCap = New-Object System.Windows.Forms.Label
    $lblDomainCap.Text = 'DOMAIN MAPPING'
    $lblDomainCap.Font = $FontCap
    $lblDomainCap.ForeColor = $clrMuted
    $lblDomainCap.Location = [System.Drawing.Point]::new($lx, $y)
    $lblDomainCap.AutoSize = $true
    $card.Controls.Add($lblDomainCap)
    $y += 18

    $lblSource = New-Object System.Windows.Forms.Label
    $lblSource.Text = 'Source Domain (current):'
    $lblSource.Location = [System.Drawing.Point]::new($lx, $y + 4)
    $lblSource.AutoSize = $true
    $card.Controls.Add($lblSource)

    $tbSourceDomain = New-Object System.Windows.Forms.TextBox
    $tbSourceDomain.Location = [System.Drawing.Point]::new(180, $y)
    $tbSourceDomain.Size = [System.Drawing.Size]::new(250, 24)
    $tbSourceDomain.Font = $FontBody
    $tbSourceDomain.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $tbSourceDomain.Text = '@olddomain.com'
    $card.Controls.Add($tbSourceDomain)
    $y += 32

    $lblTarget = New-Object System.Windows.Forms.Label
    $lblTarget.Text = 'Target Domain (new):'
    $lblTarget.Location = [System.Drawing.Point]::new($lx, $y + 4)
    $lblTarget.AutoSize = $true
    $card.Controls.Add($lblTarget)

    $tbTargetDomain = New-Object System.Windows.Forms.TextBox
    $tbTargetDomain.Location = [System.Drawing.Point]::new(180, $y)
    $tbTargetDomain.Size = [System.Drawing.Size]::new(250, 24)
    $tbTargetDomain.Font = $FontBody
    $tbTargetDomain.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $tbTargetDomain.Text = '@newdomain.com'
    $card.Controls.Add($tbTargetDomain)
    $y += 38

    $sep2 = New-Object System.Windows.Forms.Label
    $sep2.Location = [System.Drawing.Point]::new($lx, $y)
    $sep2.Size = [System.Drawing.Size]::new(848, 1)
    $sep2.BackColor = $clrBorder
    $card.Controls.Add($sep2)
    $y += 10

    # User List
    $lblUsersCap = New-Object System.Windows.Forms.Label
    $lblUsersCap.Text = 'USERS TO UPDATE'
    $lblUsersCap.Font = $FontCap
    $lblUsersCap.ForeColor = $clrMuted
    $lblUsersCap.Location = [System.Drawing.Point]::new($lx, $y)
    $lblUsersCap.AutoSize = $true
    $card.Controls.Add($lblUsersCap)
    $y += 18

    $lvUsers = New-Object System.Windows.Forms.ListView
    $lvUsers.Location = [System.Drawing.Point]::new($lx, $y)
    $lvUsers.Size = [System.Drawing.Size]::new(848, 280)
    $lvUsers.View = [System.Windows.Forms.View]::Details
    $lvUsers.FullRowSelect = $true
    $lvUsers.GridLines = $true
    $lvUsers.CheckBoxes = $true
    $lvUsers.Font = $FontBody
    $lvUsers.BackColor = [System.Drawing.Color]::FromArgb(250, 251, 253)
    $lvUsers.ForeColor = $clrText
    $lvUsers.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    [void]$lvUsers.Columns.Add('Display Name', 200)
    [void]$lvUsers.Columns.Add('Current UPN', 250)
    [void]$lvUsers.Columns.Add('New UPN', 250)
    [void]$lvUsers.Columns.Add('Status', 140)
    $card.Controls.Add($lvUsers)
    $y += 285

    $lblUserCount = New-Object System.Windows.Forms.Label
    $lblUserCount.Text = '0 users loaded'
    $lblUserCount.ForeColor = $clrMuted
    $lblUserCount.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $lblUserCount.Location = [System.Drawing.Point]::new($lx, $y)
    $lblUserCount.AutoSize = $true
    $card.Controls.Add($lblUserCount)

    $btnSelectAll = New-Object System.Windows.Forms.Button
    $btnSelectAll.Text = 'Select All'
    $btnSelectAll.Location = [System.Drawing.Point]::new(646, $y - 2)
    $btnSelectAll.Size = [System.Drawing.Size]::new(100, 26)
    $btnSelectAll.Font = $FontBold
    $btnSelectAll.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnSelectAll.FlatAppearance.BorderSize = 0
    $btnSelectAll.BackColor = [System.Drawing.Color]::FromArgb(225, 228, 238)
    $btnSelectAll.ForeColor = $clrText
    $btnSelectAll.Cursor = [System.Windows.Forms.Cursors]::Hand
    $card.Controls.Add($btnSelectAll)

    $btnDeselectAll = New-Object System.Windows.Forms.Button
    $btnDeselectAll.Text = 'Deselect All'
    $btnDeselectAll.Location = [System.Drawing.Point]::new(756, $y - 2)
    $btnDeselectAll.Size = [System.Drawing.Size]::new(100, 26)
    $btnDeselectAll.Font = $FontBold
    $btnDeselectAll.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnDeselectAll.FlatAppearance.BorderSize = 0
    $btnDeselectAll.BackColor = [System.Drawing.Color]::FromArgb(225, 228, 238)
    $btnDeselectAll.ForeColor = $clrText
    $btnDeselectAll.Cursor = [System.Windows.Forms.Cursors]::Hand
    $card.Controls.Add($btnDeselectAll)
    $y += 34

    $sep3 = New-Object System.Windows.Forms.Label
    $sep3.Location = [System.Drawing.Point]::new($lx, $y)
    $sep3.Size = [System.Drawing.Size]::new(848, 1)
    $sep3.BackColor = $clrBorder
    $card.Controls.Add($sep3)
    $y += 12

    # Options and Run
    $chkWhatIf = New-Object System.Windows.Forms.CheckBox
    $chkWhatIf.Text = 'WhatIf mode (preview changes only - no updates will be made)'
    $chkWhatIf.Location = [System.Drawing.Point]::new($lx, $y + 6)
    $chkWhatIf.AutoSize = $true
    $chkWhatIf.ForeColor = $clrText
    $chkWhatIf.Checked = $true
    $card.Controls.Add($chkWhatIf)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = 'Update UPNs & Aliases'
    $btnRun.Location = [System.Drawing.Point]::new(486, $y)
    $btnRun.Size = [System.Drawing.Size]::new(180, 34)
    $btnRun.Font = $FontBold
    $btnRun.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnRun.FlatAppearance.BorderSize = 0
    $btnRun.BackColor = $clrAccent
    $btnRun.ForeColor = [System.Drawing.Color]::White
    $btnRun.Cursor = [System.Windows.Forms.Cursors]::Hand
    $card.Controls.Add($btnRun)

    $btnADSync = New-Object System.Windows.Forms.Button
    $btnADSync.Text = 'Run AD Sync'
    $btnADSync.Location = [System.Drawing.Point]::new(676, $y)
    $btnADSync.Size = [System.Drawing.Size]::new(180, 34)
    $btnADSync.Font = $FontBold
    $btnADSync.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnADSync.FlatAppearance.BorderSize = 0
    $btnADSync.BackColor = [System.Drawing.Color]::FromArgb(0, 130, 70)
    $btnADSync.ForeColor = [System.Drawing.Color]::White
    $btnADSync.Cursor = [System.Windows.Forms.Cursors]::Hand
    $card.Controls.Add($btnADSync)

    # ── Progress bar ──────────────────────────────────────────────────────────
    $progress = New-Object System.Windows.Forms.ProgressBar
    $progress.Location = [System.Drawing.Point]::new(12, $card.Bottom + 8)
    $progress.Size = [System.Drawing.Size]::new(876, 8)
    $progress.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $form.Controls.Add($progress)

    # ── Log RTB ───────────────────────────────────────────────────────────────
    $script:rtbLog = New-Object System.Windows.Forms.RichTextBox
    $script:rtbLog.Location = [System.Drawing.Point]::new(12, $progress.Bottom + 8)
    $script:rtbLog.Size = [System.Drawing.Size]::new(876, 162)
    $script:rtbLog.BackColor = $clrLogBg
    $script:rtbLog.ForeColor = [System.Drawing.Color]::FromArgb(200, 210, 230)
    $script:rtbLog.Font = $FontMono
    $script:rtbLog.ReadOnly = $true
    $script:rtbLog.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $script:rtbLog.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    $form.Controls.Add($script:rtbLog)

    # ── Event Handlers ────────────────────────────────────────────────────────

    # Browse
    $btnBrowse.Add_Click({
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description = 'Select the folder containing CSV files'
        $fbd.ShowNewFolderButton = $false
        if ($tbFolder.Text -and (Test-Path $tbFolder.Text)) {
            $fbd.SelectedPath = $tbFolder.Text
        }
        if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $tbFolder.Text = $fbd.SelectedPath
            Write-Log "Folder selected: $($fbd.SelectedPath)"
        }
    })

    # Scan/Load CSVs
    $btnScan.Add_Click({
        $folder = $tbFolder.Text.Trim().Trim('"')
        if (-not $folder -or -not (Test-Path $folder)) {
            [System.Windows.Forms.MessageBox]::Show('Please select a valid folder.', 'Invalid Folder', 'OK', 'Warning') | Out-Null
            return
        }

        Write-Log "=== Loading CSV files from: $folder ==="
        $lblScanStatus.Text = 'Loading CSV files...'
        [System.Windows.Forms.Application]::DoEvents()

        $csvFiles = Get-ChildItem -Path $folder -Filter *.csv -File
        if ($csvFiles.Count -eq 0) {
            Write-Log "No CSV files found in: $folder" 'WARN'
            $lblScanStatus.Text = 'No CSV files found'
            return
        }

        Write-Log "Found $($csvFiles.Count) CSV file(s)"
        $lvUsers.Items.Clear()
        $allUsers = [System.Collections.Generic.List[pscustomobject]]::new()

        foreach ($csvFile in $csvFiles) {
            Write-Log "  Loading: $($csvFile.Name)"
            try {
                $rows = Import-Csv -Path $csvFile.FullName -Encoding UTF8

                # Check for UserPrincipalName column
                if (-not ($rows[0].PSObject.Properties.Name -contains 'UserPrincipalName')) {
                    Write-Log "    Skipped: No UserPrincipalName column" 'WARN'
                    continue
                }

                foreach ($row in $rows) {
                    if ($row.UserPrincipalName) {
                        $allUsers.Add([pscustomobject]@{
                            UserPrincipalName = $row.UserPrincipalName
                            DisplayName = if ($row.DisplayName) { $row.DisplayName } else { $row.UserPrincipalName.Split('@')[0] }
                            CSVFile = $csvFile.Name
                        })
                    }
                }
            } catch {
                Write-Log "    Error reading $($csvFile.Name): $_" 'ERROR'
            }
        }

        Write-Log "Loaded $($allUsers.Count) user(s) from CSV files"

        # Populate ListView
        foreach ($user in $allUsers) {
            $lvi = New-Object System.Windows.Forms.ListViewItem($user.DisplayName)
            $lvi.Checked = $true
            [void]$lvi.SubItems.Add($user.UserPrincipalName)
            [void]$lvi.SubItems.Add('-')  # New UPN placeholder
            [void]$lvi.SubItems.Add('Ready')
            $lvi.Tag = $user
            [void]$lvUsers.Items.Add($lvi)
        }

        $lblUserCount.Text = "$($allUsers.Count) user(s) loaded, $($allUsers.Count) selected"
        $lblScanStatus.Text = "Loaded $($allUsers.Count) user(s) from $($csvFiles.Count) CSV file(s)"
    })

    # Update New UPN preview when domain changes
    $updateNewUPNs = {
        $src = $tbSourceDomain.Text.Trim()
        $tgt = $tbTargetDomain.Text.Trim()

        if (-not $src -or -not $tgt) { return }

        # Normalize domain format (ensure @ prefix)
        if (-not $src.StartsWith('@')) { $src = '@' + $src }
        if (-not $tgt.StartsWith('@')) { $tgt = '@' + $tgt }

        foreach ($lvi in $lvUsers.Items) {
            $currentUPN = $lvi.SubItems[1].Text
            if ($currentUPN -like "*$src") {
                $newUPN = $currentUPN -replace [regex]::Escape($src) + '$', $tgt
                $lvi.SubItems[2].Text = $newUPN
            } else {
                $lvi.SubItems[2].Text = $currentUPN
            }
        }
    }

    $tbSourceDomain.Add_TextChanged($updateNewUPNs)
    $tbTargetDomain.Add_TextChanged($updateNewUPNs)

    # Select/Deselect All
    $btnSelectAll.Add_Click({
        foreach ($lvi in $lvUsers.Items) { $lvi.Checked = $true }
        $checked = ($lvUsers.Items | Where-Object { $_.Checked }).Count
        $lblUserCount.Text = "$($lvUsers.Items.Count) user(s) loaded, $checked selected"
    })

    $btnDeselectAll.Add_Click({
        foreach ($lvi in $lvUsers.Items) { $lvi.Checked = $false }
        $lblUserCount.Text = "$($lvUsers.Items.Count) user(s) loaded, 0 selected"
    })

    $lvUsers.Add_ItemChecked({
        $checked = ($lvUsers.Items | Where-Object { $_.Checked }).Count
        $lblUserCount.Text = "$($lvUsers.Items.Count) user(s) loaded, $checked selected"
    })

    # Run Update
    $btnRun.Add_Click({
        $selectedUsers = @($lvUsers.Items | Where-Object { $_.Checked })
        if ($selectedUsers.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('Please select at least one user.', 'No Users Selected', 'OK', 'Warning') | Out-Null
            return
        }

        $src = $tbSourceDomain.Text.Trim()
        $tgt = $tbTargetDomain.Text.Trim()

        if (-not $src -or -not $tgt) {
            [System.Windows.Forms.MessageBox]::Show('Please enter both source and target domains.', 'Missing Domain', 'OK', 'Warning') | Out-Null
            return
        }

        # Normalize domain format
        if (-not $src.StartsWith('@')) { $src = '@' + $src }
        if (-not $tgt.StartsWith('@')) { $tgt = '@' + $tgt }

        $whatIfMode = $chkWhatIf.Checked

        $result = [System.Windows.Forms.MessageBox]::Show(
            "Update $($selectedUsers.Count) user(s)?`n`nSource: $src`nTarget: $tgt`n`nMode: $(if ($whatIfMode) { 'WhatIf (preview only)' } else { 'LIVE UPDATE' })",
            'Confirm Update',
            'YesNo',
            'Question'
        )

        if ($result -ne 'Yes') { return }

        Write-Log "=== Starting UPN update ===" 'OK'
        Write-Log "Source domain: $src"
        Write-Log "Target domain: $tgt"
        Write-Log "Users to update: $($selectedUsers.Count)"
        Write-Log "WhatIf mode: $whatIfMode"

        $btnRun.Enabled = $false
        $btnScan.Enabled = $false
        $btnBrowse.Enabled = $false
        $progress.Value = 0

        $success = 0
        $failed = 0
        $skipped = 0

        for ($i = 0; $i -lt $selectedUsers.Count; $i++) {
            $lvi = $selectedUsers[$i]
            $currentUPN = $lvi.SubItems[1].Text

            # Calculate progress
            $pct = [int](($i / $selectedUsers.Count) * 100)
            try { $progress.Value = $pct } catch { }

            # Skip if UPN doesn't match source domain
            if ($currentUPN -notlike "*$src") {
                Write-Log "SKIP: $currentUPN (doesn't match source domain)" 'WARN'
                $lvi.SubItems[3].Text = 'Skipped'
                $skipped++
                continue
            }

            $newUPN = $currentUPN -replace [regex]::Escape($src) + '$', $tgt

            try {
                # Find AD user
                $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$currentUPN'" -Properties proxyAddresses, mail -ErrorAction Stop

                if (-not $adUser) {
                    Write-Log "NOT FOUND: $currentUPN" 'ERROR'
                    $lvi.SubItems[3].Text = 'Not Found in AD'
                    $failed++
                    continue
                }

                if ($whatIfMode) {
                    Write-Log "WHATIF: $currentUPN → $newUPN" 'WARN'
                    $lvi.SubItems[3].Text = 'WhatIf'
                    $success++
                } else {
                    # Update UPN
                    Set-ADUser -Identity $adUser -UserPrincipalName $newUPN -ErrorAction Stop

                    # Update mail attribute
                    $newMail = $newUPN
                    Set-ADUser -Identity $adUser -EmailAddress $newMail -ErrorAction Stop

                    # Update proxyAddresses
                    if ($adUser.proxyAddresses) {
                        $newProxyAddresses = @()
                        $primaryUpdated = $false

                        foreach ($proxy in $adUser.proxyAddresses) {
                            if ($proxy -like "SMTP:*$src") {
                                # Update primary SMTP
                                $newProxy = $proxy -replace [regex]::Escape($src) + '$', $tgt
                                $newProxyAddresses += $newProxy
                                $primaryUpdated = $true
                            } elseif ($proxy -like "smtp:*$src") {
                                # Update alias
                                $newProxy = $proxy -replace [regex]::Escape($src) + '$', $tgt
                                $newProxyAddresses += $newProxy
                            } else {
                                # Keep unchanged
                                $newProxyAddresses += $proxy
                            }
                        }

                        Set-ADUser -Identity $adUser -Replace @{proxyAddresses = $newProxyAddresses} -ErrorAction Stop
                    }

                    Write-Log "SUCCESS: $currentUPN → $newUPN" 'OK'
                    $lvi.SubItems[3].Text = 'Updated'
                    $success++
                }
            } catch {
                Write-Log "FAILED: $currentUPN - $_" 'ERROR'
                $lvi.SubItems[3].Text = 'Failed'
                $failed++
            }

            [System.Windows.Forms.Application]::DoEvents()
        }

        try { $progress.Value = 100 } catch { }

        Write-Log "=== Update complete ===" 'OK'
        Write-Log "Success: $success  |  Failed: $failed  |  Skipped: $skipped"

        $btnRun.Enabled = $true
        $btnScan.Enabled = $true
        $btnBrowse.Enabled = $true

        [System.Windows.Forms.MessageBox]::Show(
            "Update complete.`n`nSuccess: $success`nFailed: $failed`nSkipped: $skipped",
            'Complete',
            'OK',
            'Information'
        ) | Out-Null
    })

    # AD Sync
    $btnADSync.Add_Click({
        Write-Log "AD Sync button clicked"
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Trigger Azure AD Connect sync on VOL-ane-aad1?`n`nThis will start a delta sync cycle to synchronize your on-premise changes to Entra ID.",
            'Confirm AD Sync',
            'YesNo',
            'Question'
        )

        if ($result -ne 'Yes') { return }

        try {
            Write-Log "Connecting to VOL-ane-aad1..."
            $btnADSync.Enabled = $false
            $btnADSync.Text = 'Syncing...'
            [System.Windows.Forms.Application]::DoEvents()

            $syncScript = {
                Import-Module ADSync -ErrorAction Stop
                Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Stop
            }

            Write-Log "Creating remote session to VOL-ane-aad1..."
            $session = New-PSSession -ComputerName 'VOL-ane-aad1' -ErrorAction Stop
            Write-Log "Remote session established. Invoking AD Sync..."
            $syncResult = Invoke-Command -Session $session -ScriptBlock $syncScript -ErrorAction Stop
            Remove-PSSession $session

            Write-Log "AD Sync completed: $($syncResult.Result)" 'OK'

            [System.Windows.Forms.MessageBox]::Show(
                "AD Sync initiated successfully.`n`nResult: $($syncResult.Result)`n`nChanges will sync to Entra ID within a few minutes.",
                'AD Sync Complete',
                'OK',
                'Information'
            ) | Out-Null
        } catch {
            Write-Log "AD Sync failed: $_" 'ERROR'
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to run AD Sync:`n`n$_`n`nEnsure you have permission to connect to VOL-ane-aad1 and run AD Sync commands.",
                'AD Sync Error',
                'OK',
                'Error'
            ) | Out-Null
        } finally {
            $btnADSync.Enabled = $true
            $btnADSync.Text = 'Run AD Sync'
        }
    })

    Write-Log "UI initialized. Ready to load CSV files."
    $form.ShowDialog() | Out-Null
}

# ═════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═════════════════════════════════════════════════════════════════════════════
Show-UpdateOnPremUPNUI
