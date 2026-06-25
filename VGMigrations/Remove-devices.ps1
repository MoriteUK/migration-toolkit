#Requires -Version 7.0
<#
.SYNOPSIS
    Remove-devices.ps1 — Remove Entra ID / Intune registered devices from a CSV (GUI).

.DESCRIPTION
    Reads a CSV with DeviceId, DeviceName, and OwnerUPN columns, connects to
    Microsoft Graph, and deletes each device from Entra ID.  Supports a WhatIf
    mode that lists devices without deleting them.  Requires user confirmation
    before performing live deletions.  Graph work runs in a background runspace
    so the WinForms UI stays responsive.

.NOTES
    Dependency : lib.ps1 (colours, fonts, Write-Log), settings.ps1 (Show-SettingsDialog)
    Requires   : Microsoft.Graph.Identity.DirectoryManagement
    Log file   : logs\remove-devices-<timestamp>.log

    Change log
    ----------
    2026-05-27  Added settings.ps1 load + gear icon so the settings dialog is
                reachable from within this screen.
                Startup logging now captures lib/settings load results to file
                before the UI opens.
                Error message from lib.ps1 load failure is now captured and logged.
#>

param(
    [string]$DiscoveryFolder = '',  # reads 11_Devices.csv from this folder (headless mode)
    [string]$CsvFile         = '',  # OR provide a direct CSV path (headless mode)
    [switch]$WhatIf
)

$libPath      = Join-Path $PSScriptRoot 'lib.ps1'
$settingsPath = Join-Path $PSScriptRoot 'settings.ps1'

# ── File logging — initialised before lib load so any load error is captured ──
$_logDir = Join-Path $PSScriptRoot 'logs'
if (-not (Test-Path $_logDir)) { New-Item -ItemType Directory -Path $_logDir -Force | Out-Null }
$script:LogFile = Join-Path $_logDir "remove-devices-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

function _RawLog {
    param([string]$Msg)
    "$(Get-Date -Format 'HH:mm:ss.fff') $Msg" | Add-Content -Path $script:LogFile -Encoding UTF8 -ErrorAction SilentlyContinue
}

_RawLog "=== Remove-devices.ps1 started  PID=$PID  PSVersion=$($PSVersionTable.PSVersion) ==="
_RawLog "PSScriptRoot : $PSScriptRoot"
_RawLog "lib.ps1      : $libPath  exists=$(Test-Path $libPath)"

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

_RawLog "settings.ps1 : $settingsPath  exists=$(Test-Path $settingsPath)"
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

function Show-RemoveDevicesUI {

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'Remove Devices'
    $form.ClientSize      = [System.Drawing.Size]::new(620, 580)
    $form.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.BackColor       = $clrBg
    $form.Font            = $FontBody
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.MaximizeBox     = $false
    $_ico = Join-Path $PSScriptRoot 'FlyMigration.ico'
    if (Test-Path $_ico) { $form.Icon = [System.Drawing.Icon]::new($_ico) }

    # ── Header ────────────────────────────────────────────────────────────────
    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Size = [System.Drawing.Size]::new(620, 56); $hdr.Dock = [System.Windows.Forms.DockStyle]::Top
    $hdr.BackColor = $clrAccent; $form.Controls.Add($hdr)
    $_hdrX = Add-HeaderLogo $hdr 36
    $hdrLbl = New-Object System.Windows.Forms.Label
    $hdrLbl.Text = '  Remove Devices'; $hdrLbl.Font = $FontTitle
    $hdrLbl.ForeColor = [System.Drawing.Color]::White
    $hdrLbl.Location  = [System.Drawing.Point]::new($_hdrX, 0)
    $hdrLbl.Size      = [System.Drawing.Size]::new(520, 56)
    $hdrLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdr.Controls.Add($hdrLbl)

    $btnGear = New-Object System.Windows.Forms.Button
    $btnGear.BackColor = $clrAccent; $btnGear.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnGear.FlatAppearance.BorderSize = 0
    $btnGear.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(0, 78, 152)
    $btnGear.Size = [System.Drawing.Size]::new(38, 38); $btnGear.Location = [System.Drawing.Point]::new(574, 9)
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
    $btnClose.Location = [System.Drawing.Point]::new(514, 8)
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
    $card.Size      = [System.Drawing.Size]::new(596, 138)
    $card.BackColor = $clrPanel
    $form.Controls.Add($card)

    $lx = 16; $ex = 120; $y = 14

    # CSV row
    $lbFile = New-Object System.Windows.Forms.Label
    $lbFile.Text = 'Devices CSV:'; $lbFile.Font = $FontBold; $lbFile.ForeColor = $clrText
    $lbFile.Location = [System.Drawing.Point]::new($lx, $y + 5); $lbFile.AutoSize = $true
    $card.Controls.Add($lbFile)
    $txtFile = New-Object System.Windows.Forms.TextBox
    $txtFile.Location = [System.Drawing.Point]::new($ex, $y); $txtFile.Size = [System.Drawing.Size]::new(348, 24)
    $txtFile.Font = $FontBody; $txtFile.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $txtFile.ReadOnly = $true; $txtFile.BackColor = [System.Drawing.Color]::FromArgb(245, 246, 250)
    $card.Controls.Add($txtFile)
    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = 'Browse...'; $btnBrowse.Location = [System.Drawing.Point]::new($ex + 354, $y - 2)
    $btnBrowse.Size = [System.Drawing.Size]::new(86, 28); $btnBrowse.Font = $FontBold
    $btnBrowse.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnBrowse.FlatAppearance.BorderSize = 0
    $btnBrowse.BackColor = [System.Drawing.Color]::FromArgb(225, 228, 238); $btnBrowse.ForeColor = $clrText
    $btnBrowse.Cursor = [System.Windows.Forms.Cursors]::Hand
    $card.Controls.Add($btnBrowse)
    $y += 38

    # Hint
    $lbHint = New-Object System.Windows.Forms.Label
    $lbHint.Text = 'CSV must have columns: DeviceId (or DeviceObjectId), DeviceName, OwnerUPN'
    $lbHint.ForeColor = $clrMuted; $lbHint.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $lbHint.Location = [System.Drawing.Point]::new($lx, $y); $lbHint.AutoSize = $true
    $card.Controls.Add($lbHint)
    $y += 24

    # WhatIf + Run
    $chkWhatIf = New-Object System.Windows.Forms.CheckBox
    $chkWhatIf.Text = 'WhatIf  (list devices only - nothing deleted)'
    $chkWhatIf.Location = [System.Drawing.Point]::new($lx, $y + 5)
    $chkWhatIf.AutoSize = $true; $chkWhatIf.ForeColor = $clrText
    $card.Controls.Add($chkWhatIf)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = 'Connect and Delete'; $btnRun.Location = [System.Drawing.Point]::new(422, $y)
    $btnRun.Size = [System.Drawing.Size]::new(158, 32); $btnRun.Font = $FontBold
    $btnRun.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnRun.FlatAppearance.BorderSize = 0
    $btnRun.BackColor = $clrRed; $btnRun.ForeColor = [System.Drawing.Color]::White
    $btnRun.Cursor = [System.Windows.Forms.Cursors]::Hand
    $card.Controls.Add($btnRun)
    $y += 40

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = 'Select a CSV file then click Connect and Delete.'
    $lblStatus.ForeColor = $clrMuted; $lblStatus.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $lblStatus.Location = [System.Drawing.Point]::new($lx, $y); $lblStatus.Size = [System.Drawing.Size]::new(570, 18)
    $card.Controls.Add($lblStatus)

    # ── Progress + Log ────────────────────────────────────────────────────────
    $progress = New-Object System.Windows.Forms.ProgressBar
    $progress.Location = [System.Drawing.Point]::new(12, $card.Bottom + 8)
    $progress.Size     = [System.Drawing.Size]::new(596, 8)
    $progress.Style    = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $form.Controls.Add($progress)

    $script:rtbLog = New-Object System.Windows.Forms.RichTextBox
    $script:rtbLog.Location    = [System.Drawing.Point]::new(12, $progress.Bottom + 8)
    $script:rtbLog.Size        = [System.Drawing.Size]::new(596, $form.ClientSize.Height - $progress.Bottom - 8 - 46 - 8)
    $script:rtbLog.BackColor   = $clrLogBg
    $script:rtbLog.ForeColor   = [System.Drawing.Color]::FromArgb(200, 210, 230)
    $script:rtbLog.Font        = $FontMono
    $script:rtbLog.ReadOnly    = $true
    $script:rtbLog.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $script:rtbLog.ScrollBars  = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    $form.Controls.Add($script:rtbLog)

    $script:devTimer    = $null
    $script:devRunspace = $null
    $script:devPS       = $null
    $script:tickBusy    = $false

    # ── Browse ────────────────────────────────────────────────────────────────
    $btnBrowse.Add_Click({
        _RawLog "Browse dialog opened"
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Title = 'Select Devices CSV'; $ofd.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
        $ofd.InitialDirectory = $PSScriptRoot
        if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            _RawLog "File selected: $($ofd.FileName)"
            $txtFile.Text = $ofd.FileName
        } else {
            _RawLog "Browse dialog cancelled"
        }
    }.GetNewClosure())

    # ── Run ───────────────────────────────────────────────────────────────────
    $btnRun.Add_Click({
        Write-Log "Run clicked — CSV='$($txtFile.Text)'  WhatIf=$($chkWhatIf.Checked)"
        if (-not $txtFile.Text) {
            Write-Log 'Validation failed: no CSV selected' 'WARN'
            [System.Windows.Forms.MessageBox]::Show('Please select a CSV file.', 'Missing Input', 'OK', 'Warning') | Out-Null; return
        }

        # Load CSV on UI thread
        Write-Log "Reading CSV: $($txtFile.Text)"
        $devices = $null
        try { $devices = @(Import-Csv -Path $txtFile.Text -Encoding UTF8) }
        catch {
            Write-Log "CSV read failed: $($_.Exception.Message)" 'ERROR'
            [System.Windows.Forms.MessageBox]::Show("Could not read CSV:`n$($_.Exception.Message)", 'File Error', 'OK', 'Error') | Out-Null; return
        }
        if ($devices.Count -eq 0) {
            Write-Log 'Validation failed: CSV is empty' 'WARN'
            [System.Windows.Forms.MessageBox]::Show('CSV file is empty.', 'No Data', 'OK', 'Warning') | Out-Null; return
        }
        $cols = @($devices[0].PSObject.Properties.Name)
        Write-Log "CSV columns: $($cols -join ', ')"
        # Check for DeviceId or DeviceObjectId column
        $deviceIdColumn = $null
        if ($cols -contains 'DeviceId') {
            $deviceIdColumn = 'DeviceId'
        } elseif ($cols -contains 'DeviceObjectId') {
            $deviceIdColumn = 'DeviceObjectId'
        } else {
            Write-Log "Validation failed: 'DeviceId' or 'DeviceObjectId' column not found. Columns: $($cols -join ', ')" 'WARN'
            [System.Windows.Forms.MessageBox]::Show(
                "CSV must contain a 'DeviceId' or 'DeviceObjectId' column.`nFound: $($cols -join ', ')", 'Column Missing', 'OK', 'Warning') | Out-Null; return
        }
        Write-Log "CSV valid: $($devices.Count) device(s) loaded"

        if (-not $chkWhatIf.Checked) {
            Write-Log "Showing deletion confirmation for $($devices.Count) device(s)"
            $confirm = [System.Windows.Forms.MessageBox]::Show(
                "This will permanently delete $($devices.Count) device(s) from Entra ID.`n`nThis cannot be undone. Proceed?",
                'Confirm Deletion', 'YesNo', 'Warning')
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
                Write-Log 'Deletion cancelled by user at confirmation prompt' 'WARN'
                return
            }
            Write-Log 'User confirmed deletion'
        }

        $btnRun.Enabled = $false; $btnBrowse.Enabled = $false; $btnClose.Enabled = $false
        $progress.Value = 0
        _RawLog "Buttons locked — starting runspace"

        Write-Log '=== Remove Devices started ==='
        Write-Log "CSV: $($txtFile.Text)  ($($devices.Count) device(s))"
        if ($chkWhatIf.Checked) { Write-Log 'WhatIf mode - no deletions will occur.' 'WARN' }

        $whatIf      = $chkWhatIf.Checked
        $logFilePath = $script:LogFile

        $rs = [hashtable]::Synchronized(@{
            Done       = $false
            FatalError = $null
            Success    = 0
            Failed     = 0
            Done_i     = 0
            Total      = $devices.Count
            LogQueue   = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())
        })

        $script:devRunspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $script:devRunspace.ApartmentState = 'STA'
        $script:devRunspace.ThreadOptions  = 'ReuseThread'
        $script:devRunspace.Open()

        $script:devPS = [System.Management.Automation.PowerShell]::Create()
        $script:devPS.Runspace = $script:devRunspace

        [void]$script:devPS.AddScript({
            param($devices, $whatIf, $rs, $logFilePath, $deviceIdColumn)

            function QLog {
                param([string]$Msg, [string]$Level = 'INFO')
                $rs.LogQueue.Enqueue("[$Level] $Msg")
                try { "$(Get-Date -Format 'HH:mm:ss') [$Level] $Msg" | Add-Content $logFilePath -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
            }

            try {
                QLog 'Checking for Microsoft.Graph.Identity.DirectoryManagement module...'
                if (-not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Identity.DirectoryManagement' -ErrorAction SilentlyContinue)) {
                    throw 'Module not installed: Microsoft.Graph.Identity.DirectoryManagement. Run: Install-Module Microsoft.Graph -Scope CurrentUser'
                }
                Import-Module Microsoft.Graph.Identity.DirectoryManagement -DisableNameChecking -ErrorAction Stop

                try { Set-MgGraphOption -DisableLoginByWAM $true -ErrorAction SilentlyContinue } catch {}

                QLog 'Connecting to Microsoft Graph - sign in when the browser opens...'
                Connect-MgGraph -Scopes 'Device.ReadWrite.All', 'Directory.ReadWrite.All' -NoWelcome -ErrorAction Stop
                $ctx = Get-MgContext
                QLog "Connected as $($ctx.Account)" 'OK'

                $i = 0
                foreach ($device in $devices) {
                    $i++
                    $rs.Done_i = $i
                    $name  = $device.DeviceName
                    $id    = $device.$deviceIdColumn
                    $owner = if ($device.PSObject.Properties['OwnerUPN']) { $device.OwnerUPN } else { '' }

                    if ($whatIf) {
                        QLog "[$i/$($devices.Count)] WhatIf - would delete: $name  [DeviceId: $id  Owner: $owner]" 'WARN'
                        $rs.Success++; continue
                    }

                    try {
                        Remove-MgDevice -DeviceId $id -Confirm:$false -ErrorAction Stop
                        QLog "[$i/$($devices.Count)] Deleted: $name  [DeviceId: $id  Owner: $owner]" 'OK'
                        $rs.Success++
                    } catch {
                        QLog "[$i/$($devices.Count)] Failed: $name  $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'
                        $rs.Failed++
                    }
                }

                QLog "Complete. Deleted: $($rs.Success)  Failed: $($rs.Failed)"
            } catch {
                $rs.FatalError = $_.Exception.Message
                QLog "Fatal: $($_.Exception.Message)" 'ERROR'
            } finally {
                $rs.Done = $true
            }
        })
        [void]$script:devPS.AddParameters(@{ devices = $devices; whatIf = $whatIf; rs = $rs; logFilePath = $logFilePath; deviceIdColumn = $deviceIdColumn })
        $script:devHandle = $script:devPS.BeginInvoke()
        $lblStatus.Text = 'Connecting to Microsoft Graph - sign in when prompted...'

        $rs2 = $rs
        $script:devTimer = New-Object System.Windows.Forms.Timer
        $script:devTimer.Interval = 300
        $script:devTimer.Add_Tick({
            if ($script:tickBusy) { return }
            $script:tickBusy = $true
            try {
                while ($rs2.LogQueue.Count -gt 0) {
                    $raw   = $rs2.LogQueue.Dequeue()
                    $level = if ($raw -match '^\[(\w+)\]') { $matches[1] } else { 'INFO' }
                    Write-Log ($raw -replace '^\[\w+\] ', '') $level
                }
                if ($rs2.Total -gt 0 -and $rs2.Done_i -gt 0) {
                    $progress.Value = [int]([math]::Min(($rs2.Done_i / $rs2.Total) * 99, 99))
                    $lblStatus.Text = "Processing $($rs2.Done_i) / $($rs2.Total)  Deleted: $($rs2.Success)  Failed: $($rs2.Failed)"
                }
                if ($rs2.Done) {
                    $script:devTimer.Stop(); $progress.Value = 100
                    $lblStatus.Text = if ($rs2.FatalError) { "Failed - $($rs2.FatalError)" }
                                      else { "Done - Deleted: $($rs2.Success)  Failed: $($rs2.Failed)" }
                    Write-Log (if ($rs2.FatalError) { '=== Remove Devices failed ===' } else { '=== Remove Devices complete ===' }) (if ($rs2.FatalError) { 'ERROR' } else { 'OK' })
                    $btnRun.Enabled = $true; $btnBrowse.Enabled = $true; $btnClose.Enabled = $true
                    try { $script:devRunspace.Close(); $script:devRunspace.Dispose() } catch {}
                }
            } finally { $script:tickBusy = $false }
        }.GetNewClosure())
        $script:devTimer.Start()
    }.GetNewClosure())

    $form.Add_FormClosing({
        param($s, $e)
        _RawLog "FormClosing event  timerRunning=$($script:devTimer -and $script:devTimer.Enabled)"
        if ($script:devTimer -and $script:devTimer.Enabled) {
            $r = [System.Windows.Forms.MessageBox]::Show('Deletion is still running. Close anyway?', 'In Progress', 'YesNo', 'Warning')
            if ($r -eq [System.Windows.Forms.DialogResult]::No) {
                _RawLog "FormClosing cancelled — user chose to keep running"
                $e.Cancel = $true; return
            }
            _RawLog "FormClosing confirmed by user despite running timer"
        }
        _RawLog "FormClosing: cleaning up timer/PS/runspace"
        if ($script:devTimer)    { try { $script:devTimer.Stop();    $script:devTimer.Dispose() }    catch { _RawLog "Timer dispose error: $_" } }
        if ($script:devPS)       { try { $script:devPS.Stop();       $script:devPS.Dispose() }        catch { _RawLog "PS dispose error: $_" } }
        if ($script:devRunspace) { try { $script:devRunspace.Close(); $script:devRunspace.Dispose() } catch { _RawLog "Runspace dispose error: $_" } }
        _RawLog "FormClosing cleanup complete"
    }.GetNewClosure())

    $form.Add_Shown({ $form.BringToFront(); $form.Activate() }.GetNewClosure())
    [System.Windows.Forms.Application]::Run($form)
}

function Invoke-RemoveDevicesHeadless {
    # Resolve CSV path
    $csvPath = ''
    if ($CsvFile) {
        $csvPath = $CsvFile.Trim().Trim('"')
    } elseif ($DiscoveryFolder) {
        $folder    = $DiscoveryFolder.Trim().Trim('"')
        $candidate = Join-Path $folder 'Discovery'
        if ((Split-Path $folder -Leaf) -ne 'Discovery' -and (Test-Path $candidate)) { $folder = $candidate }
        $csvPath = Join-Path $folder '11_Devices.csv'
    }

    if (-not $csvPath -or -not (Test-Path $csvPath)) {
        Write-Host "ERROR: Device CSV not found: $csvPath"
        Write-Host 'Usage: Remove-devices.ps1 -DiscoveryFolder <path>  OR  -CsvFile <path>'
        exit 1
    }

    $devices = @(Import-Csv -Path $csvPath -Encoding UTF8)
    Write-Host "=== Remove Devices$(if ($WhatIf) { ' [WhatIf]' }) ==="
    Write-Host "CSV          : $csvPath"
    Write-Host "Device count : $($devices.Count)"

    if ($devices.Count -eq 0) { Write-Host 'No devices in CSV.'; exit 0 }

    # Detect ID column
    $cols = @($devices[0].PSObject.Properties.Name)
    $idCol = if ($cols -contains 'DeviceObjectId') { 'DeviceObjectId' }
             elseif ($cols -contains 'DeviceId')   { 'DeviceId' }
             else { Write-Host "ERROR: CSV must have a 'DeviceObjectId' or 'DeviceId' column. Found: $($cols -join ', ')"; exit 1 }
    Write-Host "ID column    : $idCol"
    Write-Host ''

    Write-Host 'Connecting to Microsoft Graph...'
    try {
        Connect-MgGraph -Scopes 'Device.ReadWrite.All','Directory.ReadWrite.All' `
            -NoWelcome -ErrorAction Stop
        Write-Host 'Connected.'
    } catch {
        Write-Host "ERROR: Could not connect to Microsoft Graph: $($_.Exception.Message.Split([Environment]::NewLine)[0])"
        exit 1
    }

    $ok = 0; $fail = 0; $i = 0
    foreach ($d in $devices) {
        $i++
        $id    = $d.$idCol
        $name  = $d.DeviceName
        $owner = if ($d.PSObject.Properties['OwnerUPN']) { $d.OwnerUPN } else { '' }

        if (-not $id) { Write-Host "  [$i/$($devices.Count)] SKIPPED — no ID: $name"; continue }

        if ($WhatIf) {
            Write-Host "  [$i/$($devices.Count)] WhatIf : $name  [$id]$(if ($owner) { "  (owner: $owner)" })"
            $ok++
        } else {
            try {
                Remove-MgDevice -DeviceId $id -Confirm:$false -ErrorAction Stop
                Write-Host "  [$i/$($devices.Count)] Deleted : $name  [$id]$(if ($owner) { "  (owner: $owner)" })"
                $ok++
            } catch {
                Write-Host "  [$i/$($devices.Count)] FAILED  : $name — $($_.Exception.Message.Split([Environment]::NewLine)[0])"
                $fail++
            }
        }
    }

    Write-Host ''
    Write-Host "=== Complete: $ok deleted  |  $fail failed ==="

    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
}

if ($DiscoveryFolder -or $CsvFile) {
    Invoke-RemoveDevicesHeadless
} else {
    Write-Log '=== Remove-devices.ps1 starting ==='
    try { Show-RemoveDevicesUI }
    catch {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        [System.Windows.Forms.MessageBox]::Show("Failed to start: $($_.Exception.Message)", 'Launch Error', 'OK', 'Error') | Out-Null
    }
}
