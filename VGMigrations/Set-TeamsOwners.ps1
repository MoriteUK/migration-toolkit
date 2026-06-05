#Requires -Version 7.0
<#
.SYNOPSIS
    Set-TeamsOwners.ps1 — Add a user as owner to Teams and M365 Groups (GUI).

.DESCRIPTION
    Reads a CSV containing team email addresses ("address" column), connects to
    Microsoft Graph, and adds the specified user as an owner (and member if not
    already one) of each Team or M365 Group.  Runs the Graph work in a background
    runspace and polls every 300 ms so the WinForms UI stays responsive.

.NOTES
    Dependency : lib.ps1 (colours, fonts, Write-Log), settings.ps1 (Show-SettingsDialog)
    Requires   : Microsoft.Graph.Groups, Microsoft.Graph.Users
    Log file   : logs\set-teamsowners-<timestamp>.log

    Change log
    ----------
    2026-05-27  Added settings.ps1 load + gear icon so the settings dialog is
                reachable from within this screen.
                Improved startup logging — lib/settings load results now written
                to the log file before the UI opens.
                Error message from lib.ps1 load failure is now captured and logged.
#>

$libPath      = Join-Path $PSScriptRoot 'lib.ps1'
$settingsPath = Join-Path $PSScriptRoot 'settings.ps1'

# ── File logging — initialised before lib load so any load error is captured ──
$_logDir = Join-Path $PSScriptRoot 'logs'
if (-not (Test-Path $_logDir)) { New-Item -ItemType Directory -Path $_logDir -Force | Out-Null }
$script:LogFile = Join-Path $_logDir "set-teamsowners-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

function _RawLog {
    param([string]$Msg)
    "$(Get-Date -Format 'HH:mm:ss.fff') $Msg" | Add-Content -Path $script:LogFile -Encoding UTF8 -ErrorAction SilentlyContinue
}

_RawLog "=== Set-TeamsOwners.ps1 started  PID=$PID  PSVersion=$($PSVersionTable.PSVersion) ==="
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

# ── Main GUI ──────────────────────────────────────────────────────────────────
function Show-TeamsOwnersUI {

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'Set Teams Owners'
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
    $hdrLbl.Text = '  Set Teams Owners'; $hdrLbl.Font = $FontTitle
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
    $card.Size      = [System.Drawing.Size]::new(596, 176)
    $card.BackColor = $clrPanel
    $form.Controls.Add($card)

    $lx = 16; $ex = 130; $y = 14

    # CSV file row
    $lbFile = New-Object System.Windows.Forms.Label
    $lbFile.Text = 'Teams CSV:'; $lbFile.Font = $FontBold; $lbFile.ForeColor = $clrText
    $lbFile.Location = [System.Drawing.Point]::new($lx, $y + 5); $lbFile.AutoSize = $true
    $card.Controls.Add($lbFile)
    $txtFile = New-Object System.Windows.Forms.TextBox
    $txtFile.Location = [System.Drawing.Point]::new($ex, $y); $txtFile.Size = [System.Drawing.Size]::new(342, 24)
    $txtFile.Font = $FontBody; $txtFile.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $txtFile.ReadOnly = $true; $txtFile.BackColor = [System.Drawing.Color]::FromArgb(245, 246, 250)
    $card.Controls.Add($txtFile)
    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = 'Browse...'; $btnBrowse.Location = [System.Drawing.Point]::new($ex + 348, $y - 2)
    $btnBrowse.Size = [System.Drawing.Size]::new(86, 28); $btnBrowse.Font = $FontBold
    $btnBrowse.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnBrowse.FlatAppearance.BorderSize = 0
    $btnBrowse.BackColor = [System.Drawing.Color]::FromArgb(225, 228, 238); $btnBrowse.ForeColor = $clrText
    $btnBrowse.Cursor = [System.Windows.Forms.Cursors]::Hand
    $card.Controls.Add($btnBrowse)
    $y += 38

    # Owner UPN row
    $lbUser = New-Object System.Windows.Forms.Label
    $lbUser.Text = 'Owner UPN:'; $lbUser.Font = $FontBold; $lbUser.ForeColor = $clrText
    $lbUser.Location = [System.Drawing.Point]::new($lx, $y + 5); $lbUser.AutoSize = $true
    $card.Controls.Add($lbUser)
    $txtUser = New-Object System.Windows.Forms.TextBox
    $txtUser.Location = [System.Drawing.Point]::new($ex, $y); $txtUser.Size = [System.Drawing.Size]::new(434, 24)
    $txtUser.Font = $FontBody; $txtUser.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    try { $txtUser.PlaceholderText = 'user@tenant.onmicrosoft.com' } catch {}
    $card.Controls.Add($txtUser)
    $y += 38

    # WhatIf + column hint row
    $chkWhatIf = New-Object System.Windows.Forms.CheckBox
    $chkWhatIf.Text = 'WhatIf  (resolve groups only - no changes made)'
    $chkWhatIf.Location = [System.Drawing.Point]::new($lx, $y + 5)
    $chkWhatIf.AutoSize = $true; $chkWhatIf.ForeColor = $clrText
    $card.Controls.Add($chkWhatIf)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = 'Connect and Run'; $btnRun.Location = [System.Drawing.Point]::new(426, $y)
    $btnRun.Size = [System.Drawing.Size]::new(154, 32); $btnRun.Font = $FontBold
    $btnRun.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnRun.FlatAppearance.BorderSize = 0
    $btnRun.BackColor = $clrGreen; $btnRun.ForeColor = [System.Drawing.Color]::White
    $btnRun.Cursor = [System.Windows.Forms.Cursors]::Hand
    $card.Controls.Add($btnRun)
    $y += 40

    # CSV hint + status
    $lbHint = New-Object System.Windows.Forms.Label
    $lbHint.Text = 'CSV must have an "address" column containing the team email addresses.'
    $lbHint.ForeColor = $clrMuted; $lbHint.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $lbHint.Location = [System.Drawing.Point]::new($lx, $y); $lbHint.AutoSize = $true
    $card.Controls.Add($lbHint)
    $y += 20

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = 'Select the CSV and enter the owner UPN, then click Connect and Run.'
    $lblStatus.ForeColor = $clrMuted; $lblStatus.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $lblStatus.Location = [System.Drawing.Point]::new($lx, $y); $lblStatus.Size = [System.Drawing.Size]::new(570, 18)
    $card.Controls.Add($lblStatus)

    # ── Progress bar ──────────────────────────────────────────────────────────
    $progress = New-Object System.Windows.Forms.ProgressBar
    $progress.Location = [System.Drawing.Point]::new(12, $card.Bottom + 8)
    $progress.Size     = [System.Drawing.Size]::new(596, 8)
    $progress.Style    = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $form.Controls.Add($progress)

    # ── Log RTB ───────────────────────────────────────────────────────────────
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

    # ── Runspace state ────────────────────────────────────────────────────────
    $script:ownerTimer    = $null
    $script:ownerRunspace = $null
    $script:ownerPS       = $null
    $script:ownerHandle   = $null
    $script:tickBusy      = $false

    # ── Browse handler ────────────────────────────────────────────────────────
    $btnBrowse.Add_Click({
        _RawLog "Browse dialog opened"
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Title = 'Select Teams CSV'
        $ofd.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
        $ofd.InitialDirectory = $PSScriptRoot
        if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            _RawLog "File selected: $($ofd.FileName)"
            $txtFile.Text = $ofd.FileName
        } else {
            _RawLog "Browse dialog cancelled"
        }
    }.GetNewClosure())

    # ── Run handler ───────────────────────────────────────────────────────────
    $btnRun.Add_Click({
        Write-Log "Run clicked — CSV='$($txtFile.Text)'  User='$($txtUser.Text.Trim())'  WhatIf=$($chkWhatIf.Checked)"
        if (-not $txtFile.Text) {
            Write-Log 'Validation failed: no CSV selected' 'WARN'
            [System.Windows.Forms.MessageBox]::Show('Please select the Teams CSV file.', 'Missing Input', 'OK', 'Warning') | Out-Null; return
        }
        $userEmail = $txtUser.Text.Trim()
        if (-not $userEmail) {
            Write-Log 'Validation failed: no owner UPN entered' 'WARN'
            [System.Windows.Forms.MessageBox]::Show('Please enter the owner UPN.', 'Missing Input', 'OK', 'Warning') | Out-Null; return
        }
        if ($userEmail -notmatch '^[^@]+@[^@]+\.[^@]+$') {
            Write-Log "Validation failed: UPN format invalid '$userEmail'" 'WARN'
            [System.Windows.Forms.MessageBox]::Show('Owner UPN does not look like a valid email address.', 'Invalid Input', 'OK', 'Warning') | Out-Null; return
        }

        # Read and validate CSV on UI thread
        $csvPath = $txtFile.Text
        Write-Log "Reading CSV: $csvPath"
        $entries = $null
        try {
            $entries = @(Import-Csv -Path $csvPath -Encoding UTF8)
        } catch {
            Write-Log "CSV read failed: $($_.Exception.Message)" 'ERROR'
            [System.Windows.Forms.MessageBox]::Show("Could not read CSV: $($_.Exception.Message)", 'File Error', 'OK', 'Error') | Out-Null; return
        }
        if ($entries.Count -eq 0) {
            Write-Log 'Validation failed: CSV is empty' 'WARN'
            [System.Windows.Forms.MessageBox]::Show('The CSV file is empty.', 'No Data', 'OK', 'Warning') | Out-Null; return
        }
        $cols = @($entries[0].PSObject.Properties.Name)
        Write-Log "CSV columns: $($cols -join ', ')"
        if ($cols -notcontains 'address') {
            Write-Log "Validation failed: 'address' column not found in CSV" 'WARN'
            [System.Windows.Forms.MessageBox]::Show(
                "CSV must have an 'address' column.`nFound columns: $($cols -join ', ')",
                'Column Not Found', 'OK', 'Warning') | Out-Null; return
        }

        $addresses = @($entries | ForEach-Object { $_.address.Trim() } | Where-Object { $_ })
        if ($addresses.Count -eq 0) {
            Write-Log 'Validation failed: no non-empty addresses in CSV' 'WARN'
            [System.Windows.Forms.MessageBox]::Show('No addresses found in the CSV.', 'No Data', 'OK', 'Warning') | Out-Null; return
        }
        Write-Log "CSV valid: $($addresses.Count) address(es) loaded"

        $btnRun.Enabled = $false; $btnBrowse.Enabled = $false; $btnClose.Enabled = $false
        $progress.Value = 0
        _RawLog "Buttons locked — starting runspace"

        Write-Log '=== Set Teams Owners started ==='
        Write-Log "CSV   : $csvPath  ($($addresses.Count) entries)"
        Write-Log "Owner : $userEmail"
        if ($chkWhatIf.Checked) { Write-Log 'WhatIf mode - no changes will be made.' 'WARN' }

        $whatIf      = $chkWhatIf.Checked
        $logFilePath = $script:LogFile

        $rs = [hashtable]::Synchronized(@{
            Done       = $false
            FatalError = $null
            Added      = 0
            Skipped    = 0
            Failed     = 0
            Done_i     = 0
            Total      = $addresses.Count
            LogQueue   = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())
        })

        $script:ownerRunspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $script:ownerRunspace.ApartmentState = 'STA'
        $script:ownerRunspace.ThreadOptions  = 'ReuseThread'
        $script:ownerRunspace.Open()

        $script:ownerPS = [System.Management.Automation.PowerShell]::Create()
        $script:ownerPS.Runspace = $script:ownerRunspace

        [void]$script:ownerPS.AddScript({
            param($userEmail, $addresses, $whatIf, $rs, $logFilePath)

            function QLog {
                param([string]$Msg, [string]$Level = 'INFO')
                $rs.LogQueue.Enqueue("[$Level] $Msg")
                try { "$(Get-Date -Format 'HH:mm:ss') [$Level] $Msg" | Add-Content $logFilePath -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
            }

            try {
                QLog 'Checking for Microsoft.Graph modules...'
                $missing = @()
                foreach ($m in 'Microsoft.Graph.Groups', 'Microsoft.Graph.Users') {
                    if (-not (Get-Module -ListAvailable -Name $m -ErrorAction SilentlyContinue)) { $missing += $m }
                }
                if ($missing.Count -gt 0) {
                    throw "Required modules not installed: $($missing -join ', '). Run: Install-Module Microsoft.Graph -Scope CurrentUser"
                }

                QLog 'Connecting to Microsoft Graph - sign in when the browser window opens...'
                Connect-MgGraph -Scopes 'Group.ReadWrite.All', 'GroupMember.ReadWrite.All', 'User.Read.All' -ErrorAction Stop
                QLog 'Connected to Microsoft Graph.' 'OK'

                Import-Module Microsoft.Graph.Groups, Microsoft.Graph.Users -DisableNameChecking -ErrorAction Stop

                QLog "Resolving owner account: $userEmail"
                $targetUser = Get-MgUser -Filter "userPrincipalName eq '$userEmail'" -Property Id, DisplayName -ErrorAction Stop |
                              Select-Object -First 1
                if (-not $targetUser) { throw "User not found: $userEmail" }
                QLog "Owner resolved: $($targetUser.DisplayName)  [$($targetUser.Id)]" 'OK'

                $selectProps = 'id,displayName,mail,mailNickname,groupTypes,resourceProvisioningOptions'
                $i = 0
                foreach ($address in $addresses) {
                    $i++
                    $rs.Done_i = $i
                    $mailNick  = $address -replace '@.*', ''

                    try {
                        $group = Get-MgGroup -Filter "mail eq '$address'" -Property $selectProps -ErrorAction Stop |
                                 Select-Object -First 1
                        if (-not $group) {
                            $group = Get-MgGroup -Filter "mailNickname eq '$mailNick'" -Property $selectProps -ErrorAction Stop |
                                     Select-Object -First 1
                        }
                        if (-not $group) {
                            QLog "[$i/$($addresses.Count)] Not found: $address" 'WARN'
                            $rs.Failed++; continue
                        }

                        $isTeam    = $group.AdditionalProperties['resourceProvisioningOptions'] -contains 'Team'
                        $isUnified = $group.GroupTypes -contains 'Unified'
                        $label = if ($isTeam -and $isUnified) { 'Team/M365 Group' }
                                 elseif ($isTeam)              { 'Team' }
                                 elseif ($isUnified)           { 'M365 Group' }
                                 else                          { 'Group' }

                        $owners = @(Get-MgGroupOwner -GroupId $group.Id -Property Id | Select-Object -ExpandProperty Id)
                        if ($owners -contains $targetUser.Id) {
                            QLog "[$i/$($addresses.Count)] Already owner  [$label]: $($group.DisplayName)" 'WARN'
                            $rs.Skipped++; continue
                        }

                        if ($whatIf) {
                            QLog "[$i/$($addresses.Count)] WhatIf - would add as owner  [$label]: $($group.DisplayName)" 'WARN'
                            $rs.Skipped++; continue
                        }

                        $ref = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($targetUser.Id)" }
                        New-MgGroupOwner -GroupId $group.Id -BodyParameter $ref -ErrorAction Stop

                        $members = @(Get-MgGroupMember -GroupId $group.Id -Property Id | Select-Object -ExpandProperty Id)
                        if ($members -notcontains $targetUser.Id) {
                            New-MgGroupMember -GroupId $group.Id -BodyParameter $ref -ErrorAction SilentlyContinue
                        }

                        QLog "[$i/$($addresses.Count)] Added as owner  [$label]: $($group.DisplayName)" 'OK'
                        $rs.Added++
                    } catch {
                        QLog "[$i/$($addresses.Count)] Failed: $address - $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'
                        $rs.Failed++
                    }
                }

                QLog "Complete. Added: $($rs.Added)  Skipped/already owner: $($rs.Skipped)  Failed: $($rs.Failed)"
            } catch {
                $rs.FatalError = $_.Exception.Message
                QLog "Fatal error: $($_.Exception.Message)" 'ERROR'
            } finally {
                $rs.Done = $true
            }
        })
        [void]$script:ownerPS.AddParameters(@{
            userEmail   = $userEmail
            addresses   = $addresses
            whatIf      = $whatIf
            rs          = $rs
            logFilePath = $logFilePath
        })

        $script:ownerHandle = $script:ownerPS.BeginInvoke()
        $lblStatus.Text = 'Connecting to Microsoft Graph - sign in when prompted...'

        $rs2 = $rs
        $script:ownerTimer = New-Object System.Windows.Forms.Timer
        $script:ownerTimer.Interval = 300
        $script:ownerTimer.Add_Tick({
            if ($script:tickBusy) { return }
            $script:tickBusy = $true
            try {
                while ($rs2.LogQueue.Count -gt 0) {
                    $raw   = $rs2.LogQueue.Dequeue()
                    $level = if ($raw -match '^\[(\w+)\]') { $matches[1] } else { 'INFO' }
                    $text  = $raw -replace '^\[\w+\] ', ''
                    Write-Log $text $level
                }
                if ($rs2.Total -gt 0 -and $rs2.Done_i -gt 0) {
                    $pct = [int]([math]::Min(($rs2.Done_i / $rs2.Total) * 99, 99))
                    $progress.Value = $pct
                    $lblStatus.Text = "Processing $($rs2.Done_i) / $($rs2.Total) - Added: $($rs2.Added)  Skipped: $($rs2.Skipped)  Failed: $($rs2.Failed)"
                }
                if ($rs2.Done) {
                    $script:ownerTimer.Stop()
                    $progress.Value = 100
                    if ($rs2.FatalError) {
                        $lblStatus.Text = "Failed - $($rs2.FatalError)"
                        Write-Log '=== Set Teams Owners failed ===' 'ERROR'
                    } else {
                        $lblStatus.Text = "Done - Added: $($rs2.Added)  Skipped: $($rs2.Skipped)  Failed: $($rs2.Failed)"
                        Write-Log '=== Set Teams Owners complete ===' 'OK'
                    }
                    $btnRun.Enabled = $true; $btnBrowse.Enabled = $true; $btnClose.Enabled = $true
                    try { $script:ownerRunspace.Close(); $script:ownerRunspace.Dispose() } catch {}
                }
            } finally {
                $script:tickBusy = $false
            }
        }.GetNewClosure())
        $script:ownerTimer.Start()
    }.GetNewClosure())

    # ── Cleanup on close ──────────────────────────────────────────────────────
    $form.Add_FormClosing({
        param($s, $e)
        _RawLog "FormClosing event  timerRunning=$($script:ownerTimer -and $script:ownerTimer.Enabled)"
        if ($script:ownerTimer -and $script:ownerTimer.Enabled) {
            $r = [System.Windows.Forms.MessageBox]::Show(
                'Processing is still running. Close anyway?', 'In Progress', 'YesNo', 'Warning')
            if ($r -eq [System.Windows.Forms.DialogResult]::No) {
                _RawLog "FormClosing cancelled — user chose to keep running"
                $e.Cancel = $true; return
            }
            _RawLog "FormClosing confirmed by user despite running timer"
        }
        _RawLog "FormClosing: cleaning up timer/PS/runspace"
        if ($script:ownerTimer)    { try { $script:ownerTimer.Stop();    $script:ownerTimer.Dispose() }    catch { _RawLog "Timer dispose error: $_" } }
        if ($script:ownerPS)       { try { $script:ownerPS.Stop();       $script:ownerPS.Dispose() }        catch { _RawLog "PS dispose error: $_" } }
        if ($script:ownerRunspace) { try { $script:ownerRunspace.Close(); $script:ownerRunspace.Dispose() } catch { _RawLog "Runspace dispose error: $_" } }
        _RawLog "FormClosing cleanup complete"
    }.GetNewClosure())

    $form.Add_Shown({
        $form.BringToFront(); $form.Activate()
    }.GetNewClosure())

    [System.Windows.Forms.Application]::Run($form)
}

Write-Log '=== Set-TeamsOwners.ps1 starting ==='
try {
    Show-TeamsOwnersUI
} catch {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    [System.Windows.Forms.MessageBox]::Show(
        "Failed to start: $($_.Exception.Message)", 'Launch Error', 'OK', 'Error') | Out-Null
}
