#Requires -Version 7.0
<#
.SYNOPSIS
    Get-DomainDevices.ps1 — Export Entra registered devices for a specific domain (GUI).

.DESCRIPTION
    Connects to Microsoft Graph, retrieves all member users whose UPN ends with
    @<domain>, fetches their registered devices, and exports the results to a CSV.

.NOTES
    Dependency : lib.ps1 (colours, fonts, Write-Log), settings.ps1 (Show-SettingsDialog)
    Requires   : Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement
    Log file   : logs\get-domain-devices-<timestamp>.log

    Change log
    ----------
    2026-06-17  Initial version.
#>

$libPath      = Join-Path $PSScriptRoot 'lib.ps1'
$settingsPath = Join-Path $PSScriptRoot 'settings.ps1'

# ── File logging — initialised before lib load so any load error is captured ──
$_logDir = Join-Path $PSScriptRoot 'logs'
if (-not (Test-Path $_logDir)) { New-Item -ItemType Directory -Path $_logDir -Force | Out-Null }
$script:LogFile = Join-Path $_logDir "get-domain-devices-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

function _RawLog {
    param([string]$Msg)
    "$(Get-Date -Format 'HH:mm:ss.fff') $Msg" | Add-Content -Path $script:LogFile -Encoding UTF8 -ErrorAction SilentlyContinue
}

_RawLog "=== Get-DomainDevices.ps1 started  PID=$PID  PSVersion=$($PSVersionTable.PSVersion) ==="
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
}

function Show-GetDomainDevicesUI {

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'Get Domain Devices'
    $form.ClientSize      = [System.Drawing.Size]::new(620, 600)
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
    $hdrLbl.Text = '  Get Domain Devices'; $hdrLbl.Font = $FontTitle
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
    $card.Size      = [System.Drawing.Size]::new(596, 160)
    $card.BackColor = $clrPanel
    $form.Controls.Add($card)

    $lx = 16; $ex = 130; $y = 14

    # Domain row
    $lbDomain = New-Object System.Windows.Forms.Label
    $lbDomain.Text = 'Domain:'; $lbDomain.Font = $FontBold; $lbDomain.ForeColor = $clrText
    $lbDomain.Location = [System.Drawing.Point]::new($lx, $y + 5); $lbDomain.AutoSize = $true
    $card.Controls.Add($lbDomain)

    $txtDomain = New-Object System.Windows.Forms.TextBox
    $txtDomain.Location = [System.Drawing.Point]::new($ex, $y); $txtDomain.Size = [System.Drawing.Size]::new(440, 24)
    $txtDomain.Font = $FontBody; $txtDomain.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    try { $txtDomain.PlaceholderText = 'e.g. contoso.com' } catch {}
    $card.Controls.Add($txtDomain)
    $y += 38

    # Output CSV row
    $lbOut = New-Object System.Windows.Forms.Label
    $lbOut.Text = 'Save CSV to:'; $lbOut.Font = $FontBold; $lbOut.ForeColor = $clrText
    $lbOut.Location = [System.Drawing.Point]::new($lx, $y + 5); $lbOut.AutoSize = $true
    $card.Controls.Add($lbOut)

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Location = [System.Drawing.Point]::new($ex, $y); $txtOut.Size = [System.Drawing.Size]::new(348, 24)
    $txtOut.Font = $FontBody; $txtOut.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $txtOut.ReadOnly = $true; $txtOut.BackColor = [System.Drawing.Color]::FromArgb(245, 246, 250)
    $card.Controls.Add($txtOut)

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
    $lbHint.Text = 'Exports: OwnerUPN, OwnerName, DeviceName, OS, TrustType, LastSignIn'
    $lbHint.ForeColor = $clrMuted; $lbHint.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $lbHint.Location = [System.Drawing.Point]::new($lx, $y); $lbHint.AutoSize = $true
    $card.Controls.Add($lbHint)
    $y += 24

    # Run button
    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = 'Connect and Export'; $btnRun.Location = [System.Drawing.Point]::new(424, $y)
    $btnRun.Size = [System.Drawing.Size]::new(156, 32); $btnRun.Font = $FontBold
    $btnRun.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnRun.FlatAppearance.BorderSize = 0
    $btnRun.BackColor = [System.Drawing.Color]::FromArgb(18, 140, 60)
    $btnRun.ForeColor = [System.Drawing.Color]::White
    $btnRun.Cursor = [System.Windows.Forms.Cursors]::Hand
    $card.Controls.Add($btnRun)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = 'Enter a domain name then click Connect and Export.'
    $lblStatus.ForeColor = $clrMuted; $lblStatus.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $lblStatus.Location = [System.Drawing.Point]::new($lx, $y + 7); $lblStatus.Size = [System.Drawing.Size]::new(390, 18)
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
        $domain = $txtDomain.Text.Trim().ToLower().TrimStart('@') -replace '[\\/:*?"<>|]', '_'
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Title  = 'Save devices CSV as...'
        $sfd.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
        $sfd.FileName = if ($domain) { "${domain}_devices.csv" } else { 'devices.csv' }
        $sfd.InitialDirectory = $env:USERPROFILE
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            _RawLog "Save path selected: $($sfd.FileName)"
            $txtOut.Text = $sfd.FileName
        }
    }.GetNewClosure())

    # Auto-suggest CSV filename when domain is typed
    $txtDomain.Add_Leave({
        $d = $txtDomain.Text.Trim().ToLower().TrimStart('@') -replace '[\\/:*?"<>|]', '_'
        if ($d -and -not $txtOut.Text) {
            $txtOut.Text = Join-Path $env:USERPROFILE "${d}_devices.csv"
        }
    }.GetNewClosure())

    # ── Run ───────────────────────────────────────────────────────────────────
    $btnRun.Add_Click({
        $domain  = $txtDomain.Text.Trim().ToLower().TrimStart('@')
        $outPath = $txtOut.Text.Trim()

        if ([string]::IsNullOrWhiteSpace($domain)) {
            [System.Windows.Forms.MessageBox]::Show('Enter a domain name.', 'Missing Input', 'OK', 'Warning') | Out-Null; return
        }
        if ([string]::IsNullOrWhiteSpace($outPath)) {
            [System.Windows.Forms.MessageBox]::Show('Choose a location to save the CSV.', 'Missing Input', 'OK', 'Warning') | Out-Null; return
        }

        Write-Log "Run clicked — domain='$domain'  out='$outPath'"

        $btnRun.Enabled = $false; $btnBrowse.Enabled = $false; $btnClose.Enabled = $false
        $progress.Value = 0; $progress.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
        $lblStatus.Text = 'Connecting to Microsoft Graph — sign in when the browser opens...'
        $script:rtbLog.Clear()
        Write-Log "=== Get Domain Devices started  domain=$domain ==="

        $logFilePath = $script:LogFile

        $rs = [hashtable]::Synchronized(@{
            Done        = $false
            FatalError  = $null
            UserCount   = 0
            DeviceCount = 0
            Done_i      = 0
            Total       = 0
            LogQueue    = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())
        })

        $script:devRunspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $script:devRunspace.ApartmentState = 'STA'
        $script:devRunspace.ThreadOptions  = 'ReuseThread'
        $script:devRunspace.Open()

        $script:devPS = [System.Management.Automation.PowerShell]::Create()
        $script:devPS.Runspace = $script:devRunspace

        [void]$script:devPS.AddScript({
            param($domain, $outPath, $rs, $logFilePath)

            function QLog {
                param([string]$Msg, [string]$Level = 'INFO')
                $rs.LogQueue.Enqueue("[$Level] $Msg")
                try { "$(Get-Date -Format 'HH:mm:ss') [$Level] $Msg" | Add-Content $logFilePath -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
            }

            function InvokeGraphGetAll {
                param([string]$Uri)
                $results = [System.Collections.Generic.List[object]]::new()
                $next = $Uri
                do {
                    $resp = Invoke-MgGraphRequest -Uri $next -Method GET -OutputType PSObject
                    $page = @($resp.value)
                    if ($page) { $results.AddRange($page) }
                    $next = if ($resp.PSObject.Properties['@odata.nextLink']) { $resp.'@odata.nextLink' } else { $null }
                } while ($next)
                return $results
            }

            try {
                QLog 'Checking for Microsoft.Graph modules...'
                foreach ($mod in @('Microsoft.Graph.Authentication','Microsoft.Graph.Identity.DirectoryManagement')) {
                    if (-not (Get-Module -ListAvailable -Name $mod -ErrorAction SilentlyContinue)) {
                        throw "Module not installed: $mod. Run: Install-Module Microsoft.Graph -Scope CurrentUser"
                    }
                    Import-Module $mod -DisableNameChecking -ErrorAction Stop
                }

                try { Set-MgGraphOption -DisableLoginByWAM $true -ErrorAction SilentlyContinue } catch {}

                QLog "Connecting to Microsoft Graph (tenant: $domain)..."
                Connect-MgGraph -Scopes 'User.Read.All','Device.Read.All' -TenantId $domain -NoWelcome -ErrorAction Stop
                $ctx = Get-MgContext
                QLog "Connected as $($ctx.Account)  tenant=$($ctx.TenantId)" 'OK'

                # Fetch all member users then filter by UPN suffix
                QLog "Fetching member users..."
                $userUri = "https://graph.microsoft.com/v1.0/users?`$filter=userType eq 'Member'&`$select=id,displayName,userPrincipalName,accountEnabled"
                $allUsers = @(InvokeGraphGetAll $userUri)
                $domainSuffix = "@$domain"
                $domainUsers = @($allUsers | Where-Object {
                    $upn = if ($_.PSObject.Properties['userPrincipalName']) { $_.userPrincipalName } else { $null }
                    $upn -and $upn.ToLower().EndsWith($domainSuffix)
                })
                QLog "Tenant has $($allUsers.Count) member user(s); $($domainUsers.Count) match '@$domain'." 'OK'

                $rs.Total    = $domainUsers.Count
                $rs.UserCount = $domainUsers.Count
                $DeviceResults = [System.Collections.Generic.List[object]]::new()

                $i = 0
                foreach ($u in $domainUsers) {
                    $i++
                    $rs.Done_i = $i
                    $upn     = if ($u.PSObject.Properties['userPrincipalName']) { $u.userPrincipalName } else { '' }
                    $uid     = if ($u.PSObject.Properties['id'])                { $u.id }                else { '' }
                    $dn      = if ($u.PSObject.Properties['displayName'])       { $u.displayName }       else { '' }
                    $enabled = if ($u.PSObject.Properties['accountEnabled'])    { $u.accountEnabled }    else { $null }

                    try {
                        $devUri = "https://graph.microsoft.com/v1.0/users/$uid/registeredDevices?`$select=id,deviceId,displayName,operatingSystem,operatingSystemVersion,trustType,approximateLastSignInDateTime"
                        $devices = @(InvokeGraphGetAll $devUri)

                        foreach ($d in $devices) {
                            $name    = if ($d.PSObject.Properties['displayName'])                       { $d.displayName }                       else { '' }
                            $devId   = if ($d.PSObject.Properties['deviceId'])                          { $d.deviceId }                          else { '' }
                            $objId   = if ($d.PSObject.Properties['id'])                                { $d.id }                                else { '' }
                            $os      = if ($d.PSObject.Properties['operatingSystem'])                   { $d.operatingSystem }                   else { '' }
                            $osVer   = if ($d.PSObject.Properties['operatingSystemVersion'])            { $d.operatingSystemVersion }            else { '' }
                            $trust   = if ($d.PSObject.Properties['trustType'])                         { $d.trustType }                         else { '' }
                            $lastIn  = if ($d.PSObject.Properties['approximateLastSignInDateTime'])     { $d.approximateLastSignInDateTime }     else { '' }

                            $DeviceResults.Add([PSCustomObject]@{
                                OwnerUPN       = $upn
                                OwnerName      = $dn
                                AccountEnabled = $enabled
                                DeviceName     = $name
                                DeviceObjectId = $objId
                                EntraDeviceId  = $devId
                                OS             = $os
                                OSVersion      = $osVer
                                TrustType      = $trust
                                LastSignIn     = $lastIn
                            }) | Out-Null
                        }

                        if ($devices.Count -gt 0) {
                            QLog "  $upn — $($devices.Count) device(s)"
                        }
                    } catch {
                        QLog "  Device lookup failed for ${upn}: $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'WARN'
                    }
                }

                $rs.DeviceCount = $DeviceResults.Count
                QLog "Scan complete — users: $($domainUsers.Count)  devices: $($DeviceResults.Count)" 'OK'

                # Export CSV
                $dir = Split-Path $outPath -Parent
                if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
                $DeviceResults | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8
                QLog "Exported to: $outPath" 'OK'

            } catch {
                $rs.FatalError = $_.Exception.Message
                QLog "Fatal: $($_.Exception.Message)" 'ERROR'
            } finally {
                $rs.Done = $true
            }
        })
        [void]$script:devPS.AddParameters(@{ domain = $domain; outPath = $outPath; rs = $rs; logFilePath = $logFilePath })
        $script:devHandle = $script:devPS.BeginInvoke()

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
                    $pct = [int]([math]::Min(($rs2.Done_i / $rs2.Total) * 99, 99))
                    $progress.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
                    $progress.Value = $pct
                    $lblStatus.Text = "Scanning user $($rs2.Done_i) / $($rs2.Total)..."
                }
                if ($rs2.Done) {
                    $script:devTimer.Stop()
                    $progress.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
                    $progress.Value = 100
                    if ($rs2.FatalError) {
                        $lblStatus.Text = "Failed — $($rs2.FatalError)"
                    } else {
                        $lblStatus.Text = "Done — $($rs2.DeviceCount) device(s) exported"
                    }
                    Write-Log (if ($rs2.FatalError) { '=== Get Domain Devices failed ===' } else { '=== Get Domain Devices complete ===' }) (if ($rs2.FatalError) { 'ERROR' } else { 'OK' })
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
            $r = [System.Windows.Forms.MessageBox]::Show('Export is still running. Close anyway?', 'In Progress', 'YesNo', 'Warning')
            if ($r -eq [System.Windows.Forms.DialogResult]::No) {
                _RawLog "FormClosing cancelled — user chose to keep running"
                $e.Cancel = $true; return
            }
            _RawLog "FormClosing confirmed by user despite running timer"
        }
        _RawLog "FormClosing: cleaning up timer/PS/runspace"
        if ($script:devTimer)    { try { $script:devTimer.Stop();    $script:devTimer.Dispose() }    catch {} }
        if ($script:devPS)       { try { $script:devPS.Stop();       $script:devPS.Dispose() }        catch {} }
        if ($script:devRunspace) { try { $script:devRunspace.Close(); $script:devRunspace.Dispose() } catch {} }
        _RawLog "FormClosing cleanup complete"
    }.GetNewClosure())

    $form.Add_Shown({ $form.BringToFront(); $form.Activate() }.GetNewClosure())
    [System.Windows.Forms.Application]::Run($form)
}

Write-Log '=== Get-DomainDevices.ps1 starting ==='
try { Show-GetDomainDevicesUI }
catch {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    [System.Windows.Forms.MessageBox]::Show("Failed to start: $($_.Exception.Message)", 'Launch Error', 'OK', 'Error') | Out-Null
}
