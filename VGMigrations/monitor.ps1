#Requires -Version 7.0
# ═════════════════════════════════════════════════════════════════════════════
# PROJECT MONITOR
# ═════════════════════════════════════════════════════════════════════════════

# Load required libraries if not already loaded
if (-not (Get-Variable -Name 'clrBg' -Scope Script -ErrorAction SilentlyContinue)) {
    . "$PSScriptRoot\lib.ps1"
}

function Show-ProjectMonitorForm {

    $WorkloadDefs = $script:FlyWorkloadDefs  # defined once in lib.ps1 (uses .Status and .Report keys)

    $wlSuffix = [ordered]@{
        'SharePoint'='SharePoint'; 'Exchange'='Exchange'; 'OneDrive'='OneDrive'
        'Teams'='Teams'; 'Teams Chat'='Teams Chat'; 'Groups'='Groups'
    }
    try {
        $wlRaw = Get-Content (Join-Path $PSScriptRoot 'workloads.json') -Raw -ErrorAction Stop | ConvertFrom-Json
        foreach ($wlKey in @($wlSuffix.Keys)) {
            if ($wlRaw.PSObject.Properties[$wlKey] -and
                $wlRaw.$wlKey.PSObject.Properties['ProjectSuffix'] -and
                $wlRaw.$wlKey.ProjectSuffix) {
                $wlSuffix[$wlKey] = [string]$wlRaw.$wlKey.ProjectSuffix
            }
        }
    } catch {}

    $getFlyConfig = {
        $p = Join-Path $env:APPDATA 'FlyMigration\config.json'
        if (-not (Test-Path $p)) { return $null }
        try {
            $c      = Get-Content $p -Raw | ConvertFrom-Json
            $secure = $c.EncSecret | ConvertTo-SecureString
            $bstr   = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
            $plain  = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
            return @{ Url = $c.Url; ClientId = $c.ClientId; ClientSecret = $plain }
        } catch { return $null }
    }

    $loadProjects = {
        try {
            $c = Get-Content $script:SharedConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json
            if ($c.PSObject.Properties['MonitorProjects']) { return [string[]]($c.MonitorProjects) }
        } catch {}
        return @()
    }

    # ── Form ─────────────────────────────────────────────────────────────────
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text          = "AvePoint Fly - Project Monitor"
    $Form.Size          = [System.Drawing.Size]::new(1000, 640)
    $Form.MinimumSize   = [System.Drawing.Size]::new(860, 440)
    $Form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $Form.BackColor     = $clrBg
    $Form.Font          = $FontBody
    $_ico = Join-Path $PSScriptRoot 'FlyMigration.ico'
    if (Test-Path $_ico) { $Form.Icon = [System.Drawing.Icon]::new($_ico) }

    # ── Outer layout: footer (Dock=Bottom) + TLP (Dock=Fill) ─────────────────
    # Only two controls on the form. One Dock=Bottom, one Dock=Fill.
    # The TLP owns all three content rows (header / selector bar / DGV).

    # Footer — add first so it claims its space before Fill runs
    $footer = New-Object System.Windows.Forms.Panel
    $footer.Height    = 46
    $footer.Dock      = [System.Windows.Forms.DockStyle]::Bottom
    $footer.BackColor = $clrFooterAlt
    $Form.Controls.Add($footer)

    $dotConn = New-Object System.Windows.Forms.Panel
    $dotConn.Size      = [System.Drawing.Size]::new(12, 12)
    $dotConn.Location  = [System.Drawing.Point]::new(14, 17)
    $dotConn.BackColor = $clrAmber
    $footer.Controls.Add($dotConn)

    $lblConn = New-Object System.Windows.Forms.Label
    $lblConn.Text      = "Checking connection..."
    $lblConn.Location  = [System.Drawing.Point]::new(34, 14)
    $lblConn.Size      = [System.Drawing.Size]::new(440, 20)
    $lblConn.ForeColor = $clrGridText
    $footer.Controls.Add($lblConn)

    $btnConnect = New-Object System.Windows.Forms.Button
    $btnConnect.Text      = "Connect Now"
    $btnConnect.Location  = [System.Drawing.Point]::new(480, 8)
    $btnConnect.Size      = [System.Drawing.Size]::new(100, 30)
    $btnConnect.BackColor = $clrAccent
    $btnConnect.ForeColor = [System.Drawing.Color]::White
    $btnConnect.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnConnect.FlatAppearance.BorderSize = 0
    $btnConnect.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $btnConnect.Visible   = $false
    $footer.Controls.Add($btnConnect)

    # Close button — fixed 90px from right edge
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text      = "Close"
    $btnClose.Size      = [System.Drawing.Size]::new(90, 30)
    $btnClose.Location  = [System.Drawing.Point]::new(900, 8)
    $btnClose.BackColor = $clrCloseRed
    $btnClose.ForeColor = [System.Drawing.Color]::White
    $btnClose.Font      = $FontBold
    $btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnClose.FlatAppearance.BorderSize = 0
    $btnClose.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $btnClose.Add_Click({ $Form.Close() }.GetNewClosure())
    $footer.Controls.Add($btnClose)
    $footer.Add_SizeChanged({ $btnClose.Left = $footer.Width - 100 }.GetNewClosure())

    # Main TLP — 3 rows: header (46px) / selector (50px) / DGV (fill)
    $tlp = New-Object System.Windows.Forms.TableLayoutPanel
    $tlp.Dock        = [System.Windows.Forms.DockStyle]::Fill
    $tlp.RowCount    = 3
    $tlp.ColumnCount = 1
    $tlp.Padding     = [System.Windows.Forms.Padding]::new(0)
    $tlp.Margin      = [System.Windows.Forms.Padding]::new(0)
    [void]$tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 46)))
    [void]$tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 50)))
    [void]$tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$tlp.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $Form.Controls.Add($tlp)

    # Row 0 — blue banner header
    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Dock      = [System.Windows.Forms.DockStyle]::Fill
    $hdr.BackColor = $clrAccent
    $tlp.Controls.Add($hdr, 0, 0)

    $_hdrX = Add-HeaderLogo $hdr 30
    $hdrTitle = New-Object System.Windows.Forms.Label
    $hdrTitle.Text      = "  Project Monitor"
    $hdrTitle.Font      = $FontBold
    $hdrTitle.ForeColor = [System.Drawing.Color]::White
    $hdrTitle.Location  = [System.Drawing.Point]::new($_hdrX, 0)
    $hdrTitle.Size      = [System.Drawing.Size]::new(400, 46)
    $hdrTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdr.Controls.Add($hdrTitle)

    $btnGear = New-Object System.Windows.Forms.Button
    $btnGear.BackColor = $clrAccent
    $btnGear.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnGear.FlatAppearance.BorderSize = 0
    $btnGear.FlatAppearance.MouseOverBackColor = $clrAccentHover
    $btnGear.Size     = [System.Drawing.Size]::new(38, 38)
    $btnGear.Location = [System.Drawing.Point]::new(950, 4)
    $btnGear.Cursor   = [System.Windows.Forms.Cursors]::Hand
    $btnGear.Add_Click({ Show-SettingsDialog })
    if ($script:GearBitmap) {
        $btnGear.Image = $script:GearBitmap; $btnGear.ImageAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    } else {
        $btnGear.Text = [char]0x2699; $btnGear.Font = New-Object System.Drawing.Font("Segoe UI Symbol", 16)
        $btnGear.ForeColor = [System.Drawing.Color]::White
    }
    $hdr.Controls.Add($btnGear)
    $hdr.Add_SizeChanged({ $btnGear.Left = $hdr.ClientSize.Width - 46 }.GetNewClosure())

    # Row 1 — project selector bar (light blue)
    $selBar = New-Object System.Windows.Forms.Panel
    $selBar.Dock      = [System.Windows.Forms.DockStyle]::Fill
    $selBar.BackColor = $clrAccentTint
    $tlp.Controls.Add($selBar, 0, 1)

    $lblPfx = New-Object System.Windows.Forms.Label
    $lblPfx.Text      = "Project:"
    $lblPfx.Location  = [System.Drawing.Point]::new(12, 15)
    $lblPfx.AutoSize  = $true
    $lblPfx.Font      = $FontBold
    $lblPfx.ForeColor = $clrText
    $selBar.Controls.Add($lblPfx)

    $cmbProject = New-Object System.Windows.Forms.ComboBox
    $cmbProject.Location      = [System.Drawing.Point]::new(76, 11)
    $cmbProject.Size          = [System.Drawing.Size]::new(210, 26)
    $cmbProject.Font          = $FontBody
    $cmbProject.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
    $cmbProject.BackColor     = [System.Drawing.Color]::White
    $cmbProject.ForeColor     = $clrText
    $selBar.Controls.Add($cmbProject)

    $saved = & $loadProjects
    foreach ($p in $saved) { [void]$cmbProject.Items.Add($p) }
    $sharedCfg = Read-SharedConfig
    $allPfx = if ($sharedCfg.PSObject.Properties['Prefixes']) { @($sharedCfg.Prefixes) } `
              elseif ($sharedCfg.TenantName) { @($sharedCfg.TenantName) } else { @() }
    foreach ($pfx in $allPfx) {
        if ($pfx -and -not $cmbProject.Items.Contains($pfx)) { [void]$cmbProject.Items.Add($pfx) }
    }
    if ($cmbProject.Items.Count -gt 0) { $cmbProject.SelectedIndex = 0 }

    $btnRef = New-Object System.Windows.Forms.Button
    $btnRef.Text      = "Refresh Now"
    $btnRef.Location  = [System.Drawing.Point]::new(292, 11)
    $btnRef.Size      = [System.Drawing.Size]::new(100, 28)
    $btnRef.BackColor = $clrAccent
    $btnRef.ForeColor = [System.Drawing.Color]::White
    $btnRef.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnRef.FlatAppearance.BorderSize = 0
    $btnRef.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $selBar.Controls.Add($btnRef)

    $chkAuto = New-Object System.Windows.Forms.CheckBox
    $chkAuto.Text      = "Auto refresh:"
    $chkAuto.Location  = [System.Drawing.Point]::new(402, 15)
    $chkAuto.AutoSize  = $true
    $chkAuto.ForeColor = $clrText
    $selBar.Controls.Add($chkAuto)

    $intervalMap = [ordered]@{
        "1 min"  = 60000
        "2 min"  = 120000
        "5 min"  = 300000
        "10 min" = 600000
        "15 min" = 900000
        "30 min" = 1800000
    }
    $cmbInterval = New-Object System.Windows.Forms.ComboBox
    $cmbInterval.Location      = [System.Drawing.Point]::new(497, 12)
    $cmbInterval.Size          = [System.Drawing.Size]::new(76, 26)
    $cmbInterval.Font          = $FontBody
    $cmbInterval.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbInterval.BackColor     = [System.Drawing.Color]::White
    $cmbInterval.ForeColor     = $clrText
    foreach ($k in $intervalMap.Keys) { [void]$cmbInterval.Items.Add($k) }
    $cmbInterval.SelectedIndex = 0
    $selBar.Controls.Add($cmbInterval)

    $lblRefreshed = New-Object System.Windows.Forms.Label
    $lblRefreshed.Text      = "Select a project then press Refresh"
    $lblRefreshed.Location  = [System.Drawing.Point]::new(582, 16)
    $lblRefreshed.AutoSize  = $true
    $lblRefreshed.ForeColor = $clrMuted
    $selBar.Controls.Add($lblRefreshed)

    # Row 2 — DataGridView (fills remaining space)
    $monDgv = New-Object System.Windows.Forms.DataGridView
    $monDgv.Dock                          = [System.Windows.Forms.DockStyle]::Fill
    $monDgv.BackgroundColor               = $clrLogBg
    $monDgv.GridColor                     = $clrGridLine
    $monDgv.BorderStyle                   = [System.Windows.Forms.BorderStyle]::None
    $monDgv.EnableHeadersVisualStyles     = $false
    $monDgv.ColumnHeadersHeightSizeMode   = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $monDgv.ColumnHeadersHeight           = 34
    $monDgv.ColumnHeadersDefaultCellStyle.BackColor = $clrAccent
    $monDgv.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $monDgv.ColumnHeadersDefaultCellStyle.Font      = New-Object System.Drawing.Font($FontBody.FontFamily, $FontBody.Size, [System.Drawing.FontStyle]::Bold)
    $monDgv.ColumnHeadersBorderStyle      = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::Single
    $monDgv.DefaultCellStyle.BackColor    = $clrLogBg
    $monDgv.DefaultCellStyle.ForeColor    = $clrGridText
    $monDgv.DefaultCellStyle.SelectionBackColor = $clrAccentHover
    $monDgv.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
    $monDgv.RowHeadersVisible             = $false
    $monDgv.ReadOnly                      = $true
    $monDgv.AllowUserToAddRows            = $false
    $monDgv.AllowUserToDeleteRows         = $false
    $monDgv.SelectionMode                 = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $monDgv.MultiSelect                   = $false
    $monDgv.RowTemplate.Height            = 26
    $monDgv.AutoSizeColumnsMode           = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $tlp.Controls.Add($monDgv, 0, 2)

    foreach ($cd in @(
        @{ N='Project';      W=200 }
        @{ N='Total';        W=55  }
        @{ N='Not Started';  W=80  }
        @{ N='In Progress';  W=80  }
        @{ N='Complete';     W=75  }
        @{ N='Failed';       W=65  }
        @{ N='Warnings';     W=75  }
        @{ N='Last Refresh'; W=85  }
    )) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.HeaderText = $cd.N; $col.Name = $cd.N -replace ' ',''; $col.FillWeight = $cd.W
        [void]$monDgv.Columns.Add($col)
    }

    # ── Timer ────────────────────────────────────────────────────────────────
    $monTimer = New-Object System.Windows.Forms.Timer
    $monTimer.Interval = 60000

    # ── Connection check ──────────────────────────────────────────────────────
    $tryConnect = {
        $cfg = & $getFlyConfig
        if (-not $cfg) {
            $dotConn.BackColor  = $clrRed
            $lblConn.Text       = "No saved credentials — run Setup AOS Tenant & App first"
            $lblConn.ForeColor  = [System.Drawing.Color]::FromArgb(255, 140, 140)
            $btnConnect.Visible = $false
            return $false
        }
        try {
            if (-not (Get-Module -Name Fly.Client -ListAvailable -ErrorAction SilentlyContinue)) {
                throw "Fly.Client module not found. Run Install-Module Fly.Client."
            }
            Import-Module Fly.Client -ErrorAction Stop
            Connect-Fly -Url $cfg.Url -ClientId $cfg.ClientId -ClientSecret $cfg.ClientSecret -ErrorAction Stop
            $dotConn.BackColor  = $clrGreen
            $lblConn.Text       = "Connected: $($cfg.Url)"
            $lblConn.ForeColor  = [System.Drawing.Color]::FromArgb(120, 230, 120)
            $btnConnect.Visible = $false
            return $true
        } catch {
            $dotConn.BackColor  = $clrAmber
            $lblConn.Text       = "Not connected: $($_.Exception.Message)"
            $lblConn.ForeColor  = [System.Drawing.Color]::FromArgb(235, 195, 80)
            $btnConnect.Visible = $true
            return $false
        }
    }.GetNewClosure()

    # ── Refresh ───────────────────────────────────────────────────────────────
    $doRefresh = {
        param([string]$pfxOverride = '')
        $monTimer.Stop()
        $pfx = if ($pfxOverride) { $pfxOverride } else { $cmbProject.Text.Trim() }
        if ([string]::IsNullOrWhiteSpace($pfx)) { $lblRefreshed.Text = "Select or type a project name"; return }
        $Form.Text = "AvePoint Fly - Project Monitor — $pfx"
        $monDgv.Rows.Clear()
        $tmpDir = [System.IO.Path]::GetTempPath()

        foreach ($wl in $WorkloadDefs.Keys) {
            $projName = "$pfx - $($wlSuffix[$wl])"
            $cmds     = $WorkloadDefs[$wl]
            $tmpFile  = Join-Path $tmpDir ("FlyMon_$($wl -replace '[^A-Za-z0-9]','_')_$(Get-Random).csv")
            $total=0; $notStarted=0; $inProgress=0; $complete=0; $failed=0; $warnings=0
            $rowErr = $null

            try {
                & $cmds.Status -Project $projName -OutFile $tmpFile -ErrorAction Stop
                if (Test-Path $tmpFile) {
                    $rows = @(Import-Csv $tmpFile -ErrorAction Stop)
                    Remove-Item $tmpFile -ErrorAction SilentlyContinue
                    $total = $rows.Count
                    $stCol = $null
                    if ($total -gt 0) {
                        foreach ($cand in @('Stage status','StageStatus','Stage Status','Status','MigrationStatus','Migration Status','State')) {
                            if ($rows[0].PSObject.Properties.Name -contains $cand) { $stCol = $cand; break }
                        }
                    }
                    foreach ($r in $rows) {
                        $st = if ($stCol) { ([string]$r.$stCol).Trim() } else { '' }
                        switch -Regex ($st) {
                            '^(Not started|NotStarted|Not Started)$'                                    { $notStarted++ }
                            '^(Stopped)$'                                                               { $notStarted++ }
                            '^(In progress|In queue|In queue with priority|Scheduled|InProgress|Waiting|Queued)$' { $inProgress++ }
                            '^(Exceptions|Exceptioned|CompletedWithException|FinishedWithException)$'  { $warnings++; $complete++ }
                            '^(Finished|Complete|Completed|Successful|Success)$'                       { $complete++ }
                            '^(Failed)$'                                                                { $failed++ }
                            default                                                                     { $notStarted++ }
                        }
                    }
                }
            } catch {
                $rowErr = $_.Exception.Message
                Remove-Item $tmpFile -ErrorAction SilentlyContinue
            }

            $now  = Get-Date -Format 'HH:mm:ss'
            $vals = if ($rowErr) {
                @($projName, 'ERR', $rowErr, '', '', '', '', $now)
            } else {
                @($projName, $total, $notStarted, $inProgress, $complete, $failed, $warnings, $now)
            }

            [void]$monDgv.Rows.Add($vals)
            $newRow = $monDgv.Rows[$monDgv.Rows.Count - 1]
            if ($rowErr -or $failed -gt 0) {
                $newRow.DefaultCellStyle.BackColor = $clrRowFailBg
                $newRow.DefaultCellStyle.ForeColor = $clrRowFailFg
            } elseif ($warnings -gt 0) {
                $newRow.DefaultCellStyle.BackColor = $clrRowWarnBg
                $newRow.DefaultCellStyle.ForeColor = $clrRowWarnFg
            }
        }
        $lblRefreshed.Text = "Last refresh: $(Get-Date -Format 'HH:mm:ss')"
        if ($chkAuto.Checked) {
            $ms = $intervalMap[$cmbInterval.SelectedItem]
            if ($ms) { $monTimer.Interval = $ms }
            $monTimer.Start()
        }
    }.GetNewClosure()

    $btnRef.Add_Click({ & $doRefresh }.GetNewClosure())
    $btnConnect.Add_Click({ if (& $tryConnect) { & $doRefresh } }.GetNewClosure())
    $monTimer.Add_Tick({ & $doRefresh }.GetNewClosure())

    $applyInterval = {
        $ms = $intervalMap[$cmbInterval.SelectedItem]
        if ($ms) { $monTimer.Interval = $ms }
        if ($chkAuto.Checked) { $monTimer.Stop(); $monTimer.Start() }
    }.GetNewClosure()

    $chkAuto.Add_CheckedChanged({
        if ($chkAuto.Checked) {
            $ms = $intervalMap[$cmbInterval.SelectedItem]
            if ($ms) { $monTimer.Interval = $ms }
            $monTimer.Start()
        } else {
            $monTimer.Stop()
        }
    }.GetNewClosure())
    $cmbInterval.Add_SelectedIndexChanged({ & $applyInterval }.GetNewClosure())
    $cmbProject.Add_DropDown({ $monTimer.Stop() }.GetNewClosure())
    $cmbProject.Add_SelectionChangeCommitted({
        param($s, $e)
        & $doRefresh ([string]$s.SelectedItem)
    }.GetNewClosure())
    $monDgv.Add_CellMouseDoubleClick({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        $projName = [string]$monDgv.Rows[$e.RowIndex].Cells['Project'].Value
        if ([string]::IsNullOrWhiteSpace($projName)) { return }
        $cfg     = Read-SharedConfig
        $baseUrl = if ($cfg.PortalUrl) { $cfg.PortalUrl.TrimEnd('/') } else { $null }
        if (-not $baseUrl) {
            [System.Windows.Forms.MessageBox]::Show(
                "No Portal URL configured. Set it in Settings.",
                "Portal URL not set", 'OK', 'Information') | Out-Null
            return
        }
        try {
            $proj = Get-FlyMigrationProject -Name $projName -ErrorAction SilentlyContinue
            $url  = if ($proj -and $proj.Id) { "$baseUrl/#/project/$($proj.Id)/mappings" } else { $baseUrl }
        } catch { $url = $baseUrl }
        Start-Process $url
    }.GetNewClosure())
    $Form.Add_Shown({
        $ok = & $tryConnect
        if ($ok -and $cmbProject.Text.Trim()) { & $doRefresh }
    }.GetNewClosure())
    $Form.Add_FormClosed({
        $monTimer.Stop()
        $monTimer.Dispose()
        $script:MonitorFormInstance = $null
    }.GetNewClosure())
    $script:MonitorFormInstance = $Form
    $Form.Show()
}

# When run directly (not dot-sourced), show the monitor form
if ($MyInvocation.InvocationName -ne '.') {
    Show-ProjectMonitorForm

    # Keep PowerShell window open while form is displayed
    while ($script:MonitorFormInstance -and -not $script:MonitorFormInstance.IsDisposed) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 100
    }
}
