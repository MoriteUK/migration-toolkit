# ═════════════════════════════════════════════════════════════════════════════
# MIGRATION RUNNER  (Connections + Mappings — single page)
# ═════════════════════════════════════════════════════════════════════════════
function Show-MigrationRunnerForm {

    $WorkloadDefs = $script:FlyWorkloadDefs  # defined once in lib.ps1

    $script:ConfigPath         = Join-Path $env:APPDATA "FlyMigration\config.json"
    $script:WorkloadConfigPath = Join-Path $PSScriptRoot "workloads.json"
    $script:LogDir             = Join-Path $PSScriptRoot "logs"
    $script:LogFile            = Join-Path $script:LogDir ("FlyRunner_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".log")
    if (-not (Test-Path $script:LogDir)) { New-Item -ItemType Directory -Path $script:LogDir -Force | Out-Null }

    $script:connectorJsPath = Join-Path $PSScriptRoot 'fly-connector.js'
    $script:connTable       = $null
    $script:connPendingById = @{}
    $script:connProc        = $null
    $script:connTasks       = @()

    function Read-WorkloadConfig {
        if (-not (Test-Path $script:WorkloadConfigPath)) {
            [ordered]@{
                SharePoint   = [ordered]@{ Policy = ""; Source = ""; Destination = "" }
                Exchange     = [ordered]@{ Policy = ""; Source = ""; Destination = "" }
                OneDrive     = [ordered]@{ Policy = ""; Source = ""; Destination = "" }
                Teams        = [ordered]@{ Policy = ""; Source = ""; Destination = "" }
                'Teams Chat' = [ordered]@{ Policy = ""; Source = ""; Destination = "" }
                Groups       = [ordered]@{ Policy = ""; Source = ""; Destination = "" }
            } | ConvertTo-Json -Depth 3 | Set-Content $script:WorkloadConfigPath -Encoding UTF8
        }
        try   { return Get-Content $script:WorkloadConfigPath -Raw | ConvertFrom-Json }
        catch { return [pscustomobject]@{} }
    }

    function Get-FlyConfig {
        if (-not (Test-Path $script:ConfigPath)) { return $null }
        try {
            $cfg    = Get-Content $script:ConfigPath -Raw | ConvertFrom-Json
            $secure = $cfg.EncSecret | ConvertTo-SecureString
            $bstr   = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
            $plain  = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
            return @{ Url = $cfg.Url; ClientId = $cfg.ClientId; ClientSecret = $plain }
        }
        catch { return $null }
    }

    function Find-WFNodeExe {
        try { $cmd = Get-Command node.exe -ErrorAction Stop; if ($cmd.Source) { return $cmd.Source } } catch {}
        foreach ($p in @(
            "$env:ProgramFiles\nodejs\node.exe",
            "${env:ProgramFiles(x86)}\nodejs\node.exe",
            "$env:LOCALAPPDATA\Programs\nodejs\node.exe"
        )) { if ($p -and (Test-Path $p)) { return $p } }
        return $null
    }

    function Update-ConnSummary {
        $c = ($script:connTable.Select("Status = 'CREATED'")).Count
        $s = ($script:connTable.Select("Status = 'SKIPPED'")).Count
        $f = ($script:connTable.Select("Status = 'FAILED'" )).Count
        $parts = @()
        if ($c) { $parts += "$c created" }
        if ($s) { $parts += "$s skipped" }
        if ($f) { $parts += "$f failed"  }
        $script:lblConnStatus.Text = if ($parts) { $parts -join '  |  ' } else { "Ready" }
    }

    function Add-ConnRow($id, $workload, $connectionName) {
        $row = $script:connTable.NewRow()
        $row['Id']             = $id
        $row['Timestamp']      = (Get-Date).ToString('HH:mm:ss')
        $row['Workload']       = $workload
        $row['ConnectionName'] = $connectionName
        $row['Status']         = 'WORKING'
        $row['Message']        = 'queued'
        $script:connTable.Rows.Add($row)
        $script:connPendingById[$id] = $row
        Update-ConnSummary
        [System.Windows.Forms.Application]::DoEvents()
    }

    function Update-ConnRow($id, $status, $message) {
        if ($script:connPendingById.ContainsKey($id)) {
            $row = $script:connPendingById[$id]
            $row['Timestamp'] = (Get-Date).ToString('HH:mm:ss')
            $row['Status']    = $status
            $row['Message']   = $message
            if ($status -eq 'CREATED') {
                $labelToKey = @{
                    'Exchange Online'       = 'Exchange'
                    'SharePoint Online'     = 'SharePoint'
                    'OneDrive'             = 'OneDrive'
                    'Microsoft Teams'      = 'Teams'
                    'Microsoft Teams Chat' = 'Teams Chat'
                    'Microsoft 365 Groups' = 'Groups'
                }
                $task = $script:connTasks | Where-Object { $_.id -eq $id } | Select-Object -First 1
                if ($task) {
                    $wlKey = $labelToKey[$task.workloadLabel]
                    if ($wlKey) {
                        try {
                            $wlJson = if (Test-Path $script:WorkloadConfigPath) {
                                try { Get-Content $script:WorkloadConfigPath -Raw | ConvertFrom-Json } catch { [pscustomobject]@{} }
                            } else { [pscustomobject]@{} }
                            if (-not $wlJson.PSObject.Properties[$wlKey]) {
                                $wlJson | Add-Member -NotePropertyName $wlKey -NotePropertyValue ([pscustomobject]@{ Policy = ""; Source = ""; Destination = "" })
                            }
                            $wlJson.$wlKey.($task.connKeyName) = $task.connectionName
                            $wlJson | ConvertTo-Json -Depth 3 | Set-Content $script:WorkloadConfigPath -Encoding UTF8
                        } catch {}
                    }
                }
            }
            $script:dgvConn.Refresh()
            Update-ConnSummary
        }
    }

    function Invoke-WFConnLine($line) {
        try { $obj = $line | ConvertFrom-Json -ErrorAction Stop } catch {
            $script:lblConnStatus.Text = "(non-JSON) $line"
            return
        }
        if ($obj.event) {
            switch ($obj.event) {
                'info'     { $script:lblConnStatus.Text = $obj.message }
                'warn'     { $script:lblConnStatus.Text = "WARN: $($obj.message)" }
                'error'    { $script:lblConnStatus.Text = "ERROR: $($obj.message)" }
                'fatal'    { $script:lblConnStatus.Text = "FATAL: $($obj.message)" }
                'login-ok' { $script:lblConnStatus.Text = "Signed in. $($obj.message)" }
                'done'     { $script:lblConnStatus.Text = "Connector finished." }
                default    { $script:lblConnStatus.Text = "$($obj.event): $($obj.message)" }
            }
            [System.Windows.Forms.Application]::DoEvents()
            return
        }
        if ($obj.id -and $obj.status) { Update-ConnRow $obj.id $obj.status $obj.message }
    }

    function Invoke-WFConnector {
        param([string]$Mode, [string[]]$StdinLines = @(), [string]$DisplayName = '')
        $nodeExe = Find-WFNodeExe
        if (-not $nodeExe) {
            [System.Windows.Forms.MessageBox]::Show(
                "Node.js was not found on PATH.`n`nInstall Node 18+ from nodejs.org and reopen this GUI.",
                "Node.js missing", 'OK', 'Error') | Out-Null
            return $false
        }
        if (-not (Test-Path $script:connectorJsPath)) {
            [System.Windows.Forms.MessageBox]::Show(
                "fly-connector.js not found next to this script.`nExpected: $script:connectorJsPath",
                "Connector missing", 'OK', 'Error') | Out-Null
            return $false
        }
        $argList = "`"$($script:connectorJsPath)`" --mode=$Mode"
        if ($DisplayName) {
            $clean    = $DisplayName -replace '[^A-Za-z0-9._-]', '_'
            $argList += " --display-name=$clean"
        }
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName               = $nodeExe
        $psi.WorkingDirectory       = $PSScriptRoot
        $psi.Arguments              = $argList
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.RedirectStandardInput  = $true
        $psi.UseShellExecute        = $false
        $psi.CreateNoWindow         = $true
        $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
        $psi.StandardErrorEncoding  = [System.Text.Encoding]::UTF8
        $script:connProc = New-Object System.Diagnostics.Process
        $script:connProc.StartInfo = $psi
        [void]$script:connProc.Start()
        foreach ($ln in $StdinLines) { $script:connProc.StandardInput.WriteLine($ln) }
        $script:connProc.StandardInput.Close()
        $reader = $script:connProc.StandardOutput
        while (-not $script:connProc.HasExited -or $reader.Peek() -ge 0) {
            while ($reader.Peek() -ge 0) {
                $line = $reader.ReadLine()
                if ($line) { Invoke-WFConnLine $line }
            }
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 50
        }
        while ($reader.Peek() -ge 0) {
            $line = $reader.ReadLine()
            if ($line) { Invoke-WFConnLine $line }
        }
        $exitCode = $script:connProc.ExitCode
        $script:connProc = $null
        return ($exitCode -eq 0)
    }

    # ── FORM ────────────────────────────────────────────────────────────────────
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text            = "AvePoint Fly - Migration Mappings"
    $Form.WindowState     = [System.Windows.Forms.FormWindowState]::Maximized
    $Form.StartPosition   = [System.Windows.Forms.FormStartPosition]::WindowsDefaultBounds
    $Form.BackColor       = $clrBg
    $Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $Form.MinimumSize     = [System.Drawing.Size]::new(1000, 1000)
    $Form.Font            = $FontBody
    $_ico = Join-Path $PSScriptRoot 'FlyMigration.ico'; if (Test-Path $_ico) { $Form.Icon = [System.Drawing.Icon]::new($_ico) }

    # ── HEADER ──────────────────────────────────────────────────────────────────
    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Height    = 46; $hdr.Dock = [System.Windows.Forms.DockStyle]::Top
    $hdr.BackColor = $clrAccent
    $Form.Controls.Add($hdr)

    # ── FOOTER ──────────────────────────────────────────────────────────────────
    $footer = New-Object System.Windows.Forms.Panel
    $footer.Height    = 46
    $footer.Dock      = [System.Windows.Forms.DockStyle]::Bottom
    $footer.BackColor = [System.Drawing.Color]::FromArgb(28, 32, 48)
    $Form.Controls.Add($footer)

    $btnFlyConn = New-Object System.Windows.Forms.Button
    $btnFlyConn.Text      = "Connect"
    $btnFlyConn.Location  = [System.Drawing.Point]::new(16, 8)
    $btnFlyConn.Size      = [System.Drawing.Size]::new(90, 30)
    $btnFlyConn.BackColor = $clrAccent; $btnFlyConn.ForeColor = [System.Drawing.Color]::White
    $btnFlyConn.Font      = $FontBold; $btnFlyConn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnFlyConn.FlatAppearance.BorderSize = 0; $btnFlyConn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $footer.Controls.Add($btnFlyConn)

    $dotFly = New-Object System.Windows.Forms.Panel
    $dotFly.Size      = [System.Drawing.Size]::new(12, 12)
    $dotFly.Location  = [System.Drawing.Point]::new(114, 17)
    $dotFly.BackColor = $clrGrey
    $footer.Controls.Add($dotFly)

    $lblRunStatus = New-Object System.Windows.Forms.Label
    $lblRunStatus.Location  = [System.Drawing.Point]::new(136, 14)
    $lblRunStatus.Size      = [System.Drawing.Size]::new(400, 20)
    $lblRunStatus.Font      = $FontBody
    $lblRunStatus.ForeColor = [System.Drawing.Color]::FromArgb(190, 210, 255)
    $lblRunStatus.Text      = "Connect to Fly first"
    $footer.Controls.Add($lblRunStatus)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text      = "Close"; $btnClose.Size = [System.Drawing.Size]::new(90, 30)
    $btnClose.Location  = [System.Drawing.Point]::new(900, 8)
    $btnClose.BackColor = [System.Drawing.Color]::FromArgb(200, 55, 55)
    $btnClose.ForeColor = [System.Drawing.Color]::White; $btnClose.Font = $FontBold
    $btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnClose.FlatAppearance.BorderSize = 0
    $btnClose.Cursor    = [System.Windows.Forms.Cursors]::Hand; $btnClose.Add_Click({ $Form.Close() })
    $footer.Controls.Add($btnClose)
    $footer.Add_SizeChanged({
        $btnClose.Left      = $footer.Width - 100
        $lblRunStatus.Width = $footer.Width - 240
    })

    # Header controls — gear added BEFORE title so it renders on top
    $_hdrX = Add-HeaderLogo $hdr 30
    $btnGear = New-Object System.Windows.Forms.Button
    $btnGear.BackColor = $clrAccent
    $btnGear.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnGear.FlatAppearance.BorderSize = 0
    $btnGear.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(0, 78, 152)
    $btnGear.Size     = [System.Drawing.Size]::new(38, 38)
    $btnGear.Location = [System.Drawing.Point]::new(900, 4)
    $btnGear.Anchor   = $AnchorTR
    $btnGear.Cursor   = [System.Windows.Forms.Cursors]::Hand; $btnGear.Add_Click({ Show-SettingsDialog })
    if ($script:GearBitmap) {
        $btnGear.Image = $script:GearBitmap; $btnGear.ImageAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    } else {
        $btnGear.Text = [char]0x2699; $btnGear.Font = New-Object System.Drawing.Font("Segoe UI", 16)
        $btnGear.ForeColor = [System.Drawing.Color]::White
    }
    $hdr.Controls.Add($btnGear)
    $hdrTitle = New-Object System.Windows.Forms.Label
    $hdrTitle.Text      = "  Migration Mappings"
    $hdrTitle.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 12)
    $hdrTitle.ForeColor = [System.Drawing.Color]::White
    $hdrTitle.Location  = [System.Drawing.Point]::new($_hdrX, 0)
    $hdrTitle.Size      = [System.Drawing.Size]::new($Form.Width - $_hdrX, 46)
    $hdrTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdrTitle.Anchor    = $AnchorTLR
    $hdr.Controls.Add($hdrTitle)

    # ── MAIN TABLE LAYOUT ───────────────────────────────────────────────────────
    $tlp = New-Object System.Windows.Forms.TableLayoutPanel
    $tlp.Dock        = [System.Windows.Forms.DockStyle]::Fill
    $tlp.ColumnCount = 1
    $tlp.RowCount    = 5
    $tlp.Padding     = New-Object System.Windows.Forms.Padding(8, 6, 8, 6)
    $tlp.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 296))) | Out-Null  # connections
    $tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,  82))) | Out-Null  # project config
    $tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 320))) | Out-Null  # workloads
    $tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,  50))) | Out-Null  # run row
    $tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,  100))) | Out-Null  # log
    $Form.Controls.Add($tlp)

    # ════════════════════════════════════════════════════════════════════════════
    # ROW 0 — CREATE CONNECTIONS
    # ════════════════════════════════════════════════════════════════════════════
    $c0 = New-CardPanel "CREATE CONNECTIONS"
    $tlp.Controls.Add($c0.Parent, 0, 0)

    # Workload checkboxes
    $chkConnExo    = New-Object System.Windows.Forms.CheckBox
    $chkConnSpo    = New-Object System.Windows.Forms.CheckBox
    $chkConnOd4b   = New-Object System.Windows.Forms.CheckBox
    $chkConnTeams  = New-Object System.Windows.Forms.CheckBox
    $chkConnChat   = New-Object System.Windows.Forms.CheckBox
    $chkConnGroups = New-Object System.Windows.Forms.CheckBox
    $connChkDefs = @(
        @{ Ctrl=$chkConnExo;    Text="Exchange";    X= 16 }
        @{ Ctrl=$chkConnSpo;    Text="SharePoint";  X=110 }
        @{ Ctrl=$chkConnOd4b;   Text="OneDrive";    X=210 }
        @{ Ctrl=$chkConnTeams;  Text="Teams";       X=294 }
        @{ Ctrl=$chkConnChat;   Text="Teams Chat";  X=358 }
        @{ Ctrl=$chkConnGroups; Text="M365 Groups"; X=456 }
    )
    foreach ($d in $connChkDefs) {
        $d.Ctrl.Text = $d.Text; $d.Ctrl.Checked = $true
        $d.Ctrl.Font = $FontBody; $d.Ctrl.AutoSize = $true
        $d.Ctrl.Location = [System.Drawing.Point]::new($d.X, 16)
        $c0.Controls.Add($d.Ctrl)
    }

    # Action buttons (right-aligned)
    $btnConnSignIn  = New-Btn $c0 "Sign in to AOS..."   0 44 138 26 $true  $true
    $btnConnStop    = New-Btn $c0 "Stop"                0 44  60 26 $false $true
    $btnConnSaveLog = New-Btn $c0 "Save Log"            0 44  76 26 $false $true
    $btnConnClear   = New-Btn $c0 "Clear"               0 44  58 26 $false $true
    $btnConnLogs    = New-Btn $c0 "Open Logs"           0 44  80 26 $false $true
    $btnConnStop.Enabled = $false

    # Status label
    $script:lblConnStatus = New-Object System.Windows.Forms.Label
    $script:lblConnStatus.Text      = "Ready"
    $script:lblConnStatus.Font      = $FontBody
    $script:lblConnStatus.ForeColor = $clrMuted
    $script:lblConnStatus.Location  = [System.Drawing.Point]::new(16, 76)
    $script:lblConnStatus.AutoSize  = $true
    $c0.Controls.Add($script:lblConnStatus)

    New-HSep $c0 16 94 16

    # DataGridView
    $script:dgvConn = New-Object System.Windows.Forms.DataGridView
    $script:dgvConn.Location          = [System.Drawing.Point]::new(16, 98)
    $script:dgvConn.Size              = [System.Drawing.Size]::new(800, 182)
    $script:dgvConn.ReadOnly          = $true
    $script:dgvConn.AllowUserToAddRows    = $false
    $script:dgvConn.AllowUserToDeleteRows = $false
    $script:dgvConn.AutoSizeRowsMode  = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::None
    $script:dgvConn.RowTemplate.Height = 26
    $script:dgvConn.SelectionMode     = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $script:dgvConn.MultiSelect       = $false
    $script:dgvConn.BorderStyle       = [System.Windows.Forms.BorderStyle]::FixedSingle
    $script:dgvConn.BackgroundColor   = $clrPanel
    $script:dgvConn.GridColor         = $clrBorder
    $script:dgvConn.DefaultCellStyle.BackColor  = $clrPanel
    $script:dgvConn.DefaultCellStyle.ForeColor  = $clrText
    $script:dgvConn.DefaultCellStyle.Font       = $FontBody
    $script:dgvConn.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(245, 246, 250)
    $script:dgvConn.ColumnHeadersDefaultCellStyle.Font       = $FontBold
    $script:dgvConn.ColumnHeadersDefaultCellStyle.BackColor  = $clrBg
    $script:dgvConn.ColumnHeadersDefaultCellStyle.ForeColor  = $clrText
    $script:dgvConn.ColumnHeadersHeight           = 28
    $script:dgvConn.ColumnHeadersHeightSizeMode   = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $script:dgvConn.EnableHeadersVisualStyles     = $false
    $c0.Controls.Add($script:dgvConn)

    $script:connTable = New-Object System.Data.DataTable
    $script:connTable.Columns.Add("Id",             [string]) | Out-Null
    $script:connTable.Columns.Add("Timestamp",      [string]) | Out-Null
    $script:connTable.Columns.Add("Workload",       [string]) | Out-Null
    $script:connTable.Columns.Add("ConnectionName", [string]) | Out-Null
    $script:connTable.Columns.Add("Status",         [string]) | Out-Null
    $script:connTable.Columns.Add("Message",        [string]) | Out-Null
    $script:dgvConn.DataSource = $script:connTable

    $script:dgvConn.Columns["Id"].Visible            = $false
    $script:dgvConn.Columns["Timestamp"].Width       = 72
    $script:dgvConn.Columns["Workload"].Width        = 140
    $script:dgvConn.Columns["ConnectionName"].Width  = 240
    $script:dgvConn.Columns["Status"].Width          = 82
    $script:dgvConn.Columns["Message"].AutoSizeMode  = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill

    $script:dgvConn.Add_CellFormatting({
        param($s, $e)
        if ($e.ColumnIndex -eq $script:dgvConn.Columns["Status"].Index -and $e.Value) {
            $e.CellStyle.Font = $FontBold
            $e.CellStyle.ForeColor = switch ($e.Value) {
                'CREATED' { [System.Drawing.Color]::FromArgb(18, 155, 60)   }
                'FAILED'  { [System.Drawing.Color]::FromArgb(200, 50, 50)   }
                'SKIPPED' { [System.Drawing.Color]::FromArgb(130, 130, 130) }
                'WORKING' { [System.Drawing.Color]::FromArgb(0, 100, 180)   }
                default   { $clrText }
            }
        }
    }.GetNewClosure())

    # Resize the DGV and reposition right-anchored buttons when card resizes
    $c0.Add_SizeChanged({
        $script:dgvConn.Size = [System.Drawing.Size]::new($c0.Width - 32, $c0.Height - 102)
        $rx = $c0.Width - 8
        $btnConnSignIn.Left  = $rx - $btnConnSignIn.Width;  $rx -= ($btnConnSignIn.Width  + 4)
        $btnConnStop.Left    = $rx - $btnConnStop.Width;    $rx -= ($btnConnStop.Width    + 4)
        $btnConnSaveLog.Left = $rx - $btnConnSaveLog.Width; $rx -= ($btnConnSaveLog.Width + 4)
        $btnConnClear.Left   = $rx - $btnConnClear.Width;   $rx -= ($btnConnClear.Width   + 4)
        $btnConnLogs.Left    = $rx - $btnConnLogs.Width
    }.GetNewClosure())

    # ════════════════════════════════════════════════════════════════════════════
    # ROW 1 — PROJECT CONFIGURATION
    # ════════════════════════════════════════════════════════════════════════════
    $c2 = New-CardPanel "STEP 1  -  PROJECT CONFIGURATION"
    $tlp.Controls.Add($c2.Parent, 0, 1)

    New-Lbl $c2 "Customer Prefix" 16 28 | Out-Null
    $tbPrefix = New-Object System.Windows.Forms.ComboBox
    $tbPrefix.Location      = [System.Drawing.Point]::new(16, 44)
    $tbPrefix.Size          = [System.Drawing.Size]::new(220, 24)
    $tbPrefix.Font          = $FontBody
    $tbPrefix.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
    $c2.Controls.Add($tbPrefix)

    $lblConnType = New-Object System.Windows.Forms.Label
    $lblConnType.Text = "Type:"; $lblConnType.Font = $FontBold; $lblConnType.ForeColor = $clrText
    $lblConnType.AutoSize = $true; $lblConnType.Location = [System.Drawing.Point]::new(250, 48)
    $c2.Controls.Add($lblConnType)

    $rdoConnDest = New-Object System.Windows.Forms.RadioButton
    $rdoConnDest.Text = "Destination"; $rdoConnDest.Checked = $true
    $rdoConnDest.Font = $FontBody; $rdoConnDest.AutoSize = $true
    $rdoConnDest.Location = [System.Drawing.Point]::new(290, 44)
    $c2.Controls.Add($rdoConnDest)

    $rdoConnSrc = New-Object System.Windows.Forms.RadioButton
    $rdoConnSrc.Text = "Source"
    $rdoConnSrc.Font = $FontBody; $rdoConnSrc.AutoSize = $true
    $rdoConnSrc.Location = [System.Drawing.Point]::new(400, 44)
    $c2.Controls.Add($rdoConnSrc)

    $btnConnCreate = New-Btn $c2 "Create Connections" 480 40 160 26 $true $false

    $chkCreateProject = New-Object System.Windows.Forms.CheckBox
    $chkCreateProject.Text     = "Create project if not exists"
    $chkCreateProject.Font     = $FontBody
    $chkCreateProject.Location = [System.Drawing.Point]::new(0, 47)
    $chkCreateProject.Size     = [System.Drawing.Size]::new(230, 24)
    $chkCreateProject.Checked  = $true
    $chkCreateProject.Anchor   = $AnchorTR
    $c2.Controls.Add($chkCreateProject)
    $c2.Add_SizeChanged({ $chkCreateProject.Left = $c2.Width - 238 })

    # ════════════════════════════════════════════════════════════════════════════
    # ROW 2 — WORKLOADS
    # ════════════════════════════════════════════════════════════════════════════
    $c3 = New-CardPanel "STEP 2  -  WORKLOADS"
    $tlp.Controls.Add($c3.Parent, 0, 2)

    $chkAll = New-Object System.Windows.Forms.CheckBox
    $chkAll.Text      = "Workload"; $chkAll.Font = $FontBold; $chkAll.ForeColor = $clrText
    $chkAll.Location  = [System.Drawing.Point]::new(16, 27)
    $chkAll.Size      = [System.Drawing.Size]::new(155, 20)
    $chkAll.Checked   = $true; $chkAll.Anchor = $AnchorTL
    $c3.Controls.Add($chkAll)
    New-Lbl $c3 "CSV File"   182  30 $true | Out-Null
    New-Lbl $c3 "Operation"    0  30 $true | Out-Null
    New-Lbl $c3 "Status"       0  30 $true | Out-Null

    $lblOpHdr  = $c3.Controls | Where-Object { $_.Text -eq "Operation" }
    $lblDotHdr = $c3.Controls | Where-Object { $_.Text -eq "Status" }

    New-HSep $c3 16 50 16

    $script:WLControls = [ordered]@{}
    $i = 0
    foreach ($wl in $WorkloadDefs.Keys) {
        $y = 60 + ($i * 40)

        $chk = New-Object System.Windows.Forms.CheckBox
        $chk.Text = $wl; $chk.Font = $FontBold
        $chk.Location = [System.Drawing.Point]::new(16, $y)
        $chk.Size     = [System.Drawing.Size]::new(155, 24)
        $chk.Checked  = $true; $chk.Anchor = $AnchorTL
        $c3.Controls.Add($chk)

        $tbCsv = New-Object System.Windows.Forms.TextBox
        $tbCsv.Location    = [System.Drawing.Point]::new(182, $y)
        $tbCsv.Size        = [System.Drawing.Size]::new(400, 24)
        $tbCsv.Font        = $FontBody
        $tbCsv.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $tbCsv.Text        = "No file selected"
        $tbCsv.ForeColor   = [System.Drawing.Color]::FromArgb(90, 90, 90)
        $tbCsv.Anchor      = $AnchorTLR
        $c3.Controls.Add($tbCsv)

        $btnBrowse = New-Object System.Windows.Forms.Button
        $btnBrowse.Text      = "Browse..."
        $btnBrowse.Font      = $FontBold
        $btnBrowse.Size      = [System.Drawing.Size]::new(76, 26)
        $btnBrowse.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnBrowse.FlatAppearance.BorderSize = 0
        $btnBrowse.BackColor = [System.Drawing.Color]::FromArgb(225, 228, 238)
        $btnBrowse.ForeColor = $clrText
        $btnBrowse.Anchor    = $AnchorTL
        $btnBrowse.Location  = [System.Drawing.Point]::new(0, $y)
        $capturedCsv = $tbCsv; $capturedWl = $wl
        $btnBrowse.Add_Click({
            $d = New-Object System.Windows.Forms.OpenFileDialog
            $d.Filter = "CSV files (*.csv)|*.csv"
            $d.Title  = "Select $capturedWl mapping CSV"
            if ($d.ShowDialog() -eq 'OK') {
                $capturedCsv.Tag       = $d.FileName
                $capturedCsv.ForeColor = $clrText
                try {
                    $rowCount = (Import-Csv $d.FileName | Measure-Object).Count
                    $capturedCsv.Text = "$(Split-Path $d.FileName -Leaf)  ($rowCount rows)"
                } catch { $capturedCsv.Text = $d.FileName }
            }
        }.GetNewClosure())
        $c3.Controls.Add($btnBrowse)

        $cmbOp = New-Object System.Windows.Forms.ComboBox
        $cmbOp.Font = $FontBody; $cmbOp.Size = [System.Drawing.Size]::new(120, 24)
        $cmbOp.Location = [System.Drawing.Point]::new(0, $y)
        $cmbOp.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
        $cmbOp.Anchor = $AnchorTR
        @("Import Only","Verification","Pre-Scan","Migration","Report") | ForEach-Object { $cmbOp.Items.Add($_) | Out-Null }
        $cmbOp.SelectedIndex = 0
        $c3.Controls.Add($cmbOp)

        $dot = New-Object System.Windows.Forms.Label
        $dot.Text = [char]0x25CF; $dot.Font = New-Object System.Drawing.Font("Segoe UI", 13)
        $dot.ForeColor = $clrGrey; $dot.AutoSize = $true
        $dot.Location = [System.Drawing.Point]::new(0, $y); $dot.Anchor = $AnchorTR
        $c3.Controls.Add($dot)

        $script:WLControls[$wl] = @{ Check = $chk; CSV = $tbCsv; Browse = $btnBrowse; Operation = $cmbOp; Dot = $dot; Y = $y }
        $i++
    }

    $chkAll.Add_CheckedChanged({
        foreach ($wl in $script:WLControls.Keys) { $script:WLControls[$wl].Check.Checked = $chkAll.Checked }
    }.GetNewClosure())

    $ySep = 60 + ($WorkloadDefs.Count * 40) + 4
    New-HSep $c3 16 $ySep 16
    $yOut = $ySep + 10
    New-Lbl $c3 "Report folder:" 16 $yOut | Out-Null
    $tbOutFolder  = New-TB $c3 110 ($yOut - 2) 400 180
    $btnOutBrowse = New-Btn $c3 "Browse..." 0 ($yOut - 4) 76 26 $false $true
    $btnOutBrowse.Left = $c3.Width - 84
    $btnOutBrowse.Add_Click({
        $fd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fd.SelectedPath = $tbOutFolder.Text
        if ($fd.ShowDialog() -eq "OK") { $tbOutFolder.Text = $fd.SelectedPath }
    })
    $tbOutFolder.Text = "$env:USERPROFILE\Desktop"

    $c3.Add_SizeChanged({
        $rightEdge = $c3.Width - 16
        $dotW = 22; $opW = 128; $browseW = 76
        $dotLeft    = $rightEdge - $dotW
        $opLeft     = $dotLeft   - 6 - $opW
        $browseLeft = $opLeft    - 8 - $browseW
        $lblOpHdr.Left  = $opLeft; $lblDotHdr.Left = $dotLeft
        foreach ($wl in $script:WLControls.Keys) {
            $wlCtrl = $script:WLControls[$wl]
            $wlCtrl.Dot.Left       = $dotLeft
            $wlCtrl.Operation.Left = $opLeft;  $wlCtrl.Operation.Width = $opW
            $wlCtrl.Browse.Left    = $browseLeft
            $wlCtrl.CSV.Width      = $wlCtrl.Browse.Left - $wlCtrl.CSV.Left - 6
        }
        $btnOutBrowse.Left = $rightEdge - 80
        $tbOutFolder.Width = $btnOutBrowse.Left - $tbOutFolder.Left - 6
    })

    # ════════════════════════════════════════════════════════════════════════════
    # ROW 3 — RUN ROW
    # ════════════════════════════════════════════════════════════════════════════
    $c4 = New-CardPanel ""
    $tlp.Controls.Add($c4.Parent, 0, 3)

    $btnRun = New-Btn $c4 "Run Selected Workloads" 16 12 200 30
    $btnRun.Enabled = $false

    $btnMonitor = New-Btn $c4 "Monitor Projects" 228 12 120 30 $false
    $btnMonitor.Add_Click({ Show-ProjectMonitorForm })

    # ════════════════════════════════════════════════════════════════════════════
    # ROW 4 — LOG
    # ════════════════════════════════════════════════════════════════════════════
    $c5 = New-CardPanel "LOG"
    $tlp.Controls.Add($c5.Parent, 0, 4)

    $script:rtbLog = New-Object System.Windows.Forms.RichTextBox
    $script:rtbLog.Dock        = [System.Windows.Forms.DockStyle]::Fill
    $script:rtbLog.Font        = $FontMono
    $script:rtbLog.BackColor   = $clrLogBg
    $script:rtbLog.ForeColor   = [System.Drawing.Color]::FromArgb(190, 210, 255)
    $script:rtbLog.ReadOnly    = $true
    $script:rtbLog.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $script:rtbLog.ScrollBars  = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    $script:rtbLog.Margin      = New-Object System.Windows.Forms.Padding(16, 26, 8, 8)
    $c5.Controls.Add($script:rtbLog)

    # ════════════════════════════════════════════════════════════════════════════
    # CONNECTION EVENT HANDLERS
    # ════════════════════════════════════════════════════════════════════════════
    $connSvcDef = @(
        [pscustomobject]@{ Label='Exchange Online';       Chk=$chkConnExo    }
        [pscustomobject]@{ Label='SharePoint Online';     Chk=$chkConnSpo    }
        [pscustomobject]@{ Label='OneDrive';              Chk=$chkConnOd4b   }
        [pscustomobject]@{ Label='Microsoft Teams';       Chk=$chkConnTeams  }
        [pscustomobject]@{ Label='Microsoft Teams Chat';  Chk=$chkConnChat   }
        [pscustomobject]@{ Label='Microsoft 365 Groups';  Chk=$chkConnGroups }
    )


    $btnConnSignIn.Add_Click({
        $btnConnSignIn.Enabled = $false; $btnConnCreate.Enabled = $false
        $script:lblConnStatus.Text = "Launching browser for Microsoft SSO sign-in..."
        try {
            Invoke-WFConnector -Mode 'login' | Out-Null
            $script:lblConnStatus.Text = "Sign-in complete."
        } catch {
            $script:lblConnStatus.Text = "Sign-in error: $($_.Exception.Message)"
        } finally {
            $btnConnSignIn.Enabled = $true; $btnConnCreate.Enabled = $true
        }
    }.GetNewClosure())

    $btnConnClear.Add_Click({
        $script:connTable.Rows.Clear(); $script:connPendingById.Clear()
        $script:lblConnStatus.Text = "Results cleared."
        Update-ConnSummary
    }.GetNewClosure())

    $btnConnLogs.Add_Click({
        $logsDir = Join-Path $PSScriptRoot 'logs'
        if (-not (Test-Path $logsDir)) { New-Item -ItemType Directory -Path $logsDir -Force | Out-Null }
        Start-Process explorer.exe $logsDir
    }.GetNewClosure())

    $btnConnSaveLog.Add_Click({
        if ($script:connTable.Rows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Nothing to save.", "Save Log", 'OK', 'Information') | Out-Null; return
        }
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Filter   = "CSV (*.csv)|*.csv"
        $tag = if ($tbPrefix.Text.Trim()) { $tbPrefix.Text.Trim() } else { 'tenant' }
        $dlg.FileName = "fly-connections-$tag-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
        if ($dlg.ShowDialog() -eq 'OK') {
            $script:connTable | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
            $script:lblConnStatus.Text = "Saved: $($dlg.FileName)"
        }
    }.GetNewClosure())

    $btnConnStop.Add_Click({
        if ($script:connProc -and -not $script:connProc.HasExited) {
            try { $script:connProc.Kill() } catch {}
            $script:lblConnStatus.Text = "Cancelled."
        }
    }.GetNewClosure())

    $btnConnCreate.Add_Click({
        $selPfx = $tbPrefix.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($selPfx)) {
            [System.Windows.Forms.MessageBox]::Show("Select a Customer Prefix in Step 1.", "Validation", 'OK', 'Warning') | Out-Null; return
        }
        $cust       = $script:customerList | Where-Object { $_.Prefix -eq $selPfx } | Select-Object -First 1
        $tenantName = if ($cust -and $cust.AccountName) { $cust.AccountName } else { $selPfx }
        $tenantSrch = if ($cust -and $cust.AccountName) { ($cust.AccountName -replace '[^A-Za-z0-9]', '').ToLower() } else { ($selPfx -replace '[^A-Za-z0-9]', '').ToLower() }
        $shared     = Read-SharedConfig
        $credName   = [string]$shared.CredentialsName
        if ([string]::IsNullOrWhiteSpace($credName)) {
            [System.Windows.Forms.MessageBox]::Show("Credentials Name is not configured. Set it in Settings (shared config).", "Validation", 'OK', 'Warning') | Out-Null; return
        }
        $selected = $connSvcDef | Where-Object { $_.Chk.Checked }
        if (-not $selected) {
            [System.Windows.Forms.MessageBox]::Show("Select at least one workload.", "Validation", 'OK', 'Warning') | Out-Null; return
        }
        $connKeyName = if ($rdoConnSrc.Checked) { 'Source' } else { 'Destination' }
        $script:connTasks = @()
        foreach ($svc in $selected) {
            $id   = [guid]::NewGuid().ToString('N').Substring(0, 8)
            $name = "$tenantName - $($svc.Label)"
            $script:connTasks += [pscustomobject]@{ id=$id; workloadLabel=$svc.Label; connectionName=$name; connKeyName=$connKeyName }
            Add-ConnRow $id $svc.Label $name
        }
        $stdinLines = $script:connTasks | ForEach-Object {
            [pscustomobject]@{ id=$_.id; tenantName=$tenantName; tenantSearch=$tenantSrch; workloadLabel=$_.workloadLabel; connectionName=$_.connectionName; credentialsName=$credName } | ConvertTo-Json -Compress
        }
        $btnConnCreate.Enabled = $false; $btnConnSignIn.Enabled = $false; $btnConnStop.Enabled = $true
        $script:lblConnStatus.Text = "Driving the AOS portal..."
        try {
            Invoke-WFConnector -Mode 'create' -StdinLines $stdinLines -DisplayName $tenantName | Out-Null
            $c = ($script:connTable.Select("Status = 'CREATED'")).Count
            $f = ($script:connTable.Select("Status = 'FAILED'" )).Count
            $script:lblConnStatus.Text = "Done. $c created, $f failed. workloads.json updated automatically."
        } catch {
            $script:lblConnStatus.Text = "Error: $($_.Exception.Message)"
        } finally {
            $btnConnCreate.Enabled = $true; $btnConnSignIn.Enabled = $true; $btnConnStop.Enabled = $false
        }
    }.GetNewClosure())

    # ════════════════════════════════════════════════════════════════════════════
    # FLY CONNECT & RUN
    # ════════════════════════════════════════════════════════════════════════════
    $btnFlyConn.Add_Click({
        $btnFlyConn.Enabled = $false; $dotFly.BackColor = $clrGrey
        try {
            $cfg = Get-FlyConfig
            if (-not $cfg) { throw "No saved credentials. Open Settings (gear icon) to configure Fly API URL, Client ID and Secret." }
            if ([string]::IsNullOrWhiteSpace($cfg.ClientId) -or [string]::IsNullOrWhiteSpace($cfg.ClientSecret)) {
                throw "Client ID or Secret not configured. Open Settings (gear icon) to configure."
            }
            Write-Log "Connecting to Fly..."
            if (-not (Get-Module -Name Fly.Client -ListAvailable)) { throw "Fly.Client module not found. Run: Install-Module Fly.Client -Scope CurrentUser" }
            Import-Module Fly.Client -ErrorAction Stop
            Connect-Fly -Url $cfg.Url -ClientId $cfg.ClientId -ClientSecret $cfg.ClientSecret -ErrorAction Stop
            Write-Log "Testing API endpoint..."
            try {
                Get-FlyMigrationProject -ErrorAction Stop | Out-Null
                Write-Log "API endpoint reachable OK" "OK"
            } catch {
                $testErr = $_.Exception.Message
                if ($testErr -like "*404*" -or $testErr -like "*Not Found*") {
                    Write-Log "API test returned 404 - check the Fly API URL in Settings." "WARN"
                } else { Write-Log "API test warning: $testErr" "WARN" }
            }
            $dotFly.BackColor = $clrGreen; $btnRun.Enabled = $true
            $lblRunStatus.Text = "Ready"; $lblRunStatus.ForeColor = $clrGreen
            Write-Log "Connected to Fly OK" "OK"
        } catch {
            $dotFly.BackColor = $clrRed; $btnFlyConn.Enabled = $true
            Write-Log "Connection failed: $($_.Exception.Message)" "ERROR"
        }
    })

    $btnRun.Add_Click({
        $btnRun.Enabled = $false
        $prefix     = $tbPrefix.Text.Trim()
        $outFolder  = $tbOutFolder.Text.Trim()
        $createProj = $chkCreateProject.Checked

        if ([string]::IsNullOrWhiteSpace($prefix)) { Write-Log "Customer Prefix is required." "ERROR"; $btnRun.Enabled = $true; return }

        $runCust        = $script:customerList | Where-Object { $_.Prefix -eq $prefix } | Select-Object -First 1
        $accountName    = if ($runCust -and $runCust.AccountName) { $runCust.AccountName } else { $prefix }
        $wlServiceLabel = @{
            SharePoint   = 'SharePoint Online'
            Exchange     = 'Exchange Online'
            OneDrive     = 'OneDrive'
            Teams        = 'Microsoft Teams'
            'Teams Chat' = 'Microsoft Teams Chat'
            Groups       = 'Microsoft 365 Groups'
        }

        $ok = 0; $fail = 0

        $RequiredCsvCols = @{
            SharePoint   = @('Source URL', 'Source object level', 'Destination URL', 'Destination object level')
            Exchange     = @('Source', 'Source type', 'Destination', 'Destination type')
            OneDrive     = @('Source user', 'Destination user')
            Teams        = @('Source Team name', 'Source Team email address', 'Destination Team name', 'Destination Team email address')
            'Teams Chat' = @('Source user', 'Destination user')
            Groups       = @('Source Group name', 'Source Group email address', 'Destination Group name', 'Destination Group email address')
        }

        foreach ($wl in $script:WLControls.Keys) {
            $wlCtrl   = $script:WLControls[$wl]
            if (-not $wlCtrl.Check.Checked) { continue }
            $csvPath   = if ($wlCtrl.CSV.Tag) { $wlCtrl.CSV.Tag } else { $wlCtrl.CSV.Text.Trim() }
            $operation = $wlCtrl.Operation.SelectedItem
            $srcConn   = [string]($script:WorkloadConfig.$wl.Source)
            $dstConn   = "$accountName - $($wlServiceLabel[$wl])"
            $policy    = [string]($script:WorkloadConfig.$wl.Policy)
            $cmds      = $WorkloadDefs[$wl]
            $projName  = "$prefix - $wl"

            if ([string]::IsNullOrWhiteSpace($srcConn)) { $wlCtrl.Dot.ForeColor = $clrRed; Write-Log "$wl - no source connection found. Run Create Connections with 'Source' selected first." "WARN"; $fail++; continue }
            if ($csvPath -eq "No file selected" -or -not (Test-Path $csvPath)) { $wlCtrl.Dot.ForeColor = $clrRed; Write-Log "$wl - no valid CSV selected, skipping" "WARN"; $fail++; continue }

            $reqCols = $RequiredCsvCols[$wl]
            if ($reqCols) {
                try {
                    $csvRows = Import-Csv $csvPath
                    if (-not $csvRows) { $wlCtrl.Dot.ForeColor = $clrRed; Write-Log "$wl CSV has no data rows." "ERROR"; $fail++; continue }
                    $csvHeaders  = ($csvRows | Select-Object -First 1).PSObject.Properties.Name
                    $missingCols = $reqCols | Where-Object { $csvHeaders -notcontains $_ }
                    if ($missingCols) { $wlCtrl.Dot.ForeColor = $clrRed; Write-Log "$wl CSV missing required columns: $($missingCols -join ', ')" "ERROR"; Write-Log "  Required: $($reqCols -join ', ')" "WARN"; $fail++; continue }
                    if ($wl -eq 'Exchange') {
                        $badRows = @($csvRows | Where-Object { [string]::IsNullOrWhiteSpace($_.'Source type') -or [string]::IsNullOrWhiteSpace($_.'Destination type') })
                        if ($badRows.Count -gt 0) { $wlCtrl.Dot.ForeColor = $clrRed; Write-Log "$wl CSV has $($badRows.Count) row(s) with empty 'Source type' or 'Destination type'." "ERROR"; $fail++; continue }
                    }
                    if ($wl -in @('OneDrive', 'Teams Chat')) {
                        $sameRows = @($csvRows | Where-Object { $_.'Source user' -eq $_.'Destination user' -and -not [string]::IsNullOrWhiteSpace($_.'Source user') })
                        if ($sameRows.Count -gt 0) { $wlCtrl.Dot.ForeColor = $clrRed; Write-Log "$wl CSV has $($sameRows.Count) row(s) where Source user equals Destination user." "ERROR"; $fail++; continue }
                    }
                    if ($wl -eq 'Groups') {
                        $sameRows = @($csvRows | Where-Object { $_.'Source Group email address' -eq $_.'Destination Group email address' -and -not [string]::IsNullOrWhiteSpace($_.'Source Group email address') })
                        if ($sameRows.Count -gt 0) { $wlCtrl.Dot.ForeColor = $clrRed; Write-Log "$wl CSV has $($sameRows.Count) row(s) where Source Group email equals Destination Group email." "ERROR"; $fail++; continue }
                    }
                } catch { $wlCtrl.Dot.ForeColor = $clrRed; Write-Log "$wl CSV validation failed: $($_.Exception.Message)" "ERROR"; $fail++; continue }
            }

            $wlCtrl.Dot.ForeColor = $clrAmber
            [System.Windows.Forms.Application]::DoEvents()
            Write-Log "-- $wl [$operation] --"

            try {
                if ($createProj) {
                    $existing = Get-FlyMigrationProject -Name $projName -ErrorAction SilentlyContinue
                    if ($existing) {
                        Write-Log "Project exists: $projName"
                    } elseif ([string]::IsNullOrWhiteSpace($policy)) {
                        Write-Log "$wl - no policy set in workloads.json, cannot create project" "WARN"
                    } else {
                        Write-Log "Creating project: $projName..."
                        try {
                            New-FlyMigrationProject -Name $projName -SourceConnection $srcConn -DestinationConnection $dstConn -Policy $policy -ErrorAction Stop | Out-Null
                            Write-Log "Project created OK" "OK"
                        } catch {
                            Write-Log "Project creation failed: $($_.Exception.Message)" "WARN"
                            $verifyExist = Get-FlyMigrationProject -Name $projName -ErrorAction SilentlyContinue
                            if (-not $verifyExist) { Write-Log "$wl - project does not exist and could not be created, skipping" "ERROR"; $wlCtrl.Dot.ForeColor = $clrRed; $fail++; continue }
                            Write-Log "Project found - continuing despite creation warning" "WARN"
                        }
                    }
                }

                $statusFile   = Join-Path $outFolder ($projName + "_MappingStatus_" + (Get-Date -Format "yyyyMMdd-HHmm") + ".csv")
                $errIdxBefore = $Error.Count
                Write-Log "Importing mappings from: $(Split-Path $csvPath -Leaf)..."
                & $cmds.Import -Project $projName -Path $csvPath -ErrorAction Stop
                Write-Log "Mappings imported OK" "OK"
                Write-Log "Exporting mapping status..."
                & $cmds.Status -Project $projName -OutFile $statusFile -ErrorAction Stop
                Write-Log "Mapping status saved: $(Split-Path $statusFile -Leaf)" "OK"

                if      ($operation -eq "Verification") { Write-Log "Starting verification..."; & $cmds.Verify  -Project $projName -ErrorAction Stop; Write-Log "Verification started OK" "OK" }
                elseif  ($operation -eq "Pre-Scan")     { if ($cmds.PreScan) { Write-Log "Starting pre-scan..."; & $cmds.PreScan -Project $projName -ErrorAction Stop; Write-Log "Pre-scan started OK" "OK" } else { Write-Log "$wl does not support Pre-Scan, skipping" "WARN" } }
                elseif  ($operation -eq "Migration")    { Write-Log "Starting migration..."; & $cmds.Start -Project $projName -ErrorAction Stop; Write-Log "Migration started OK" "OK" }
                elseif  ($operation -eq "Report")       {
                    $reportFile = Join-Path $outFolder ($projName + "_MigrationReport_" + (Get-Date -Format "yyyyMMdd-HHmm") + ".csv")
                    Write-Log "Exporting migration report..."
                    & $cmds.Report -Project $projName -OutFile $reportFile -ErrorAction Stop
                    Write-Log "Report saved: $(Split-Path $reportFile -Leaf)" "OK"
                }
                $wlCtrl.Dot.ForeColor = $clrGreen; $ok++
            } catch {
                $wlCtrl.Dot.ForeColor = $clrRed
                $errMsg = $_.Exception.Message
                if ($errMsg -like "*Additional text encountered after finished reading JSON content*") {
                    Write-Log "$wl failed: API returned an unparseable response (likely 404 from wrong base URL)." "ERROR"
                } elseif ($errMsg -like "*404*" -or $errMsg -like "*Not Found*") {
                    Write-Log "$wl failed: 404 Not Found - check your Fly API URL in Settings." "ERROR"
                } else {
                    Write-Log "$wl failed: $errMsg" "ERROR"
                    $apiLogged   = $false
                    $newErrCount = $Error.Count - $errIdxBefore
                    for ($ei = 0; $ei -lt [Math]::Min($newErrCount, 8); $ei++) {
                        $ed = $Error[$ei].ErrorDetails.Message
                        if ($ed) {
                            try {
                                $j = $ed | ConvertFrom-Json -ErrorAction Stop
                                $apiMsg = if ($j.errorMessage) { $j.errorMessage.Trim() } elseif ($j.ErrorMessage) { $j.ErrorMessage.Trim() } else { $ed.Trim() }
                                Write-Log "  API: $apiMsg" "WARN"
                                if ($apiMsg -like "*already exist*" -or $apiMsg -like "*duplicate*" -or $apiMsg -like "*ProjectMappingDuplicated*") { Write-Log "  Hint: mappings already imported - delete '$projName' in Fly and re-run." "WARN" }
                            } catch { Write-Log "  API: $($ed.Trim())" "WARN" }
                            $apiLogged = $true; break
                        }
                    }
                    if (-not $apiLogged) {
                        if ($wl -eq 'OneDrive' -and $errMsg -like "*(400)*") { Write-Log "  API: Source and destination identity should be different." "WARN" }
                        elseif ($errMsg -like "*(500)*" -or $errMsg -like "*Internal Server Error*") { Write-Log "  Hint: 500 on import often means mappings already exist - delete '$projName' in Fly and re-run." "WARN" }
                    }
                }
                $fail++
            }
        }
        Write-Log "------------------------------------------------"
        Write-Log "$ok workload$(if($ok -ne 1){'s'}) completed - $fail failed." "OK"
        $btnRun.Enabled = $true
    })

    # ── LAUNCH ──────────────────────────────────────────────────────────────────
    Write-Log "Ready - connect to Fly, configure project settings, then run."
    Write-Log "Log file: $($script:LogFile)"

    $script:WorkloadConfig = Read-WorkloadConfig
    Write-Log "Workload config: $script:WorkloadConfigPath"
    foreach ($wl in $WorkloadDefs.Keys) {
        $wlCfg = $script:WorkloadConfig.$wl
        $p = if ($wlCfg) { [string]($wlCfg.Policy) } else { "" }
        $s = if ($wlCfg) { [string]($wlCfg.Source) } else { "" }
        $d = if ($wlCfg) { [string]($wlCfg.Destination) } else { "" }
        if ($p) { Write-Log "  $wl policy: $p" "OK" }      else { Write-Log "  $wl policy: (none - configure in Settings)" "WARN" }
        if ($s) { Write-Log "  $wl source: $s" "OK" }      else { Write-Log "  $wl source: (none - configure in Settings)" "WARN" }
        if ($d) { Write-Log "  $wl destination: $d" "OK" } else { Write-Log "  $wl destination: (none - configure in Settings)" "WARN" }
    }

    $savedCfg = Get-FlyConfig
    if ($savedCfg) {
        Write-Log "Credentials loaded from saved config." "OK"
        Write-Log "Auto-connecting..."
        try {
            if (-not (Get-Module -Name Fly.Client -ListAvailable)) { throw "Fly.Client module not found." }
            Import-Module Fly.Client -ErrorAction Stop
            Connect-Fly -Url $savedCfg.Url -ClientId $savedCfg.ClientId -ClientSecret $savedCfg.ClientSecret -ErrorAction Stop
            $dotFly.BackColor = $clrGreen; $btnRun.Enabled = $true
            $lblRunStatus.Text = "Ready"; $lblRunStatus.ForeColor = $clrGreen
            Write-Log "Auto-connected to Fly OK" "OK"
        } catch {
            Write-Log "Auto-connect failed: $($_.Exception.Message)" "WARN"
            Write-Log "Click Connect to retry manually." "INFO"
        }
    }

    $sharedCfg = Read-SharedConfig
    $script:customerList = @()
    if ($sharedCfg.PSObject.Properties['Customers'] -and $sharedCfg.Customers) {
        $script:customerList = @($sharedCfg.Customers | ForEach-Object {
            [pscustomobject]@{ Prefix = "$($_.Prefix)"; AccountName = "$($_.AccountName)" }
        } | Where-Object { $_.Prefix })
    } elseif ($sharedCfg.PSObject.Properties['Prefixes'] -and $sharedCfg.Prefixes) {
        $script:customerList = @($sharedCfg.Prefixes | Where-Object { $_ } | ForEach-Object {
            [pscustomobject]@{ Prefix = "$_"; AccountName = '' }
        })
    } elseif ($sharedCfg.TenantName) {
        $script:customerList = @([pscustomobject]@{ Prefix = "$($sharedCfg.TenantName)"; AccountName = '' })
    }
    foreach ($cust in $script:customerList) { [void]$tbPrefix.Items.Add($cust.Prefix) }
    if ($tbPrefix.Items.Count -gt 0 -and [string]::IsNullOrWhiteSpace($tbPrefix.Text)) {
        $tbPrefix.Text = $tbPrefix.Items[0]
        Write-Log "Customer Prefix loaded from Settings: $($tbPrefix.Text)" "OK"
    }
    $tbPrefix.Add_SelectedIndexChanged({
        $selPfx = $tbPrefix.SelectedItem
        if (-not $selPfx) { return }
        $cust = $script:customerList | Where-Object { $_.Prefix -eq $selPfx } | Select-Object -First 1
        if ($cust -and $cust.AccountName) {
            $tbConnName.Text = $cust.AccountName
            if ([string]::IsNullOrWhiteSpace($tbConnSearch.Text)) {
                $tbConnSearch.Text = ($cust.AccountName -replace '[^A-Za-z0-9]', '').ToLower()
            }
        }
    }.GetNewClosure())
    if ($sharedCfg.SecretExpiry) {
        try {
            $expiry   = [datetime]$sharedCfg.SecretExpiry
            $daysLeft = ($expiry - (Get-Date)).Days
            if    ($daysLeft -le 0)  { Write-Log "Client secret EXPIRED ($($expiry.ToString('yyyy-MM-dd'))) — renew via Create App Registration." "ERROR" }
            elseif ($daysLeft -le 30) { Write-Log "Client secret expires in $daysLeft day(s) on $($expiry.ToString('yyyy-MM-dd')) — plan renewal soon." "WARN" }
        } catch {}
    }

    $nodeExe = Find-WFNodeExe
    if (-not $nodeExe) {
        $script:lblConnStatus.Text = "Node.js not found - install Node 18+ from nodejs.org and reopen."
    } elseif (-not (Test-Path $script:connectorJsPath)) {
        $script:lblConnStatus.Text = "fly-connector.js not found next to this script."
    } else {
        $authFile = Join-Path $PSScriptRoot 'auth\storageState.json'
        if (Test-Path $authFile) { $script:lblConnStatus.Text = "Previous session found. Sign in again only if expired." }
    }

    [void]$Form.ShowDialog()
}
