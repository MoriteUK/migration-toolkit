# ═════════════════════════════════════════════════════════════════════════════
# MIGRATION REPORTS
# ═════════════════════════════════════════════════════════════════════════════
function Show-ReportingForm {

    # ── Error pattern analysis ─────────────────────────────────────────────
    $errPatterns = @(
        [pscustomobject]@{
            Pattern = [regex]::new('access.?denied|permission.?denied|insufficient.?privil|forbidden|403', 'IgnoreCase')
            Type    = 'Access Denied'
            Fix     = 'Ensure the migration account has Site Collection Admin (SharePoint/OneDrive), Full Access (Exchange), or equivalent rights on source and destination. Re-test the connection after granting permissions.'
        }
        [pscustomobject]@{
            Pattern = [regex]::new('401|unauthorized|token.?expired|invalid.?credential|auth.*fail', 'IgnoreCase')
            Type    = 'Authentication Error'
            Fix     = 'Credentials have expired or are incorrect. Reconnect to the Fly API and re-test before retrying the migration.'
        }
        [pscustomobject]@{
            Pattern = [regex]::new('throttl|429|too.?many.?request|rate.?limit', 'IgnoreCase')
            Type    = 'Throttling'
            Fix     = 'Microsoft is rate-limiting requests. Fly will retry automatically. Consider reducing project concurrency in the Fly portal under Project Settings > Advanced.'
        }
        [pscustomobject]@{
            Pattern = [regex]::new('timeout|timed.?out|connection.?reset|socket|network', 'IgnoreCase')
            Type    = 'Network Timeout'
            Fix     = 'A network interruption occurred during transfer. The item will be retried on the next incremental pass. If persistent, check firewall rules or increase the Fly timeout setting.'
        }
        [pscustomobject]@{
            Pattern = [regex]::new('not.?found|does.?not.?exist|no.?such|404|mailbox.*missing|smtp.*invalid', 'IgnoreCase')
            Type    = 'Item / Mailbox Not Found'
            Fix     = 'The source or destination object no longer exists. Verify the URL, email address, or mailbox in the mapping CSV and ensure the destination is licensed and provisioned.'
        }
        [pscustomobject]@{
            Pattern = [regex]::new('size.?exceed|too.?large|exceeds.?limit|max.?size|file.?too.?big', 'IgnoreCase')
            Type    = 'Item Too Large'
            Fix     = 'The item exceeds the Microsoft size limit (e.g. 250 GB for SharePoint files, 150 MB for Exchange items). Split, compress, or migrate manually outside of Fly.'
        }
        [pscustomobject]@{
            Pattern = [regex]::new('duplicate|already.?exist|conflict', 'IgnoreCase')
            Type    = 'Duplicate / Conflict'
            Fix     = 'An item with the same name exists at the destination. Update the Fly project conflict resolution policy (Overwrite vs. Skip) under Project Settings > Migration Policy.'
        }
        [pscustomobject]@{
            Pattern = [regex]::new('unsupported|not.?support|cannot.?migrat', 'IgnoreCase')
            Type    = 'Unsupported Content'
            Fix     = 'This content type is not supported by Fly. Consult the AvePoint supported content matrix and migrate this item manually if required.'
        }
    )

    function Get-ErrorAnalysis([string]$Message) {
        if ([string]::IsNullOrWhiteSpace($Message)) {
            return [pscustomobject]@{ Type = '—'; Fix = '—' }
        }
        foreach ($p in $errPatterns) {
            if ($p.Pattern.IsMatch($Message)) {
                return [pscustomobject]@{ Type = $p.Type; Fix = $p.Fix }
            }
        }
        return [pscustomobject]@{
            Type = 'Unknown Error'
            Fix  = 'Review the full error message in the log. Check the Fly portal for more detail, or contact AvePoint support if the error persists.'
        }
    }

    function Find-RowField($Row, [string[]]$Candidates) {
        foreach ($c in $Candidates) {
            $prop = $Row.PSObject.Properties[$c]
            if ($prop -and -not [string]::IsNullOrWhiteSpace($prop.Value)) { return $prop.Value }
        }
        foreach ($c in $Candidates) {
            $prop = $Row.PSObject.Properties | Where-Object { $_.Name -ieq $c } | Select-Object -First 1
            if ($prop -and -not [string]::IsNullOrWhiteSpace($prop.Value)) { return $prop.Value }
        }
        return ''
    }

    # ── Report cmdlet map ──────────────────────────────────────────────────
    $rptCmdlets = [ordered]@{
        'SharePoint'  = @{ migration = 'Export-FlySharePointMigrationReport'; mapping = 'Export-FlySharePointMappingStatus'; Display = 'SharePoint Online'    }
        'Exchange'    = @{ migration = 'Export-FlyExchangeMigrationReport';   mapping = 'Export-FlyExchangeMappingStatus';   Display = 'Exchange Online'       }
        'OneDrive'    = @{ migration = 'Export-FlyOneDriveMigrationReport';   mapping = 'Export-FlyOneDriveMappingStatus';   Display = 'OneDrive for Business' }
        'Teams'       = @{ migration = 'Export-FlyTeamsMigrationReport';      mapping = 'Export-FlyTeamsMappingStatus';      Display = 'Microsoft Teams'       }
        'Teams Chat'  = @{ migration = 'Export-FlyTeamChatMigrationReport';   mapping = 'Export-FlyTeamChatMappingStatus';   Display = 'Teams Chat'            }
        'Groups'      = @{ migration = 'Export-FlyM365GroupMigrationReport';  mapping = 'Export-FlyM365GroupMappingStatus';  Display = 'Microsoft 365 Groups'  }
    }

    # ── Build Form ─────────────────────────────────────────────────────────
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text            = "AvePoint Fly - Migration Reports"
    $Form.WindowState     = [System.Windows.Forms.FormWindowState]::Maximized
    $Form.StartPosition   = [System.Windows.Forms.FormStartPosition]::WindowsDefaultBounds
    $Form.BackColor       = $clrBg
    $Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $Form.MinimumSize     = [System.Drawing.Size]::new(1024, 768)
    $Form.Font            = $FontBody
    $_ico = Join-Path $PSScriptRoot 'FlyMigration.ico'; if (Test-Path $_ico) { $Form.Icon = [System.Drawing.Icon]::new($_ico) }

    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Height = 46; $hdr.Dock = [System.Windows.Forms.DockStyle]::Top
    $hdr.BackColor = $clrAccent
    $Form.Controls.Add($hdr)

    $rptFooter = New-Object System.Windows.Forms.Panel
    $rptFooter.Height    = 46
    $rptFooter.Dock      = [System.Windows.Forms.DockStyle]::Bottom
    $rptFooter.BackColor = [System.Drawing.Color]::FromArgb(28, 32, 48)
    $Form.Controls.Add($rptFooter)

    $btnRCreds = New-Object System.Windows.Forms.Button
    $btnRCreds.Text      = "Credentials..."
    $btnRCreds.Location  = [System.Drawing.Point]::new(16, 8)
    $btnRCreds.Size      = [System.Drawing.Size]::new(110, 30)
    $btnRCreds.BackColor = $clrAccent; $btnRCreds.ForeColor = [System.Drawing.Color]::White
    $btnRCreds.Font      = $FontBold; $btnRCreds.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnRCreds.FlatAppearance.BorderSize = 0; $btnRCreds.Cursor = [System.Windows.Forms.Cursors]::Hand
    $rptFooter.Controls.Add($btnRCreds)

    $dotRConn = New-Object System.Windows.Forms.Panel
    $dotRConn.Size      = [System.Drawing.Size]::new(12, 12)
    $dotRConn.Location  = [System.Drawing.Point]::new(134, 17)
    $dotRConn.BackColor = $clrGrey
    $rptFooter.Controls.Add($dotRConn)

    $lblRStatus = New-Object System.Windows.Forms.Label
    $lblRStatus.Location  = [System.Drawing.Point]::new(156, 14)
    $lblRStatus.Size      = [System.Drawing.Size]::new(400, 20)
    $lblRStatus.Font      = $FontBody
    $lblRStatus.ForeColor = [System.Drawing.Color]::FromArgb(190, 210, 255)
    $lblRStatus.Text      = "Connecting..."
    $rptFooter.Controls.Add($lblRStatus)

    $rptBtnClose = New-Object System.Windows.Forms.Button
    $rptBtnClose.Text      = "Close"; $rptBtnClose.Size = [System.Drawing.Size]::new(90, 30)
    $rptBtnClose.Location  = [System.Drawing.Point]::new(900, 8)
    $rptBtnClose.BackColor = [System.Drawing.Color]::FromArgb(200, 55, 55)
    $rptBtnClose.ForeColor = [System.Drawing.Color]::White; $rptBtnClose.Font = $FontBold
    $rptBtnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $rptBtnClose.FlatAppearance.BorderSize = 0
    $rptBtnClose.Cursor    = [System.Windows.Forms.Cursors]::Hand; $rptBtnClose.Add_Click({ $Form.Close() })
    $rptFooter.Controls.Add($rptBtnClose)
    $rptFooter.Add_SizeChanged({
        $rptBtnClose.Left  = $rptFooter.Width - 100
        $lblRStatus.Width  = $rptFooter.Width - 260
    })

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
    $hdrTitle.Text      = "  Migration Reports"
    $hdrTitle.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 12)
    $hdrTitle.ForeColor = [System.Drawing.Color]::White
    $hdrTitle.Location  = [System.Drawing.Point]::new($_hdrX, 0)
    $hdrTitle.Size      = [System.Drawing.Size]::new(400, 46)
    $hdrTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdrTitle.Anchor    = $AnchorTL
    $hdr.Controls.Add($hdrTitle)

    $tlp = New-Object System.Windows.Forms.TableLayoutPanel
    $tlp.Dock        = [System.Windows.Forms.DockStyle]::Fill
    $tlp.ColumnCount = 1
    $tlp.RowCount    = 5
    $tlp.Padding     = New-Object System.Windows.Forms.Padding(8, 6, 8, 6)
    $tlp.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 155))) | Out-Null
    $tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,  48))) | Out-Null
    $tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,  40))) | Out-Null
    $tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,  60))) | Out-Null
    $tlp.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,  40))) | Out-Null
    $Form.Controls.Add($tlp)

    # ── Row 0: Report Options ───────────────────────────────────────────────
    $c2 = New-CardPanel "REPORT OPTIONS"
    $tlp.Controls.Add($c2.Parent, 0, 0)

    New-Lbl $c2 "Customer Prefix" 16 44 | Out-Null
    $tbRPrefix = New-Object System.Windows.Forms.ComboBox
    $tbRPrefix.Location      = [System.Drawing.Point]::new(16, 60)
    $tbRPrefix.Size          = [System.Drawing.Size]::new(240, 24)
    $tbRPrefix.Font          = $FontBody
    $tbRPrefix.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
    $c2.Controls.Add($tbRPrefix)

    New-Lbl $c2 "Report Type" 220 44 | Out-Null
    $rdoMigration = New-Object System.Windows.Forms.RadioButton
    $rdoMigration.Text = "Error Report"; $rdoMigration.Font = $FontBody
    $rdoMigration.Location = [System.Drawing.Point]::new(220, 62); $rdoMigration.AutoSize = $true; $rdoMigration.Checked = $true
    $c2.Controls.Add($rdoMigration)
    $rdoMapping = New-Object System.Windows.Forms.RadioButton
    $rdoMapping.Text = "Mapping Status"; $rdoMapping.Font = $FontBody
    $rdoMapping.Location = [System.Drawing.Point]::new(370, 62); $rdoMapping.AutoSize = $true
    $c2.Controls.Add($rdoMapping)

    New-Lbl $c2 "Workloads" 16 96 | Out-Null
    $script:rptWLChecks = [ordered]@{}
    $rptWLLabels = [ordered]@{
        'SharePoint'  = 'SharePoint Online'
        'Exchange'    = 'Exchange Online'
        'OneDrive'    = 'OneDrive for Business'
        'Teams'       = 'Microsoft Teams'
        'Teams Chat'  = 'Teams Chat'
        'Groups'      = 'Microsoft 365 Groups'
    }
    $xi = 0
    foreach ($wl in $rptWLLabels.Keys) {
        $chk = New-Object System.Windows.Forms.CheckBox
        $chk.Text = $rptWLLabels[$wl]; $chk.Font = $FontBody
        $chk.Location = [System.Drawing.Point]::new(16 + ($xi * 185), 114)
        $chk.AutoSize = $true; $chk.Checked = $true
        $c2.Controls.Add($chk)
        $script:rptWLChecks[$wl] = $chk
        $xi++
    }

    # ── Row 1: Action bar ──────────────────────────────────────────────────
    $c3 = New-CardPanel ""
    $tlp.Controls.Add($c3.Parent, 0, 1)

    $btnRRun    = New-Btn $c3 "Fetch Report" 16 10 140 28
    $btnRRun.Enabled = $false
    $btnRExport = New-Btn $c3 "Export CSV"  168 10 120 28 $false
    $btnRExport.Enabled = $false
    $btnRClear  = New-Btn $c3 "Clear"       300 10  80 28 $false

    # ── Row 2: Summary bar ─────────────────────────────────────────────────
    $c4 = New-CardPanel ""
    $tlp.Controls.Add($c4.Parent, 0, 2)

    $summaryFont = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
    $lblRTotal   = New-Object System.Windows.Forms.Label
    $lblRSuccess = New-Object System.Windows.Forms.Label
    $lblRWarn    = New-Object System.Windows.Forms.Label
    $lblRFailed  = New-Object System.Windows.Forms.Label
    foreach ($lbl in @($lblRTotal, $lblRSuccess, $lblRWarn, $lblRFailed)) {
        $lbl.Font = $summaryFont; $lbl.AutoSize = $true; $lbl.ForeColor = $clrMuted
    }
    $lblRTotal.Location   = [System.Drawing.Point]::new(16,  10)
    $lblRSuccess.Location = [System.Drawing.Point]::new(160, 10)
    $lblRWarn.Location    = [System.Drawing.Point]::new(320, 10)
    $lblRFailed.Location  = [System.Drawing.Point]::new(480, 10)
    $lblRTotal.Text = "Total: 0"; $lblRSuccess.Text = "Succeeded: 0"
    $lblRWarn.Text  = "Warnings: 0"; $lblRFailed.Text  = "Failed: 0"
    $c4.Controls.AddRange(@($lblRTotal, $lblRSuccess, $lblRWarn, $lblRFailed))

    $script:rptStats = @{ Total = 0; Success = 0; Warn = 0; Failed = 0 }

    function Update-RptStats {
        $lblRTotal.Text        = "Total: $($script:rptStats.Total)";     $lblRTotal.ForeColor   = $clrText
        $lblRSuccess.Text      = "Succeeded: $($script:rptStats.Success)"
        $lblRSuccess.ForeColor = if ($script:rptStats.Success -gt 0) { $clrGreen } else { $clrMuted }
        $lblRWarn.Text         = "Warnings: $($script:rptStats.Warn)"
        $lblRWarn.ForeColor    = if ($script:rptStats.Warn -gt 0) { $clrAmber } else { $clrMuted }
        $lblRFailed.Text       = "Failed: $($script:rptStats.Failed)"
        $lblRFailed.ForeColor  = if ($script:rptStats.Failed -gt 0) { $clrRed } else { $clrMuted }
        [System.Windows.Forms.Application]::DoEvents()
    }

    # ── Row 3: Error Analysis DataGridView ─────────────────────────────────
    $c5 = New-CardPanel "ERROR ANALYSIS"
    $tlp.Controls.Add($c5.Parent, 0, 3)

    $c5TLP = New-Object System.Windows.Forms.TableLayoutPanel
    $c5TLP.Dock = [System.Windows.Forms.DockStyle]::Fill
    $c5TLP.BackColor = [System.Drawing.Color]::White
    $c5TLP.ColumnCount = 1; $c5TLP.RowCount = 2
    $c5TLP.Padding = New-Object System.Windows.Forms.Padding(12, 0, 6, 4)
    [void]$c5TLP.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$c5TLP.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
    [void]$c5TLP.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $c5.Controls.Add($c5TLP)

    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.Dock                      = [System.Windows.Forms.DockStyle]::Fill
    $dgv.AutoGenerateColumns       = $false
    $dgv.ReadOnly                  = $true
    $dgv.AllowUserToAddRows        = $false
    $dgv.AllowUserToDeleteRows     = $false
    $dgv.BackgroundColor           = [System.Drawing.Color]::White
    $dgv.BorderStyle               = [System.Windows.Forms.BorderStyle]::None
    $dgv.ColumnHeadersBorderStyle  = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::Single
    $dgv.ColumnHeadersDefaultCellStyle.BackColor = $clrBg
    $dgv.ColumnHeadersDefaultCellStyle.ForeColor = $clrMuted
    $dgv.ColumnHeadersDefaultCellStyle.Font      = $FontCap
    $dgv.CellBorderStyle           = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $dgv.GridColor                 = $clrBorder
    $dgv.RowHeadersVisible         = $false
    $dgv.SelectionMode             = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $dgv.Font                      = New-Object System.Drawing.Font("Segoe UI", 9)
    $dgv.AutoSizeRowsMode          = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells
    $dgv.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
    $dgv.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 252)

    foreach ($cd in @(
        @{ Name='ColWorkload'; Header='Workload';        Width=140; Fill=$false }
        @{ Name='ColSource';   Header='Source Object';   Width=200; Fill=$false }
        @{ Name='ColStatus';   Header='Status';          Width=100; Fill=$false }
        @{ Name='ColErrType';  Header='Error Type';      Width=160; Fill=$false }
        @{ Name='ColErrMsg';   Header='Error Message';   Width=220; Fill=$false }
        @{ Name='ColFix';      Header='Recommended Fix'; Width=0;   Fill=$true  }
    )) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $cd.Name; $col.HeaderText = $cd.Header
        $col.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
        if ($cd.Fill) { $col.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill }
        else          { $col.Width = $cd.Width }
        $dgv.Columns.Add($col) | Out-Null
    }

    $dgv.Add_CellFormatting({
        param($s, $e)
        if ($e.RowIndex -lt 0 -or $e.ColumnIndex -ne 2) { return }
        $val = $dgv.Rows[$e.RowIndex].Cells['ColStatus'].Value
        if (-not $val) { return }
        switch -Regex ($val) {
            '^[Ff]ail|^[Ee]rror' { $e.CellStyle.ForeColor = $clrRed;   $e.CellStyle.Font = $FontBold; $e.FormattingApplied = $true }
            '^[Ww]arn|^[Ss]kip'  { $e.CellStyle.ForeColor = $clrAmber; $e.CellStyle.Font = $FontBold; $e.FormattingApplied = $true }
        }
    })
    $c5TLP.Controls.Add($dgv, 0, 1)

    # ── Row 4: Log ─────────────────────────────────────────────────────────
    $c6 = New-CardPanel "LOG"
    $tlp.Controls.Add($c6.Parent, 0, 4)

    $c6TLP = New-Object System.Windows.Forms.TableLayoutPanel
    $c6TLP.Dock = [System.Windows.Forms.DockStyle]::Fill
    $c6TLP.BackColor = [System.Drawing.Color]::White
    $c6TLP.ColumnCount = 1; $c6TLP.RowCount = 2
    $c6TLP.Padding = New-Object System.Windows.Forms.Padding(4, 0, 4, 4)
    [void]$c6TLP.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$c6TLP.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
    [void]$c6TLP.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $c6.Controls.Add($c6TLP)

    $rptLog = New-Object System.Windows.Forms.RichTextBox
    $rptLog.Dock        = [System.Windows.Forms.DockStyle]::Fill
    $rptLog.Font        = $FontMono; $rptLog.BackColor = $clrLogBg
    $rptLog.ForeColor   = [System.Drawing.Color]::FromArgb(190, 210, 255)
    $rptLog.ReadOnly    = $true; $rptLog.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $rptLog.ScrollBars  = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    $c6TLP.Controls.Add($rptLog, 0, 1)

    function Write-Log {
        param([string]$Msg, [string]$Level = "INFO")
        $ts = Get-Date -Format "HH:mm:ss"
        $rptLog.SelectionStart = $rptLog.TextLength; $rptLog.SelectionLength = 0
        $rptLog.SelectionColor = [System.Drawing.Color]::FromArgb(80, 95, 120)
        $rptLog.AppendText("$ts ")
        $rptLog.SelectionColor = switch ($Level) {
            "OK"    { [System.Drawing.Color]::FromArgb(65,  195, 110) }
            "WARN"  { [System.Drawing.Color]::FromArgb(220, 165, 45)  }
            "ERROR" { [System.Drawing.Color]::FromArgb(225, 80,  80)  }
            default { [System.Drawing.Color]::FromArgb(120, 155, 220) }
        }
        $rptLog.AppendText("[$Level] ")
        $rptLog.SelectionColor = [System.Drawing.Color]::FromArgb(205, 212, 230)
        $rptLog.AppendText("$Msg`n")
        $rptLog.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }

    $script:rptExportRows = [System.Collections.Generic.List[object]]::new()

    # ── Config helpers ─────────────────────────────────────────────────────
    $rptCfgPath = Join-Path $env:APPDATA "FlyMigration\config.json"

    function Read-RptConfig {
        if (-not (Test-Path $rptCfgPath)) { return $null }
        try {
            $cfg    = Get-Content $rptCfgPath -Raw | ConvertFrom-Json
            $secure = $cfg.EncSecret | ConvertTo-SecureString
            $bstr   = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
            $plain  = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
            return @{ Url = $cfg.Url; ClientId = $cfg.ClientId; ClientSecret = $plain }
        } catch { return $null }
    }

    function Save-RptConfig([string]$Url, [string]$ClientId, [string]$ClientSecret) {
        $dir = Split-Path $rptCfgPath
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
        $enc = $ClientSecret | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
        @{ Url = $Url; ClientId = $ClientId; EncSecret = $enc } |
            ConvertTo-Json | Set-Content -Path $rptCfgPath -Encoding UTF8
    }

    function Connect-RptFly([hashtable]$Cfg) {
        $dotRConn.BackColor   = $clrGrey
        $lblRStatus.Text      = "Connecting..."; $lblRStatus.ForeColor = $clrMuted
        [System.Windows.Forms.Application]::DoEvents()
        try {
            if (-not (Get-Module -Name Fly.Client -ListAvailable)) {
                throw "Fly.Client module not found. Run: Install-Module Fly.Client -Scope CurrentUser"
            }
            Import-Module Fly.Client -ErrorAction Stop
            Connect-Fly -Url $Cfg.Url -ClientId $Cfg.ClientId -ClientSecret $Cfg.ClientSecret -ErrorAction Stop
            Write-Log "Connected to Fly API." "OK"
            $dotRConn.BackColor   = $clrGreen
            $btnRRun.Enabled      = $true
            $lblRStatus.Text      = "Ready"; $lblRStatus.ForeColor = $clrGreen
        } catch {
            $dotRConn.BackColor   = $clrRed
            $btnRRun.Enabled      = $false
            $lblRStatus.Text      = "Not connected"; $lblRStatus.ForeColor = $clrRed
            Write-Log "Connection failed: $($_.Exception.Message)" "ERROR"
        }
        [System.Windows.Forms.Application]::DoEvents()
    }

    # ── Credentials dialog ─────────────────────────────────────────────────
    $btnRCreds.Add_Click({
        $existing = Read-RptConfig
        $cdlg = New-Object System.Windows.Forms.Form
        $cdlg.Text            = "Fly API Credentials"
        $cdlg.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterParent
        $cdlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $cdlg.MaximizeBox     = $false; $cdlg.MinimizeBox = $false
        $cdlg.BackColor       = $clrBg; $cdlg.Font = $FontBody
        $cdlg.ClientSize      = [System.Drawing.Size]::new(520, 220)

        $mkLbl = { param($t,$x,$y)
            $l = New-Object System.Windows.Forms.Label
            $l.Text = $t; $l.Location = [System.Drawing.Point]::new($x,$y)
            $l.AutoSize = $true; $l.ForeColor = $clrMuted; $l.Font = $FontCap
            $cdlg.Controls.Add($l)
        }
        $mkTb = { param($x,$y,$w,[bool]$isPass=$false)
            $tb = New-Object System.Windows.Forms.TextBox
            $tb.Location = [System.Drawing.Point]::new($x,$y)
            $tb.Size     = [System.Drawing.Size]::new($w,26)
            $tb.BackColor = [System.Drawing.Color]::White
            $tb.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
            if ($isPass) { $tb.UseSystemPasswordChar = $true }
            $cdlg.Controls.Add($tb); $tb
        }

        & $mkLbl "FLY API URL"   16 16
        $cdTbUrl = & $mkTb 16 32 488
        & $mkLbl "AOS CLIENT ID" 16 72
        $cdTbCid = & $mkTb 16 88 230
        & $mkLbl "CLIENT SECRET" 260 72
        $cdTbSec = & $mkTb 260 88 244 $true

        if ($existing) {
            $cdTbUrl.Text = $existing.Url
            $cdTbCid.Text = $existing.ClientId
            $cdTbSec.Text = $existing.ClientSecret
        }

        $btnSave = New-Object System.Windows.Forms.Button
        $btnSave.Text = "Save & Connect"; $btnSave.Size = [System.Drawing.Size]::new(130,30)
        $btnSave.Location  = [System.Drawing.Point]::new(16,168)
        $btnSave.BackColor = $clrAccent; $btnSave.ForeColor = [System.Drawing.Color]::White
        $btnSave.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $cdlg.Controls.Add($btnSave)

        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Text = "Cancel"; $btnCancel.Size = [System.Drawing.Size]::new(80,30)
        $btnCancel.Location  = [System.Drawing.Point]::new(154,168)
        $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnCancel.BackColor = $clrPanel; $btnCancel.ForeColor = $clrText
        $cdlg.Controls.Add($btnCancel)

        $btnSave.Add_Click({
            $url = $cdTbUrl.Text.Trim(); $cid = $cdTbCid.Text.Trim(); $sec = $cdTbSec.Text.Trim()
            if (-not $cid -or -not $sec) {
                [System.Windows.Forms.MessageBox]::Show("Client ID and Secret are required.", "Validation") | Out-Null
                return
            }
            Save-RptConfig -Url $url -ClientId $cid -ClientSecret $sec
            $cdlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $cdlg.Close()
        }.GetNewClosure())

        $btnCancel.Add_Click({ $cdlg.Close() })
        $cdlg.CancelButton = $btnCancel

        if ($cdlg.ShowDialog($Form) -eq [System.Windows.Forms.DialogResult]::OK) {
            $cfg = Read-RptConfig
            if ($cfg) { Connect-RptFly -Cfg $cfg }
        }
    })

    # ── Reports cache directory ────────────────────────────────────────────────
    $script:rptDir = Join-Path $PSScriptRoot 'reports'
    if (-not (Test-Path $script:rptDir)) { New-Item -ItemType Directory -Path $script:rptDir -Force | Out-Null }

    # Workload → Fly project name suffix map (defaults to the workload key; overridden by workloads.json ProjectSuffix)
    $script:rptWLSuffix = [ordered]@{
        'SharePoint' = 'SharePoint'; 'Exchange' = 'Exchange'; 'OneDrive' = 'OneDrive'
        'Teams' = 'Teams'; 'Teams Chat' = 'Teams Chat'; 'Groups' = 'Groups'
    }
    $wlSuffixSrc = Join-Path $PSScriptRoot 'workloads.json'
    if (Test-Path $wlSuffixSrc) {
        try {
            $wlRaw = Get-Content $wlSuffixSrc -Raw | ConvertFrom-Json
            foreach ($wl in @($script:rptWLSuffix.Keys)) {
                if ($wlRaw.PSObject.Properties[$wl] -and $wlRaw.$wl.PSObject.Properties['ProjectSuffix'] -and $wlRaw.$wl.ProjectSuffix) {
                    $script:rptWLSuffix[$wl] = [string]$wlRaw.$wl.ProjectSuffix
                }
            }
        } catch {}
    }

    function Get-CachedReport {
        param([string]$Prefix, [string]$Workload, [string]$ReportType)
        $cutoff   = (Get-Date).AddDays(-7)
        $safeWL   = [regex]::Escape($Workload)
        $safePfx  = if ($Prefix) { [regex]::Escape($Prefix) } else { $null }
        Get-ChildItem $script:rptDir -ErrorAction SilentlyContinue |
            Where-Object { (-not $safePfx -or $_.Name -imatch $safePfx) -and $_.Name -imatch $safeWL -and $_.Name -imatch $ReportType } |
            Where-Object { $_.LastWriteTime -ge $cutoff } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1
    }

    function Show-ReportSourceDialog {
        param([string[]]$Workloads, [string]$Prefix, [string]$ReportType, [hashtable]$CachedFiles)
        $Workloads = @($Workloads)

        $choices = [ordered]@{}
        foreach ($wl in $Workloads) {
            $choices[$wl] = [pscustomobject]@{
                Mode = if ($CachedFiles[$wl]) { 'cached' } else { 'fresh' }
                File = if ($CachedFiles[$wl]) { $CachedFiles[$wl].FullName } else { $null }
            }
        }

        $rowH   = 56
        $bodyH  = $Workloads.Count * $rowH + 32
        $dlgH   = 44 + 40 + $bodyH + 50   # header + quick-bar + body + footer

        $dlg = New-Object System.Windows.Forms.Form
        $dlg.Text            = "Select Report Sources"
        $dlg.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $dlg.MaximizeBox     = $false
        $dlg.MinimizeBox     = $false
        $dlg.BackColor       = $clrBg
        $dlg.Font            = $FontBody
        $dlg.ClientSize      = [System.Drawing.Size]::new(740, $dlgH)

        # Header bar
        $hdr = New-Object System.Windows.Forms.Panel
        $hdr.Size = [System.Drawing.Size]::new(740, 44); $hdr.BackColor = $clrAccent
        $dlg.Controls.Add($hdr)
        $hdrLbl = New-Object System.Windows.Forms.Label
        $hdrLbl.Text      = "  Select Report Sources  —  $Prefix  ($ReportType)"
        $hdrLbl.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
        $hdrLbl.ForeColor = [System.Drawing.Color]::White
        $hdrLbl.Location  = [System.Drawing.Point]::new(8, 13)
        $hdrLbl.AutoSize  = $true
        $hdr.Controls.Add($hdrLbl)

        # Quick-select buttons
        $btnAllCached = New-Btn $dlg "Use All Cached"     10 52 136 28 $false
        $btnAllFresh  = New-Btn $dlg "Generate All Fresh" 154 52 150 28 $false
        $btnAllCached.Enabled = [bool]($CachedFiles.Values | Where-Object { $_ })

        # Column header labels
        $colX = @(10, 164, 390, 580)
        foreach ($i in 0..3) {
            $lh = New-Object System.Windows.Forms.Label
            $lh.Text      = @('Workload', 'Cached report  (< 2 days old)', 'Source', 'Browse')[($i)]
            $lh.Font      = $FontCap; $lh.ForeColor = $clrMuted
            $lh.Location  = [System.Drawing.Point]::new($colX[$i], 90)
            $lh.AutoSize  = $true
            $dlg.Controls.Add($lh)
        }

        $rdoGroups = [ordered]@{}
        $y = 106

        foreach ($wl in $Workloads) {
            $cached = $CachedFiles[$wl]

            # Workload label
            $wlLbl = New-Object System.Windows.Forms.Label
            $wlLbl.Text = $wl; $wlLbl.Font = $FontBold; $wlLbl.ForeColor = $clrText
            $wlLbl.Location = [System.Drawing.Point]::new($colX[0], $y + 8); $wlLbl.AutoSize = $true
            $dlg.Controls.Add($wlLbl)

            # Cached file info
            $cacheLbl = New-Object System.Windows.Forms.Label
            $cacheLbl.Location = [System.Drawing.Point]::new($colX[1], $y)
            $cacheLbl.Size     = [System.Drawing.Size]::new(218, 48)
            $cacheLbl.Font     = $FontBody
            if ($cached) {
                $hrs    = ((Get-Date) - $cached.LastWriteTime).TotalHours
                $ageStr = if ($hrs -lt 1) { "$([int]($hrs*60))m ago" } `
                          elseif ($hrs -lt 24) { "$([Math]::Round($hrs,1))h ago" } `
                          else { "$([Math]::Round($hrs/24,1))d ago" }
                $cacheLbl.Text      = $cached.Name + "`n" + $ageStr
                $cacheLbl.ForeColor = $clrGreen
            } else {
                $cacheLbl.Text      = "None found"
                $cacheLbl.ForeColor = $clrMuted
            }
            $dlg.Controls.Add($cacheLbl)

            # Radio panel (groups the three radios so they auto-mutually-exclude per workload)
            $rdoPnl = New-Object System.Windows.Forms.Panel
            $rdoPnl.Location  = [System.Drawing.Point]::new($colX[2], $y)
            $rdoPnl.Size      = [System.Drawing.Size]::new(340, 52)
            $rdoPnl.BackColor = $clrBg
            $dlg.Controls.Add($rdoPnl)

            $rdoCached = New-Object System.Windows.Forms.RadioButton
            $rdoCached.Text = "Use cached"; $rdoCached.Font = $FontBody
            $rdoCached.Location = [System.Drawing.Point]::new(0, 2); $rdoCached.AutoSize = $true
            $rdoCached.Enabled  = [bool]$cached; $rdoCached.Checked = [bool]$cached
            $rdoPnl.Controls.Add($rdoCached)

            $rdoFresh = New-Object System.Windows.Forms.RadioButton
            $rdoFresh.Text = "Generate fresh"; $rdoFresh.Font = $FontBody
            $rdoFresh.Location = [System.Drawing.Point]::new(0, 26); $rdoFresh.AutoSize = $true
            $rdoFresh.Checked  = -not $cached
            $rdoPnl.Controls.Add($rdoFresh)

            # Browse radio + label live outside the panel (col 3)
            $rdoBrowse_rdo = New-Object System.Windows.Forms.RadioButton
            $rdoBrowse_rdo.Text = "Browse for file..."; $rdoBrowse_rdo.Font = $FontBody
            $rdoBrowse_rdo.Location = [System.Drawing.Point]::new(0, 0); $rdoBrowse_rdo.AutoSize = $true
            $rdoPnl.Controls.Add($rdoBrowse_rdo)
            $rdoBrowse_rdo.Location = [System.Drawing.Point]::new($colX[3] - $colX[2], 2)

            $brwLbl = New-Object System.Windows.Forms.Label
            $brwLbl.Text      = "(no file chosen)"
            $brwLbl.Font      = New-Object System.Drawing.Font("Segoe UI", 7.5)
            $brwLbl.ForeColor = $clrMuted
            $brwLbl.Location  = [System.Drawing.Point]::new($colX[3] - $colX[2], 26)
            $brwLbl.Size      = [System.Drawing.Size]::new(148, 18)
            $rdoPnl.Controls.Add($brwLbl)

            # Separator
            $sep = New-Object System.Windows.Forms.Label
            $sep.Location  = [System.Drawing.Point]::new(10, $y + $rowH - 2)
            $sep.Size      = [System.Drawing.Size]::new(720, 1)
            $sep.BackColor = $clrBorder
            $dlg.Controls.Add($sep)

            $rdoGroups[$wl] = @{ Cached = $rdoCached; Fresh = $rdoFresh; Browse = $rdoBrowse_rdo }

            # Wire events with closures
            $captWl     = $wl
            $captCached = $cached
            $captRdoC   = $rdoCached
            $captRdoF   = $rdoFresh
            $captRdoB   = $rdoBrowse_rdo
            $captBrwLbl = $brwLbl

            $rdoCached.Add_CheckedChanged({
                if ($captRdoC.Checked) { $choices[$captWl].Mode = 'cached'; $choices[$captWl].File = $captCached.FullName }
            }.GetNewClosure())

            $rdoFresh.Add_CheckedChanged({
                if ($captRdoF.Checked) { $choices[$captWl].Mode = 'fresh'; $choices[$captWl].File = $null }
            }.GetNewClosure())

            $rdoBrowse_rdo.Add_CheckedChanged({
                if (-not $captRdoB.Checked) { return }
                $ofd = New-Object System.Windows.Forms.OpenFileDialog
                $ofd.Filter = "CSV / report files (*.csv)|*.csv|All files (*.*)|*.*"
                $ofd.Title  = "Select report file for $captWl"
                if ($ofd.ShowDialog() -eq 'OK') {
                    $choices[$captWl].Mode = 'browse'
                    $choices[$captWl].File = $ofd.FileName
                    $captBrwLbl.Text       = [System.IO.Path]::GetFileName($ofd.FileName)
                    $captBrwLbl.ForeColor  = $clrText
                } else {
                    if ($captCached) { $captRdoC.Checked = $true } else { $captRdoF.Checked = $true }
                }
            }.GetNewClosure())

            $y += $rowH
        }

        # Quick-select handlers
        $btnAllCached.Add_Click({
            foreach ($wl in $Workloads) { if ($rdoGroups[$wl].Cached.Enabled) { $rdoGroups[$wl].Cached.Checked = $true } }
        }.GetNewClosure())
        $btnAllFresh.Add_Click({
            foreach ($wl in $Workloads) { $rdoGroups[$wl].Fresh.Checked = $true }
        }.GetNewClosure())

        # Footer
        $btnOk     = New-Btn $dlg "Proceed" 474 ($dlg.ClientSize.Height - 42) 124 30
        $btnCancel = New-Btn $dlg "Cancel"  606 ($dlg.ClientSize.Height - 42) 124 30 $false
        $btnOk.Add_Click({     $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK;     $dlg.Close() })
        $btnCancel.Add_Click({ $dlg.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $dlg.Close() })
        $dlg.AcceptButton = $btnOk; $dlg.CancelButton = $btnCancel

        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $choices }
        return $null
    }

    # ── Fetch Report ───────────────────────────────────────────────────────
    $btnRRun.Add_Click({
        $prefix     = $tbRPrefix.Text.Trim()
        $reportType = if ($rdoMigration.Checked) { 'error' } else { 'mapping' }

        if ([string]::IsNullOrWhiteSpace($prefix)) { Write-Log "Customer Prefix is required." "ERROR"; return }

        $selected = [string[]]($script:rptWLChecks.Keys | Where-Object { $script:rptWLChecks[$_].Checked })
        if (-not $selected) { Write-Log "Select at least one workload." "ERROR"; return }

        if ($reportType -eq 'error') {
            # ── Combined error report (one API call for all selected workloads) ──
            $errKey      = "Combined Error Report"
            $cachedFiles = [ordered]@{ $errKey = Get-CachedReport -Prefix $prefix -Workload 'ErrorReport' -ReportType 'error' }

            $choices = Show-ReportSourceDialog -Workloads @($errKey) -Prefix $prefix -ReportType 'error report' -CachedFiles $cachedFiles
            if (-not $choices) { return }

            $btnRRun.Enabled = $false; $btnRExport.Enabled = $false
            $dgv.Rows.Clear(); $script:rptExportRows.Clear()
            $script:rptStats = @{ Total = 0; Success = 0; Warn = 0; Failed = 0 }
            Update-RptStats
            $lblRStatus.Text = "Loading error report..."; $lblRStatus.ForeColor = $clrMuted

            $choice  = $choices[$errKey]
            $tempCsv = $null
            $tempDir = $null

            try {
                switch ($choice.Mode) {
                    'cached' {
                        $tempCsv = $choice.File
                        Write-Log "Using cached error report: $(Split-Path $tempCsv -Leaf)" "OK"
                    }
                    'browse' {
                        $tempCsv = $choice.File
                        Write-Log "Using selected file: $(Split-Path $tempCsv -Leaf)" "OK"
                    }
                    'fresh' {
                        $projNames = @($selected | ForEach-Object { "$prefix - $($script:rptWLSuffix[$_])" })
                        Write-Log "Generating combined error report for: $($projNames -join ', ')..."
                        $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ([System.IO.Path]::GetRandomFileName())
                        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
                        Export-FlyErrorReport -Projects $projNames -OutFolder $tempDir -ErrorAction Stop | Out-Null
                        $dl = Get-ChildItem $tempDir -Recurse -ErrorAction SilentlyContinue |
                              Where-Object { -not $_.PSIsContainer } | Select-Object -First 1
                        if ($dl) {
                            $saveName = "$($prefix -replace '[^\w ]') ErrorReport $(Get-Date -Format 'yyyyMMdd-HHmm').csv"
                            $savePath = Join-Path $script:rptDir $saveName
                            Copy-Item $dl.FullName $savePath -Force -ErrorAction SilentlyContinue
                            $tempCsv = $savePath
                            Write-Log "Saved to cache: $saveName" "OK"
                        }
                    }
                }

                if (-not $tempCsv -or -not (Test-Path $tempCsv)) {
                    Write-Log "No error report data returned. The selected projects may have no errors, or no migration has completed yet. If a report is visible in the Fly portal Download Centre, use 'Browse' to load it directly." "WARN"
                } else {
                    $rows = @(Import-Csv -Path $tempCsv -ErrorAction Stop)
                    Write-Log "$($rows.Count) rows in error report." "OK"
                    if ($rows.Count -gt 0) { Write-Log "CSV columns: $($rows[0].PSObject.Properties.Name -join ', ')" "INFO" }

                    foreach ($row in $rows) {
                        $wlName  = Find-RowField $row @('WorkloadType','Workload','Workload Type','Platform','Module','Migration Module','Project Name','ProjectName','Project')
                        $status  = Find-RowField $row @('Status','Migration Status','MigrationStatus','Result','State')
                        $source  = Find-RowField $row @('Source Object','SourceObject','Source Identity','SourceIdentity','SourceItem','SourceUser','SourceSite','SourceMailbox','Source','Name','Item')
                        $errMsg  = Find-RowField $row @('Error Message','ErrorMessage','Error Detail','Error Details','ErrorDetail','Error','Message','Failure Reason','FailReason','Description','Details')
                        $errType = Find-RowField $row @('Error Type','ErrorType','Error Category','ErrorCategory','Category','Failure Type','FailureType','Type')
                        $fix     = Find-RowField $row @('Recommended Solution','RecommendedSolution','Recommended Fix','RecommendedFix','Suggested Action','SuggestedAction','Resolution','Fix')
                        $display = if ($wlName) { $wlName } else { $prefix }

                        $script:rptStats.Total++
                        $isSuccess = $status -imatch '^success$|^completed$|^done$'
                        $isWarn    = $status -imatch '^warning|^skipped|^partial'
                        $isFail    = $status -imatch '^fail|^error'

                        if     ($isSuccess) { $script:rptStats.Success++ }
                        elseif ($isWarn)    { $script:rptStats.Warn++    }
                        elseif ($isFail)    { $script:rptStats.Failed++  }
                        elseif ($errMsg)    { $script:rptStats.Failed++  }
                        else                { $script:rptStats.Success++ }

                        if ($isSuccess -and -not $errMsg -and -not $isFail) { Update-RptStats; continue }
                        if (-not $errMsg -and -not $errType -and -not $isFail -and -not $isWarn) { Update-RptStats; continue }

                        $analysis   = Get-ErrorAnalysis $errMsg
                        $srcDisplay = if ($source)  { $source }                                   else { '—' }
                        $stDisplay  = if ($status)  { $status }                                   else { '—' }
                        $errDisplay = if ($errMsg)  { $errMsg }                                   else { '—' }
                        $typeDisplay= if ($errType) { $errType } elseif ($analysis.Type) { $analysis.Type } else { '—' }
                        $fixDisplay = if ($fix)     { $fix }     elseif ($analysis.Fix)  { $analysis.Fix }  else { '—' }

                        $dgv.Rows.Add($display, $srcDisplay, $stDisplay, $typeDisplay, $errDisplay, $fixDisplay) | Out-Null
                        $script:rptExportRows.Add([pscustomobject]@{
                            Workload       = $display
                            SourceObject   = $srcDisplay
                            Status         = $stDisplay
                            ErrorType      = $typeDisplay
                            ErrorMessage   = $errDisplay
                            RecommendedFix = $fixDisplay
                        }) | Out-Null

                        Update-RptStats
                        [System.Windows.Forms.Application]::DoEvents()
                    }
                    Update-RptStats
                }
            } catch {
                Write-Log "Error: $($_.Exception.Message)" "ERROR"
            } finally {
                if ($tempDir -and (Test-Path $tempDir)) { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
            }

        } else {
            # ── Mapping status path (per-workload) ─────────────────────────
            $cachedFiles = [ordered]@{}
            foreach ($wl in $selected) { $cachedFiles[$wl] = Get-CachedReport -Prefix $prefix -Workload $wl -ReportType $reportType }

            $choices = Show-ReportSourceDialog -Workloads $selected -Prefix $prefix -ReportType $reportType -CachedFiles $cachedFiles
            if (-not $choices) { return }

            $btnRRun.Enabled = $false; $btnRExport.Enabled = $false
            $dgv.Rows.Clear(); $script:rptExportRows.Clear()
            $script:rptStats = @{ Total = 0; Success = 0; Warn = 0; Failed = 0 }
            Update-RptStats
            $lblRStatus.Text = "Loading reports..."; $lblRStatus.ForeColor = $clrMuted

            foreach ($wl in $selected) {
                $cmds     = $rptCmdlets[$wl]
                $cmdName  = $cmds[$reportType]
                $display  = $cmds['Display']
                $projName = "$prefix - $($script:rptWLSuffix[$wl])"
                $choice   = $choices[$wl]

                if (-not $cmdName -and $choice.Mode -eq 'fresh') {
                    Write-Log "[$wl] No $reportType cmdlet defined — skipped." "WARN"; continue
                }

                $tempCsv = $null

                try {
                    switch ($choice.Mode) {
                        'cached' {
                            $tempCsv = $choice.File
                            Write-Log "[$wl] Using cached report: $(Split-Path $tempCsv -Leaf)" "OK"
                        }
                        'browse' {
                            $tempCsv = $choice.File
                            Write-Log "[$wl] Using selected file: $(Split-Path $tempCsv -Leaf)" "OK"
                        }
                        'fresh' {
                            Write-Log "[$wl] Generating $reportType report for '$projName'..."
                            $tempCsv = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.csv'
                            & $cmdName -Project $projName -OutFile $tempCsv -ErrorAction Stop | Out-Null
                            if (Test-Path $tempCsv) {
                                $saveName = "$($projName -replace '[^\w ]') $reportType $(Get-Date -Format 'yyyyMMdd-HHmm').csv"
                                $savePath = Join-Path $script:rptDir $saveName
                                Copy-Item $tempCsv $savePath -Force -ErrorAction SilentlyContinue
                                Write-Log "[$wl] Saved to cache: $saveName" "OK"
                            }
                        }
                    }

                    if (-not $tempCsv -or -not (Test-Path $tempCsv)) {
                        Write-Log "[$wl] No report data returned — the project may have no data yet. If a report is visible in the Fly portal Download Centre, use 'Browse' to load it directly." "WARN"; continue
                    }

                    $rows = @(Import-Csv -Path $tempCsv -ErrorAction Stop)
                    Write-Log "[$wl] $($rows.Count) rows retrieved." "OK"

                    foreach ($row in $rows) {
                        $status  = Find-RowField $row @('Status','Migration Status','MigrationStatus','Result','State')
                        $source  = Find-RowField $row @('Source Object','SourceObject','Source Identity','SourceIdentity','SourceItem','SourceUser','SourceSite','SourceMailbox','Source','Name','Item')
                        $errMsg  = Find-RowField $row @('Error Message','ErrorMessage','Error Detail','Error Details','ErrorDetail','Error','Message','Failure Reason','FailReason','Description','Details')
                        $errType = Find-RowField $row @('Error Type','ErrorType','Error Category','ErrorCategory','Category','Failure Type','FailureType','Type')
                        $fix     = Find-RowField $row @('Recommended Solution','RecommendedSolution','Recommended Fix','RecommendedFix','Suggested Action','SuggestedAction','Resolution','Fix')

                        $script:rptStats.Total++
                        $isSuccess = $status -imatch '^success$|^completed$|^done$'
                        $isWarn    = $status -imatch '^warning|^skipped|^partial'
                        $isFail    = $status -imatch '^fail|^error'

                        if     ($isSuccess) { $script:rptStats.Success++ }
                        elseif ($isWarn)    { $script:rptStats.Warn++    }
                        elseif ($isFail)    { $script:rptStats.Failed++  }
                        elseif ($errMsg)    { $script:rptStats.Failed++  }
                        else                { $script:rptStats.Success++ }

                        if ($isSuccess -and -not $errMsg -and -not $isFail) { Update-RptStats; continue }
                        if (-not $errMsg -and -not $errType -and -not $isFail -and -not $isWarn) { Update-RptStats; continue }

                        $analysis   = Get-ErrorAnalysis $errMsg
                        $srcDisplay = if ($source)  { $source }                                   else { '—' }
                        $stDisplay  = if ($status)  { $status }                                   else { '—' }
                        $errDisplay = if ($errMsg)  { $errMsg }                                   else { '—' }
                        $typeDisplay= if ($errType) { $errType } elseif ($analysis.Type) { $analysis.Type } else { '—' }
                        $fixDisplay = if ($fix)     { $fix }     elseif ($analysis.Fix)  { $analysis.Fix }  else { '—' }

                        $dgv.Rows.Add($display, $srcDisplay, $stDisplay, $typeDisplay, $errDisplay, $fixDisplay) | Out-Null
                        $script:rptExportRows.Add([pscustomobject]@{
                            Workload       = $display
                            SourceObject   = $srcDisplay
                            Status         = $stDisplay
                            ErrorType      = $typeDisplay
                            ErrorMessage   = $errDisplay
                            RecommendedFix = $fixDisplay
                        }) | Out-Null

                        Update-RptStats
                        [System.Windows.Forms.Application]::DoEvents()
                    }
                    Update-RptStats
                } catch {
                    Write-Log "[$wl] Error: $($_.Exception.Message)" "ERROR"
                }
            }
        }

        $failTxt = if ($script:rptStats.Failed -gt 0) { "$($script:rptStats.Failed) failed" } else { "no failures" }
        $lblRStatus.Text      = "Done. $($script:rptStats.Total) total — $failTxt."
        $lblRStatus.ForeColor = if ($script:rptStats.Failed -gt 0) { $clrRed } else { $clrGreen }
        $btnRRun.Enabled      = $true
        if ($script:rptExportRows.Count -gt 0) { $btnRExport.Enabled = $true }
        Write-Log "Complete: $($script:rptStats.Total) rows, $($script:rptStats.Failed) failed, $($script:rptStats.Warn) warnings." "OK"
    })

    # ── Export CSV ─────────────────────────────────────────────────────────
    $btnRExport.Add_Click({
        if ($script:rptExportRows.Count -eq 0) { Write-Log "No error rows to export." "WARN"; return }
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Filter   = "CSV files (*.csv)|*.csv"
        $dlg.FileName = "fly-report-errors-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
        if ($dlg.ShowDialog() -eq 'OK') {
            $escape  = { param($v) '"' + ([string]$v -replace '"', '""') + '"' }
            $headers = @('Workload','Source Object','Status','Error Type','Error Message','Recommended Fix')
            $lines   = [System.Collections.Generic.List[string]]::new()
            $lines.Add(($headers | ForEach-Object { & $escape $_ }) -join ',')
            foreach ($r in $script:rptExportRows) {
                $lines.Add((@($r.Workload, $r.SourceObject, $r.Status, $r.ErrorType, $r.ErrorMessage, $r.RecommendedFix) |
                    ForEach-Object { & $escape $_ }) -join ',')
            }
            $lines | Set-Content -Path $dlg.FileName -Encoding UTF8
            Write-Log "Exported: $($dlg.FileName)" "OK"
        }
    })

    # ── Clear ──────────────────────────────────────────────────────────────
    $btnRClear.Add_Click({
        $dgv.Rows.Clear(); $script:rptExportRows.Clear()
        $script:rptStats = @{ Total = 0; Success = 0; Warn = 0; Failed = 0 }
        Update-RptStats; $rptLog.Clear(); $btnRExport.Enabled = $false
        $lblRStatus.Text      = if ($btnRRun.Enabled) { "Ready" } else { "Not connected" }
        $lblRStatus.ForeColor = $clrMuted
    })

    $sharedCfg = Read-SharedConfig
    $allPfx = if ($sharedCfg.PSObject.Properties['Prefixes']) { @($sharedCfg.Prefixes) } `
              elseif ($sharedCfg.TenantName) { @($sharedCfg.TenantName) } else { @() }
    foreach ($pfx in $allPfx) { if ($pfx) { [void]$tbRPrefix.Items.Add($pfx) } }
    if ($tbRPrefix.Items.Count -gt 0 -and [string]::IsNullOrWhiteSpace($tbRPrefix.Text)) {
        $tbRPrefix.Text = $tbRPrefix.Items[0]
    }

    # Auto-connect on load using saved config
    $Form.Add_Shown({
        $initCfg = Read-RptConfig
        if ($initCfg) {
            Connect-RptFly -Cfg $initCfg
        } else {
            $dotRConn.BackColor   = $clrAmber
            $lblRStatus.Text      = "No credentials — click Credentials..."
            $lblRStatus.ForeColor = $clrAmber
            Write-Log "No saved credentials found. Click 'Credentials...' to set up Fly API access." "WARN"
        }
    })

    [void]$Form.ShowDialog()
}
