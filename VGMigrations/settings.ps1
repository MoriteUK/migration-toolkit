# ═════════════════════════════════════════════════════════════════════════════
# MAIN MENU
# ═════════════════════════════════════════════════════════════════════════════
function Show-SettingsDialog {
    $flyApiCfgPath = Join-Path $env:APPDATA "FlyMigration\config.json"
    $wlCfgPath     = Join-Path $PSScriptRoot "workloads.json"
    $wlOrder       = @('SharePoint','Exchange','OneDrive','Teams','Teams Chat','Groups')

    # ── Read existing config ────────────────────────────────────────────
    $apiUrl = ''; $apiCid = ''; $apiSec = ''
    if (Test-Path $flyApiCfgPath) {
        try {
            $raw = Get-Content $flyApiCfgPath -Raw | ConvertFrom-Json
            $apiUrl = [string]$raw.Url; $apiCid = [string]$raw.ClientId
            if ($raw.EncSecret) {
                try {
                    $s = $raw.EncSecret | ConvertTo-SecureString
                    $b = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($s)
                    $apiSec = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($b)
                    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($b)
                } catch {}
            }
        } catch {}
    }
    $sharedCfg = Read-SharedConfig
    $wlCfg = [ordered]@{}
    foreach ($wl in $wlOrder) { $wlCfg[$wl] = @{ Policy=''; Source=''; Destination=''; ProjectSuffix='' } }
    if (Test-Path $wlCfgPath) {
        try {
            $raw = Get-Content $wlCfgPath -Raw | ConvertFrom-Json
            foreach ($wl in $wlOrder) {
                if ($raw.PSObject.Properties[$wl]) {
                    $wlCfg[$wl] = @{
                        Policy        = [string]$raw.$wl.Policy
                        Source        = [string]$raw.$wl.Source
                        Destination   = [string]$raw.$wl.Destination
                        ProjectSuffix = if ($raw.$wl.PSObject.Properties['ProjectSuffix']) { [string]$raw.$wl.ProjectSuffix } else { '' }
                    }
                }
            }
        } catch {}
    }

    # ── Form ────────────────────────────────────────────────────────────
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Settings"
    $dlg.ClientSize      = [System.Drawing.Size]::new(900, 540)
    $dlg.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox     = $false; $dlg.MinimizeBox = $false
    $dlg.BackColor       = $clrBg; $dlg.Font = $FontBody

    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Height = 46; $hdr.Dock = [System.Windows.Forms.DockStyle]::Top; $hdr.BackColor = $clrAccent
    $dlg.Controls.Add($hdr)
    $hdrLbl = New-Object System.Windows.Forms.Label
    $hdrLbl.Text = "  ⚙  Settings"
    $hdrLbl.Font = $FontBold
    $hdrLbl.ForeColor = [System.Drawing.Color]::White
    $hdrLbl.Location = [System.Drawing.Point]::new(8, 10); $hdrLbl.AutoSize = $true
    $hdr.Controls.Add($hdrLbl)

    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Location = [System.Drawing.Point]::new(10, 54)
    $tabs.Size     = [System.Drawing.Size]::new(876, 432)
    $dlg.Controls.Add($tabs)

    # Helper scriptblocks for tab content
    $mkCapLbl = {
        param($parent, $text, $x, $y)
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $text; $l.Font = $FontCap; $l.ForeColor = $clrMuted
        $l.Location = [System.Drawing.Point]::new($x, $y); $l.AutoSize = $true
        $parent.Controls.Add($l)
    }
    $mkTb = {
        param($parent, $x, $y, $w, [bool]$isPass = $false)
        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Location    = [System.Drawing.Point]::new($x, $y)
        $tb.Size        = [System.Drawing.Size]::new($w, 24)
        $tb.Font        = $FontBody
        $tb.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        if ($isPass) { $tb.UseSystemPasswordChar = $true }
        $parent.Controls.Add($tb); $tb
    }
    $mkNote = {
        param($parent, $text, $x, $y)
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $text; $l.Font = $FontSub
        $l.ForeColor = $clrMuted; $l.Location = [System.Drawing.Point]::new($x, $y)
        $l.Size = [System.Drawing.Size]::new(850, 16); $parent.Controls.Add($l)
    }

    # ── Tab 1: Fly API ─────────────────────────────────────────────────
    $tp1 = New-Object System.Windows.Forms.TabPage
    $tp1.Text = "Config"; $tp1.BackColor = $clrBg; $tp1.UseVisualStyleBackColor = $true
    $tabs.TabPages.Add($tp1)

    & $mkCapLbl $tp1 "FLY API URL" 10 16
    $tbUrl = & $mkTb $tp1 10 32 850
    $tbUrl.Text = $apiUrl

    & $mkCapLbl $tp1 "AOS CLIENT ID" 10 68
    $tbCid = & $mkTb $tp1 10 84 280
    $tbCid.Text = $apiCid

    & $mkCapLbl $tp1 "CLIENT SECRET" 300 68
    $tbSec = & $mkTb $tp1 300 84 230 $true
    $tbSec.Text = $apiSec

    & $mkCapLbl $tp1 "SECRET EXPIRY (yyyy-MM-dd)" 540 68
    $tbExpiry = & $mkTb $tp1 540 84 160
    $tbExpiry.Text = if ($sharedCfg.SecretExpiry) {
        try { ([datetime]$sharedCfg.SecretExpiry).ToString('yyyy-MM-dd') } catch { [string]$sharedCfg.SecretExpiry }
    } else { '' }

    $btnTest = New-Btn $tp1 "Test Connection" 10 128 140 28
    $dotTest = New-Dot $tp1 158 134
    $lblTestResult = New-Object System.Windows.Forms.Label
    $lblTestResult.Location = [System.Drawing.Point]::new(178, 134)
    $lblTestResult.AutoSize = $true; $lblTestResult.Font = $FontBody; $lblTestResult.ForeColor = $clrMuted
    $tp1.Controls.Add($lblTestResult)

    & $mkCapLbl $tp1 "FLY PORTAL URL" 10 172
    $tbPortal = & $mkTb $tp1 10 188 530
    $tbPortal.Text = if ($sharedCfg.PortalUrl) { $sharedCfg.PortalUrl } else { '' }
    & $mkNote $tp1 "Web portal base URL — used to open projects when double-clicking in the Project Monitor" 10 216

    & $mkCapLbl $tp1 "SHAREPOINT ADMIN URL" 10 240
    $tbSpoAdmin = & $mkTb $tp1 10 256 530
    $tbSpoAdmin.Text = if ($sharedCfg.SharePointAdminUrl) { $sharedCfg.SharePointAdminUrl } else { '' }
    try { $tbSpoAdmin.PlaceholderText = 'https://tenant-admin.sharepoint.com' } catch {}

    # Check for Updates button
    & $mkCapLbl $tp1 "SOFTWARE UPDATES" 10 288
    $btnCheckUpdates = New-Btn $tp1 "Check for Updates" 10 304 160 28
    $lblUpdateStatus = New-Object System.Windows.Forms.Label
    $lblUpdateStatus.Location = [System.Drawing.Point]::new(180, 310)
    $lblUpdateStatus.AutoSize = $true
    $lblUpdateStatus.Font = $FontBody
    $lblUpdateStatus.ForeColor = $clrMuted
    $lblUpdateStatus.Text = ''
    $tp1.Controls.Add($lblUpdateStatus)

    $btnCheckUpdates.Add_Click({
        $checkUpdatesScript = Join-Path $PSScriptRoot 'Check-Updates.ps1'
        if (-not (Test-Path $checkUpdatesScript)) {
            $lblUpdateStatus.Text = 'Check-Updates.ps1 not found'
            $lblUpdateStatus.ForeColor = $clrRed
            return
        }

        $lblUpdateStatus.Text = 'Checking for updates...'
        $lblUpdateStatus.ForeColor = $clrMuted
        $btnCheckUpdates.Enabled = $false
        [System.Windows.Forms.Application]::DoEvents()

        try {
            # Run Check-Updates.ps1 with -Force to show results
            $result = & $checkUpdatesScript -Force 2>&1 | Out-String

            if ($result -match 'Already up to date') {
                $lblUpdateStatus.Text = 'Already up to date'
                $lblUpdateStatus.ForeColor = $clrGreen
            } elseif ($result -match 'Updated to|Installation complete') {
                $lblUpdateStatus.Text = 'Update installed! Restart required.'
                $lblUpdateStatus.ForeColor = $clrGreen
                [System.Windows.Forms.MessageBox]::Show(
                    "Update installed successfully!`n`nPlease close and restart the application to use the new version.",
                    'Update Complete',
                    'OK',
                    'Information'
                ) | Out-Null
            } else {
                $lblUpdateStatus.Text = 'Update check complete'
                $lblUpdateStatus.ForeColor = $clrGreen
            }
        } catch {
            $lblUpdateStatus.Text = "Update check failed: $_"
            $lblUpdateStatus.ForeColor = $clrRed
        } finally {
            $btnCheckUpdates.Enabled = $true
        }
    }.GetNewClosure())

    $btnTest.Add_Click({
        $dotTest.ForeColor = $clrGrey; $lblTestResult.Text = "Connecting..."
        [System.Windows.Forms.Application]::DoEvents()
        try {
            $u = $tbUrl.Text.Trim(); $c = $tbCid.Text.Trim(); $s = $tbSec.Text.Trim()
            if (-not $c -or -not $s) { throw "Client ID and Secret are required." }
            if (-not (Get-Module Fly.Client -ListAvailable)) { throw "Fly.Client module not found." }
            Import-Module Fly.Client -ErrorAction Stop
            Connect-Fly -Url $u -ClientId $c -ClientSecret $s -ErrorAction Stop
            $dotTest.ForeColor = $clrGreen; $lblTestResult.Text = "Connected successfully"
            $lblTestResult.ForeColor = $clrGreen
        } catch {
            $dotTest.ForeColor = $clrRed
            $lblTestResult.Text = $_.Exception.Message; $lblTestResult.ForeColor = $clrRed
        }
        [System.Windows.Forms.Application]::DoEvents()
    }.GetNewClosure())

    # ── Tab 2: Customer ────────────────────────────────────────────────
    $tp2 = New-Object System.Windows.Forms.TabPage
    $tp2.Text = "Customer"; $tp2.BackColor = $clrBg; $tp2.UseVisualStyleBackColor = $true
    $tabs.TabPages.Add($tp2)

    & $mkCapLbl $tp2 "CUSTOMER PREFIXES" 10 10
    & $mkNote $tp2 "Customer Prefix is the project name prefix. Account Name auto-fills the connection details in Migration Mappings." 10 28

    # Build DataTable from stored Customers array (or migrate from old Prefixes array)
    $pfxTable = New-Object System.Data.DataTable
    $pfxTable.Columns.Add("Prefix",             [string]) | Out-Null
    $pfxTable.Columns.Add("AccountName",        [string]) | Out-Null
    $pfxTable.Columns.Add("SharePointAdminUrl", [string]) | Out-Null

    $existingCustomers = @()
    if ($sharedCfg.PSObject.Properties['Customers'] -and $sharedCfg.Customers) {
        $existingCustomers = @($sharedCfg.Customers)
    } elseif ($sharedCfg.PSObject.Properties['Prefixes'] -and $sharedCfg.Prefixes) {
        $existingCustomers = @($sharedCfg.Prefixes | Where-Object { $_ } | ForEach-Object {
            [pscustomobject]@{ Prefix = "$_"; AccountName = ''; SharePointAdminUrl = '' }
        })
    } elseif ($sharedCfg.TenantName) {
        $existingCustomers = @([pscustomobject]@{ Prefix = $sharedCfg.TenantName; AccountName = ''; SharePointAdminUrl = '' })
    }
    foreach ($c in $existingCustomers) {
        if ($c.Prefix) {
            $r = $pfxTable.NewRow()
            $r['Prefix']             = "$($c.Prefix)"
            $r['AccountName']        = "$($c.AccountName)"
            $r['SharePointAdminUrl'] = if ($c.PSObject.Properties['SharePointAdminUrl']) { "$($c.SharePointAdminUrl)" } else { '' }
            $pfxTable.Rows.Add($r)
        }
    }

    $dgvPfx = New-Object System.Windows.Forms.DataGridView
    $dgvPfx.Location          = [System.Drawing.Point]::new(10, 46)
    $dgvPfx.Size              = [System.Drawing.Size]::new(850, 168)
    $dgvPfx.Font              = $FontBody
    $dgvPfx.ReadOnly          = $false
    $dgvPfx.AutoGenerateColumns  = $false
    $dgvPfx.AllowUserToAddRows   = $false
    $dgvPfx.AllowUserToDeleteRows = $false
    $dgvPfx.SelectionMode     = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $dgvPfx.MultiSelect       = $false
    $dgvPfx.BorderStyle       = [System.Windows.Forms.BorderStyle]::FixedSingle
    $dgvPfx.BackgroundColor   = $clrPanel
    $dgvPfx.GridColor         = $clrBorder
    $dgvPfx.DefaultCellStyle.BackColor  = $clrPanel
    $dgvPfx.DefaultCellStyle.ForeColor  = $clrText
    $dgvPfx.DefaultCellStyle.Font       = $FontBody
    $dgvPfx.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(245, 246, 250)
    $dgvPfx.ColumnHeadersDefaultCellStyle.Font       = $FontBold
    $dgvPfx.ColumnHeadersDefaultCellStyle.BackColor  = $clrAccentTint
    $dgvPfx.ColumnHeadersDefaultCellStyle.ForeColor  = $clrText
    $dgvPfx.ColumnHeadersHeight           = 28
    $dgvPfx.ColumnHeadersHeightSizeMode   = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $dgvPfx.EnableHeadersVisualStyles     = $false
    $dgvPfx.RowTemplate.Height            = 26

    $colPfx  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colPfx.DataPropertyName = "Prefix"; $colPfx.HeaderText = "Customer Prefix"; $colPfx.Width = 160
    $dgvPfx.Columns.Add($colPfx) | Out-Null

    $colAcct = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colAcct.DataPropertyName = "AccountName"; $colAcct.HeaderText = "Account Name"; $colAcct.Width = 220
    $dgvPfx.Columns.Add($colAcct) | Out-Null

    $colSpAdm = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colSpAdm.DataPropertyName = "SharePointAdminUrl"; $colSpAdm.HeaderText = "SharePoint Admin URL"
    $colSpAdm.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $dgvPfx.Columns.Add($colSpAdm) | Out-Null

    $dgvPfx.DataSource = $pfxTable

    $pfxTable.Add_ColumnChanged({
        param($s, $e)
        if ($e.Column.ColumnName -eq 'AccountName' -and -not "$($e.Row['SharePointAdminUrl'])") {
            if ("$($e.Row['AccountName'])" -match '@([\w-]+)\.onmicrosoft\.com') {
                $e.Row['SharePointAdminUrl'] = "https://$($Matches[1])-admin.sharepoint.com"
            }
        }
    })

    $tp2.Controls.Add($dgvPfx)

    $btnAddPfx = New-Btn $tp2 "+ Add Row"    10  220  90 26 $false
    $btnRemPfx = New-Btn $tp2 "- Remove Row" 106 220 106 26 $false

    $btnAddPfx.Add_Click({
        $r = $pfxTable.NewRow(); $r['Prefix'] = ''; $r['AccountName'] = ''; $r['SharePointAdminUrl'] = ''
        $pfxTable.Rows.Add($r)
        $dgvPfx.ClearSelection()
        $dgvPfx.Rows[$dgvPfx.Rows.Count - 1].Selected = $true
        $dgvPfx.CurrentCell = $dgvPfx.Rows[$dgvPfx.Rows.Count - 1].Cells[0]
        $dgvPfx.BeginEdit($true)
    }.GetNewClosure())

    $btnRemPfx.Add_Click({
        if ($dgvPfx.SelectedRows.Count -gt 0) {
            $idx = $dgvPfx.SelectedRows[0].Index
            $pfxTable.Rows.RemoveAt($idx)
        }
    }.GetNewClosure())

    # ── Tab 3: Workloads ───────────────────────────────────────────────
    $tp3 = New-Object System.Windows.Forms.TabPage
    $tp3.Text = "Workloads"; $tp3.BackColor = $clrBg; $tp3.UseVisualStyleBackColor = $true
    $tabs.TabPages.Add($tp3)

    $scroll = New-Object System.Windows.Forms.Panel
    $scroll.AutoScroll = $true; $scroll.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tp3.Controls.Add($scroll)
    & $mkNote $scroll "These are pre-filled into the Migration Runner for each workload." 10 4

    # Column headers
    $lhSfx = New-Object System.Windows.Forms.Label
    $lhSfx.Text = "PROJECT SUFFIX"; $lhSfx.Font = $FontCap; $lhSfx.ForeColor = $clrMuted
    $lhSfx.Location = [System.Drawing.Point]::new(118, 24); $lhSfx.AutoSize = $true
    $scroll.Controls.Add($lhSfx)

    $wlControls = [ordered]@{}
    $wy = 42
    foreach ($wl in $wlOrder) {
        $wlLbl = New-Object System.Windows.Forms.Label
        $wlLbl.Text = $wl; $wlLbl.Font = $FontBold; $wlLbl.ForeColor = $clrText
        $wlLbl.Location = [System.Drawing.Point]::new(10, $wy + 3); $wlLbl.Size = [System.Drawing.Size]::new(106, 20)
        $scroll.Controls.Add($wlLbl)

        $tbSuffix = New-Object System.Windows.Forms.TextBox
        $tbSuffix.Location    = [System.Drawing.Point]::new(118, $wy)
        $tbSuffix.Size        = [System.Drawing.Size]::new(736, 24)
        $tbSuffix.Font        = $FontBody
        $tbSuffix.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $scroll.Controls.Add($tbSuffix)
        $tbSuffix.Text      = if ($wlCfg[$wl].ProjectSuffix) { $wlCfg[$wl].ProjectSuffix } else { $wl }
        $tbSuffix.ForeColor = if ($wlCfg[$wl].ProjectSuffix) { $clrText } else { $clrMuted }

        $sep = New-Object System.Windows.Forms.Label
        $sep.Location = [System.Drawing.Point]::new(10, $wy + 30)
        $sep.Size     = [System.Drawing.Size]::new(856, 1)
        $sep.BackColor = $clrBorder
        $scroll.Controls.Add($sep)

        $wlControls[$wl] = @{ ProjectSuffix = $tbSuffix }
        $wy += 40
    }

    # ── Tab 4: Discovery ───────────────────────────────────────────────
    $tp4 = New-Object System.Windows.Forms.TabPage
    $tp4.Text = "Discovery"; $tp4.BackColor = $clrBg; $tp4.UseVisualStyleBackColor = $true
    $tabs.TabPages.Add($tp4)

    & $mkCapLbl $tp4 "DISCOVERY OUTPUT FOLDER" 10 16
    $tbDiscPath = & $mkTb $tp4 10 32 790
    $tbDiscPath.Text = if ($sharedCfg.PSObject.Properties['DiscoveryOutputPath'] -and $sharedCfg.DiscoveryOutputPath) { $sharedCfg.DiscoveryOutputPath } else { '' }
    & $mkNote $tp4 "Base folder where discovery CSVs and Excel files are saved. Each domain gets its own subfolder inside." 10 60

    $btnBrowseDiscPath = New-Btn $tp4 "Browse..." 10 82 90 26 $false
    $btnBrowseDiscPath.Add_Click({
        $fbd = [System.Windows.Forms.FolderBrowserDialog]::new()
        $fbd.Description  = 'Select the base output folder for discovery results'
        $fbd.SelectedPath = if ($tbDiscPath.Text -and (Test-Path $tbDiscPath.Text -ErrorAction SilentlyContinue)) { $tbDiscPath.Text } else { $env:USERPROFILE }
        if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $tbDiscPath.Text = $fbd.SelectedPath }
    }.GetNewClosure())

    # ── Footer ─────────────────────────────────────────────────────────
    $btnSave   = New-Btn $dlg "Save All" 672 496 112 32
    $btnCancel = New-Btn $dlg "Cancel"   792 496 100 32 $false

    $btnSave.Add_Click({
        try {
            # Fly API
            $u = $tbUrl.Text.Trim(); $c = $tbCid.Text.Trim(); $s = $tbSec.Text.Trim()
            $dir = Split-Path $flyApiCfgPath
            if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
            $encSec = if ($s) { $s | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString } else { $null }
            $json = @{ Url = $u; ClientId = $c; EncSecret = $encSec } | ConvertTo-Json
            [System.IO.File]::WriteAllText($flyApiCfgPath, $json, (New-Object System.Text.UTF8Encoding $false))

            # Shared config
            $customers = @($pfxTable.Rows | ForEach-Object {
                [pscustomobject]@{ Prefix = "$($_['Prefix'])"; AccountName = "$($_['AccountName'])"; SharePointAdminUrl = "$($_['SharePointAdminUrl'])" }
            } | Where-Object { $_.Prefix })
            $pfxList   = @($customers | ForEach-Object { $_.Prefix })
            $expVal       = $tbExpiry.Text.Trim()
            $portalVal    = $tbPortal.Text.Trim()
            $spoAdminVal  = $tbSpoAdmin.Text.Trim()
            $upd = @{}
            $upd['Customers']   = $customers
            $upd['Prefixes']    = $pfxList
            $upd['TenantName']  = if ($pfxList.Count -gt 0) { $pfxList[0] } else { '' }
            if ($expVal      -or $sharedCfg.SecretExpiry)       { $upd['SecretExpiry']       = $expVal }
            if ($portalVal   -or $sharedCfg.PortalUrl)          { $upd['PortalUrl']           = $portalVal }
            if ($spoAdminVal -or $sharedCfg.SharePointAdminUrl) { $upd['SharePointAdminUrl']  = $spoAdminVal }
            $discPathVal = $tbDiscPath.Text.Trim()
            if ($discPathVal -or ($sharedCfg.PSObject.Properties['DiscoveryOutputPath'] -and $sharedCfg.DiscoveryOutputPath)) { $upd['DiscoveryOutputPath'] = $discPathVal }
            Update-SharedConfig $upd

            # Workload config — preserve Policy/Source/Destination (static or auto-set by Create Connections)
            $wlOut = [ordered]@{}
            foreach ($wl in $wlOrder) {
                $sfx = $wlControls[$wl].ProjectSuffix.Text.Trim()
                $wlOut[$wl] = [ordered]@{
                    Policy        = if ($wlCfg[$wl].Policy)      { [string]$wlCfg[$wl].Policy }      else { '' }
                    Source        = if ($wlCfg[$wl].Source)      { [string]$wlCfg[$wl].Source }      else { '' }
                    Destination   = if ($wlCfg[$wl].Destination) { [string]$wlCfg[$wl].Destination } else { '' }
                    ProjectSuffix = if ($sfx) { $sfx } else { $wl }
                }
            }
            $wlOut | ConvertTo-Json -Depth 3 | Set-Content $wlCfgPath -Encoding UTF8

            [System.Windows.Forms.MessageBox]::Show("Settings saved.", "Saved", 'OK', 'Information') | Out-Null
            $dlg.Close()
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Save failed: $($_.Exception.Message)", "Error", 'OK', 'Error') | Out-Null
        }
    }.GetNewClosure())

    $btnCancel.Add_Click({ $dlg.Close() })
    $dlg.CancelButton = $btnCancel
    $dlg.ShowDialog() | Out-Null
}
