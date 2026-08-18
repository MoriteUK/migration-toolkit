#Requires -Version 7.0
<#
.SYNOPSIS
    Search-OneDriveFiles.ps1 — Searchable inventory of a single user's OneDrive, including
    files still in the Recycle Bin (both stages). Use this when a user reports missing files:
    it shows whether a file is still live, sitting in their own Recycle Bin (user-recoverable),
    or in the second-stage Recycle Bin (admin-recoverable only).

.DESCRIPTION
    Interactive WinForms tool. Connects via PnP PowerShell (browser sign-in) as a SharePoint/
    Global admin, temporarily grants that admin site collection access to the target user's
    OneDrive, then enumerates:
      - Every file currently in the "Documents" library (recursively — all folders)
      - Every item in the first-stage Recycle Bin (user-deleted, still user-recoverable)
      - Every item in the second-stage Recycle Bin (emptied by user, admin-recoverable only)

    Results load into a searchable, sortable grid and can be exported to CSV.

    Microsoft Graph has no API to list a OneDrive-for-Business Recycle Bin (its recycleBin
    endpoint only covers SharePoint Embedded containers), so this uses the classic SharePoint
    REST surface via PnP PowerShell instead. PnP's own multi-tenant Entra ID app is used for
    interactive sign-in — no app registration or admin consent is required up front (a one-time
    tenant admin consent prompt may appear on first use in a given tenant, which is expected).

.NOTES
    Dependency : lib.ps1 (colours, fonts, header/footer helpers), PnP.PowerShell module
                 (auto-installed to CurrentUser scope if missing)
    Log file   : %APPDATA%\FlyMigration\Logs\search-onedrive-files-<timestamp>.log

    IMPORTANT — access granted, not revoked: this script adds the signed-in admin as a site
    collection administrator on the target OneDrive via Set-PnPTenantSite so it can enumerate
    files and the Recycle Bin. It does not remove that access afterwards (Set-PnPTenantSite has
    no clean single-user "remove owner" call that's guaranteed not to disturb existing site
    owners). If your tenant's access-review policy requires it, remove the admin as a site
    collection admin manually afterwards via the SharePoint admin center.

    Change log
    ----------
    2026-08-18  Initial version.
#>

$libPath = Join-Path $PSScriptRoot 'lib.ps1'

$_logDir = Join-Path $env:APPDATA 'FlyMigration\Logs'
if (-not (Test-Path $_logDir)) { New-Item -ItemType Directory -Path $_logDir -Force | Out-Null }
$script:LogFile = Join-Path $_logDir "search-onedrive-files-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

function _RawLog {
    param([string]$Msg)
    "$(Get-Date -Format 'HH:mm:ss.fff') $Msg" | Add-Content -Path $script:LogFile -Encoding UTF8 -ErrorAction SilentlyContinue
}

# Top-level (not nested in a scriptblock) so it's visible from any GetNewClosure()'d event
# handler — functions defined inside a closure are NOT captured by GetNewClosure(), only
# variables are, so a nested "function Set-CtlText" would be invisible to sibling handlers
# like the Timer's Add_Tick and throw CommandNotFoundException when called from there.
function Set-CtlText {
    param($Control, [string]$Value)
    if ($Control -and -not $Control.IsDisposed) { $Control.Text = $Value }
}

_RawLog "=== Search-OneDriveFiles.ps1 started  PID=$PID  PSVersion=$($PSVersionTable.PSVersion) ==="

if (Test-Path $libPath) {
    try { . $libPath; _RawLog "lib.ps1 loaded OK" }
    catch { _RawLog "lib.ps1 LOAD ERROR: $($_.Exception.Message)" }
} else {
    _RawLog "lib.ps1 NOT FOUND — colours and helpers will be missing"
}

# ─────────────────────────────────────────────────────────────
# Saved customers (for SharePoint Admin URL auto-fill)
# ─────────────────────────────────────────────────────────────
function Get-SavedCustomers {
    $cfgPath = Join-Path $env:APPDATA 'FlyMigration\config.json'
    if (-not (Test-Path $cfgPath)) { return @() }
    try {
        $cfg = Get-Content $cfgPath -Raw | ConvertFrom-Json
        return @($cfg.Customers | Where-Object { $_.SharePointAdminUrl })
    } catch { return @() }
}

function Show-SearchOneDriveUI {

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'Search OneDrive Files'
    $form.ClientSize      = [System.Drawing.Size]::new(900, 720)
    $form.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.BackColor       = $clrBg
    $form.Font            = $FontBody
    $form.MinimumSize     = [System.Drawing.Size]::new(700, 500)
    $_ico = Join-Path $PSScriptRoot 'FlyMigration.ico'
    if (Test-Path $_ico) { $form.Icon = [System.Drawing.Icon]::new($_ico) }

    # ── Header ────────────────────────────────────────────────────────────────
    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Size = [System.Drawing.Size]::new(900, 56); $hdr.Dock = [System.Windows.Forms.DockStyle]::Top
    $hdr.BackColor = $clrAccent; $form.Controls.Add($hdr)
    $_hdrX = Add-HeaderLogo $hdr 36
    $hdrLbl = New-Object System.Windows.Forms.Label
    $hdrLbl.Text = '  Search OneDrive Files (incl. Recycle Bin)'; $hdrLbl.Font = $FontTitle
    $hdrLbl.ForeColor = [System.Drawing.Color]::White
    $hdrLbl.Location  = [System.Drawing.Point]::new($_hdrX, 0)
    $hdrLbl.Size      = [System.Drawing.Size]::new(800, 56)
    $hdrLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdr.Controls.Add($hdrLbl)

    # ── Footer ────────────────────────────────────────────────────────────────
    $footer = New-Object System.Windows.Forms.Panel
    $footer.Height = 46; $footer.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $footer.BackColor = $clrFooter
    $form.Controls.Add($footer)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'; $btnClose.Size = [System.Drawing.Size]::new(90, 30)
    $btnClose.BackColor = $clrCloseRed
    $btnClose.ForeColor = [System.Drawing.Color]::White; $btnClose.Font = $FontBold
    $btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnClose.FlatAppearance.BorderSize = 0
    $btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnClose.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $footer.Controls.Add($btnClose)
    $btnClose.Add_Click({ $form.Close() }.GetNewClosure())

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = 'Export CSV...'; $btnExport.Size = [System.Drawing.Size]::new(110, 30)
    $btnExport.BackColor = [System.Drawing.Color]::FromArgb(225, 228, 238); $btnExport.ForeColor = $clrText
    $btnExport.Font = $FontBold; $btnExport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnExport.FlatAppearance.BorderSize = 0; $btnExport.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnExport.Enabled = $false
    $btnExport.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $footer.Controls.Add($btnExport)

    $footer.Add_SizeChanged({
        $btnClose.Left  = $footer.Width - 100
        $btnExport.Left = $btnClose.Left - 120
    }.GetNewClosure())

    # ── Input card ────────────────────────────────────────────────────────────
    $card = New-Object System.Windows.Forms.Panel
    $card.Location  = [System.Drawing.Point]::new(12, 66)
    $card.Size      = [System.Drawing.Size]::new(876, 190)
    $card.Anchor    = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $card.BackColor = $clrPanel
    $form.Controls.Add($card)

    $lx = 16; $ex = 190; $ew = 400; $y = 14

    $lbUpn = New-Object System.Windows.Forms.Label
    $lbUpn.Text = "User's OneDrive UPN:"; $lbUpn.Font = $FontBold; $lbUpn.ForeColor = $clrText
    $lbUpn.Location = [System.Drawing.Point]::new($lx, $y + 5); $lbUpn.AutoSize = $true
    $card.Controls.Add($lbUpn)
    $txtUpn = New-Object System.Windows.Forms.TextBox
    $txtUpn.Location = [System.Drawing.Point]::new($ex, $y); $txtUpn.Size = [System.Drawing.Size]::new($ew, 24)
    $txtUpn.Font = $FontBody; $txtUpn.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    try { $txtUpn.PlaceholderText = 'jane.doe@contoso.com' } catch {}
    $card.Controls.Add($txtUpn)
    $y += 34

    $lbCust = New-Object System.Windows.Forms.Label
    $lbCust.Text = 'Saved customer:'; $lbCust.Font = $FontBold; $lbCust.ForeColor = $clrText
    $lbCust.Location = [System.Drawing.Point]::new($lx, $y + 5); $lbCust.AutoSize = $true
    $card.Controls.Add($lbCust)
    $cboCustomer = New-Object System.Windows.Forms.ComboBox
    $cboCustomer.Location = [System.Drawing.Point]::new($ex, $y); $cboCustomer.Size = [System.Drawing.Size]::new($ew, 24)
    $cboCustomer.Font = $FontBody; $cboCustomer.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $script:savedCustomers = @(Get-SavedCustomers)
    [void]$cboCustomer.Items.Add('(choose to auto-fill admin URL below)')
    foreach ($c in $script:savedCustomers) { [void]$cboCustomer.Items.Add("$($c.Prefix) — $($c.SharePointAdminUrl)") }
    $cboCustomer.SelectedIndex = 0
    $card.Controls.Add($cboCustomer)
    $y += 34

    $lbAdminUrl = New-Object System.Windows.Forms.Label
    $lbAdminUrl.Text = 'SharePoint Admin URL:'; $lbAdminUrl.Font = $FontBold; $lbAdminUrl.ForeColor = $clrText
    $lbAdminUrl.Location = [System.Drawing.Point]::new($lx, $y + 5); $lbAdminUrl.AutoSize = $true
    $card.Controls.Add($lbAdminUrl)
    $txtAdminUrl = New-Object System.Windows.Forms.TextBox
    $txtAdminUrl.Location = [System.Drawing.Point]::new($ex, $y); $txtAdminUrl.Size = [System.Drawing.Size]::new($ew, 24)
    $txtAdminUrl.Font = $FontBody; $txtAdminUrl.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    try { $txtAdminUrl.PlaceholderText = 'https://contoso-admin.sharepoint.com' } catch {}
    $card.Controls.Add($txtAdminUrl)
    $cboCustomer.Add_SelectedIndexChanged({
        if ($cboCustomer.SelectedIndex -ge 1) {
            $txtAdminUrl.Text = $script:savedCustomers[$cboCustomer.SelectedIndex - 1].SharePointAdminUrl
        }
    }.GetNewClosure())
    $y += 34

    $lbYourUpn = New-Object System.Windows.Forms.Label
    $lbYourUpn.Text = 'Your admin UPN:'; $lbYourUpn.Font = $FontBold; $lbYourUpn.ForeColor = $clrText
    $lbYourUpn.Location = [System.Drawing.Point]::new($lx, $y + 5); $lbYourUpn.AutoSize = $true
    $card.Controls.Add($lbYourUpn)
    $txtYourUpn = New-Object System.Windows.Forms.TextBox
    $txtYourUpn.Location = [System.Drawing.Point]::new($ex, $y); $txtYourUpn.Size = [System.Drawing.Size]::new($ew, 24)
    $txtYourUpn.Font = $FontBody; $txtYourUpn.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    try { $txtYourUpn.PlaceholderText = 'you@contoso.com — needs SharePoint/Global admin role' } catch {}
    $card.Controls.Add($txtYourUpn)
    $y += 40

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = 'Sign In and Search'; $btnRun.Location = [System.Drawing.Point]::new($ex, $y)
    $btnRun.Size = [System.Drawing.Size]::new(180, 32); $btnRun.Font = $FontBold
    $btnRun.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnRun.FlatAppearance.BorderSize = 0
    $btnRun.BackColor = [System.Drawing.Color]::FromArgb(18, 140, 60)
    $btnRun.ForeColor = [System.Drawing.Color]::White
    $btnRun.Cursor = [System.Windows.Forms.Cursors]::Hand
    $card.Controls.Add($btnRun)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = 'Enter the target user, admin URL, and your own admin UPN, then click Sign In and Search.'
    $lblStatus.ForeColor = $clrMuted; $lblStatus.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $lblStatus.Location = [System.Drawing.Point]::new($ex + 190, $y + 7); $lblStatus.Size = [System.Drawing.Size]::new($ew - 190, 36)
    $card.Controls.Add($lblStatus)

    # ── Progress ──────────────────────────────────────────────────────────────
    $progress = New-Object System.Windows.Forms.ProgressBar
    $progress.Location = [System.Drawing.Point]::new(12, $card.Bottom + 8)
    $progress.Size     = [System.Drawing.Size]::new(876, 8)
    $progress.Anchor   = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $progress.Style    = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $form.Controls.Add($progress)

    # ── Search box ────────────────────────────────────────────────────────────
    $lbFilter = New-Object System.Windows.Forms.Label
    $lbFilter.Text = 'Filter:'; $lbFilter.Font = $FontBold; $lbFilter.ForeColor = $clrText
    $lbFilter.Location = [System.Drawing.Point]::new(12, $progress.Bottom + 12); $lbFilter.AutoSize = $true
    $lbFilter.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
    $form.Controls.Add($lbFilter)

    $txtFilter = New-Object System.Windows.Forms.TextBox
    $txtFilter.Location = [System.Drawing.Point]::new(60, $progress.Bottom + 8); $txtFilter.Size = [System.Drawing.Size]::new(828, 24)
    $txtFilter.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $txtFilter.Font = $FontBody; $txtFilter.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $txtFilter.Enabled = $false
    try { $txtFilter.PlaceholderText = 'Type to filter by file name, folder, or status...' } catch {}
    $form.Controls.Add($txtFilter)

    # ── Results grid ──────────────────────────────────────────────────────────
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = [System.Drawing.Point]::new(12, $txtFilter.Bottom + 8)
    $grid.Size     = [System.Drawing.Size]::new(876, $form.ClientSize.Height - $txtFilter.Bottom - 8 - 46 - 8)
    $grid.Anchor   = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grid.RowHeadersVisible = $false
    $grid.BackgroundColor = [System.Drawing.Color]::White
    $grid.Font = $FontBody
    $form.Controls.Add($grid)

    $dt = New-Object System.Data.DataTable
    [void]$dt.Columns.Add('Status')
    [void]$dt.Columns.Add('Name')
    [void]$dt.Columns.Add('Location')
    [void]$dt.Columns.Add('SizeKB', [double])
    [void]$dt.Columns.Add('Date')
    [void]$dt.Columns.Add('ModifiedOrDeletedBy')
    $dv = New-Object System.Data.DataView($dt)
    $grid.DataSource = $dv

    $txtFilter.Add_TextChanged({
        $q = $txtFilter.Text.Trim().Replace("'", "''")
        if ([string]::IsNullOrWhiteSpace($q)) { $dv.RowFilter = '' }
        else {
            $dv.RowFilter = "Name LIKE '%$q%' OR Location LIKE '%$q%' OR Status LIKE '%$q%' OR ModifiedOrDeletedBy LIKE '%$q%'"
        }
    }.GetNewClosure())

    # ── Export ────────────────────────────────────────────────────────────────
    $btnExport.Add_Click({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Title  = 'Save results as...'
        $sfd.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
        $upnSafe = $txtUpn.Text.Trim() -replace '[\\/:*?"<>|@]', '_'
        $sfd.FileName = "${upnSafe}_onedrive_files.csv"
        $sfd.InitialDirectory = $env:USERPROFILE
        if ($sfd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        try {
            $dt.DefaultView.ToTable() | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8 -Force
            [System.Windows.Forms.MessageBox]::Show("Exported $($dt.Rows.Count) row(s) to:`n$($sfd.FileName)", 'Export Complete', 'OK', 'Information') | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Export failed: $($_.Exception.Message)", 'Export Error', 'OK', 'Error') | Out-Null
        }
    }.GetNewClosure())

    # ── Run ───────────────────────────────────────────────────────────────────
    $script:sofTimer    = $null
    $script:sofRunspace = $null
    $script:sofPS       = $null
    $script:tickBusy    = $false

    $btnRun.Add_Click({
      try {
        $targetUpn  = $txtUpn.Text.Trim()
        $adminUrl   = $txtAdminUrl.Text.Trim().TrimEnd('/')
        $yourUpn    = $txtYourUpn.Text.Trim()

        if ([string]::IsNullOrWhiteSpace($targetUpn) -or $targetUpn -notmatch '^[^@]+@[^@]+\.[^@]+$') {
            [System.Windows.Forms.MessageBox]::Show('Enter a valid UPN for the target user.', 'Missing Input', 'OK', 'Warning') | Out-Null; return
        }
        if ([string]::IsNullOrWhiteSpace($adminUrl) -or $adminUrl -notmatch '^https://.+-admin\.sharepoint\.com$') {
            [System.Windows.Forms.MessageBox]::Show("Enter a valid SharePoint Admin URL, e.g. https://contoso-admin.sharepoint.com", 'Missing Input', 'OK', 'Warning') | Out-Null; return
        }
        if ([string]::IsNullOrWhiteSpace($yourUpn) -or $yourUpn -notmatch '^[^@]+@[^@]+\.[^@]+$') {
            [System.Windows.Forms.MessageBox]::Show('Enter your own admin UPN (needed to grant temporary OneDrive access).', 'Missing Input', 'OK', 'Warning') | Out-Null; return
        }

        $tenantHost = $adminUrl -replace '^https://(.+)-admin\.sharepoint\.com$', '$1'
        $oneDriveAccount = ($targetUpn -replace '[.@]', '_')
        $oneDriveUrl = "https://$tenantHost-my.sharepoint.com/personal/$oneDriveAccount"

        _RawLog "Run clicked — target=$targetUpn  oneDriveUrl=$oneDriveUrl  admin=$yourUpn"

        $btnRun.Enabled = $false; $btnExport.Enabled = $false; $btnClose.Enabled = $false
        $txtFilter.Enabled = $false; $txtFilter.Text = ''
        $dt.Rows.Clear()
        $progress.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
        $lblStatus.Text = 'Connecting — sign in with the browser window that opens...'

        $logFilePath = $script:LogFile

        $rs = [hashtable]::Synchronized(@{
            Done       = $false
            FatalError = $null
            Rows       = [System.Collections.Generic.List[object]]::new()
            LogQueue   = [System.Collections.Queue]::Synchronized([System.Collections.Queue]::new())
        })

        $script:sofRunspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $script:sofRunspace.ApartmentState = 'STA'
        $script:sofRunspace.ThreadOptions  = 'ReuseThread'
        $script:sofRunspace.Open()

        $script:sofPS = [System.Management.Automation.PowerShell]::Create()
        $script:sofPS.Runspace = $script:sofRunspace

        [void]$script:sofPS.AddScript({
            param($adminUrl, $oneDriveUrl, $yourUpn, $rs, $logFilePath)

            function QLog {
                param([string]$Msg, [string]$Level = 'INFO')
                $rs.LogQueue.Enqueue("[$Level] $Msg")
                try { "$(Get-Date -Format 'HH:mm:ss') [$Level] $Msg" | Add-Content $logFilePath -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
            }

            function GetProp {
                param($Obj, [string]$Name)
                if ($null -eq $Obj) { return $null }
                $p = $Obj.PSObject.Properties[$Name]
                if ($p) { return $p.Value }
                return $null
            }

            try {
                QLog 'Checking for PnP.PowerShell module...'
                if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
                    QLog 'PnP.PowerShell not found — installing from PSGallery (CurrentUser scope)...' 'WARN'
                    Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                }
                Import-Module PnP.PowerShell -ErrorAction Stop
                QLog 'PnP.PowerShell loaded.' 'OK'

                QLog "Connecting to SharePoint Admin ($adminUrl) — sign in as $yourUpn..."
                Connect-PnPOnline -Url $adminUrl -Interactive -ErrorAction Stop
                QLog 'Connected to SharePoint Admin.' 'OK'

                QLog "Granting temporary site collection admin on the target OneDrive to $yourUpn..."
                try {
                    Set-PnPTenantSite -Url $oneDriveUrl -Owners @($yourUpn) -ErrorAction Stop
                    QLog 'Access granted.' 'OK'
                } catch {
                    throw "Could not grant access to the target OneDrive ($oneDriveUrl). Confirm the UPN is correct and that the OneDrive has been provisioned. Underlying error: $($_.Exception.Message)"
                }

                Disconnect-PnPOnline -ErrorAction SilentlyContinue
                QLog "Connecting to the target OneDrive ($oneDriveUrl)..."
                Connect-PnPOnline -Url $oneDriveUrl -Interactive -ErrorAction Stop
                QLog 'Connected to target OneDrive.' 'OK'

                # ── Live files ────────────────────────────────────────────────
                QLog 'Locating the Documents library...'
                $lib = Get-PnPList -ErrorAction Stop | Where-Object { $_.BaseType -eq 'DocumentLibrary' -and -not $_.Hidden } | Select-Object -First 1
                if (-not $lib) { $lib = Get-PnPList -Identity 'Documents' -ErrorAction SilentlyContinue }

                if ($lib) {
                    QLog "Enumerating files in '$($lib.Title)' (this can take a while for large OneDrives)..."
                    $items = Get-PnPListItem -List $lib -PageSize 500 -Fields 'FileLeafRef','FileRef','FSObjType','Modified','Editor','File_x0020_Size' -ErrorAction Stop
                    $fileCount = 0
                    foreach ($item in $items) {
                        $fsType = GetProp $item.FieldValues 'FSObjType'
                        if ($fsType -ne 0) { continue }   # 0 = file, 1 = folder
                        $fileCount++
                        $name     = GetProp $item.FieldValues 'FileLeafRef'
                        $path     = GetProp $item.FieldValues 'FileRef'
                        $modified = GetProp $item.FieldValues 'Modified'
                        $editor   = GetProp $item.FieldValues 'Editor'
                        $editorName = if ($editor -and $editor.Email) { $editor.Email } elseif ($editor -and $editor.LookupValue) { $editor.LookupValue } else { '' }
                        $sizeBytes = GetProp $item.FieldValues 'File_x0020_Size'
                        $sizeKB = if ($sizeBytes) { [math]::Round([double]$sizeBytes / 1KB, 1) } else { 0 }

                        $rs.Rows.Add([ordered]@{
                            Status               = 'Live'
                            Name                 = $name
                            Location             = (Split-Path $path -Parent)
                            SizeKB               = $sizeKB
                            Date                 = if ($modified) { $modified.ToString('yyyy-MM-dd HH:mm') } else { '' }
                            ModifiedOrDeletedBy  = $editorName
                        })
                        if ($fileCount % 250 -eq 0) { QLog "  ...$fileCount live file(s) so far" }
                    }
                    QLog "Live files: $fileCount" 'OK'
                } else {
                    QLog 'Could not locate a Documents library — skipping live file enumeration.' 'WARN'
                }

                # ── Recycle bin (both stages) ────────────────────────────────
                foreach ($stage in @('FirstStage','SecondStage')) {
                    QLog "Enumerating Recycle Bin ($stage)..."
                    $binItems = if ($stage -eq 'FirstStage') { Get-PnPRecycleBinItem -FirstStage -ErrorAction Stop } else { Get-PnPRecycleBinItem -SecondStage -ErrorAction Stop }
                    $binItems = @($binItems)
                    $label = if ($stage -eq 'FirstStage') { 'Recycle Bin (user-recoverable)' } else { 'Recycle Bin (admin-recoverable only)' }
                    foreach ($bi in $binItems) {
                        $itemType = GetProp $bi 'ItemType'
                        if ($itemType -and $itemType -ne 'File') { continue }
                        $title      = GetProp $bi 'Title'
                        $dir        = GetProp $bi 'DirName'
                        $deletedBy  = GetProp $bi 'DeletedByEmail'
                        if (-not $deletedBy) { $deletedBy = GetProp $bi 'DeletedBy' }
                        $deletedDt  = GetProp $bi 'DeletedDate'
                        $size       = GetProp $bi 'Size'
                        $sizeKB = if ($size) { [math]::Round([double]$size / 1KB, 1) } else { 0 }

                        $rs.Rows.Add([ordered]@{
                            Status               = $label
                            Name                 = $title
                            Location             = $dir
                            SizeKB               = $sizeKB
                            Date                 = if ($deletedDt) { ([datetime]$deletedDt).ToString('yyyy-MM-dd HH:mm') } else { '' }
                            ModifiedOrDeletedBy  = $deletedBy
                        })
                    }
                    QLog "$label items: $($binItems.Count)" 'OK'
                }

                QLog "Done — $($rs.Rows.Count) total row(s)." 'OK'

            } catch {
                $rs.FatalError = $_.Exception.Message
                QLog "Fatal: $($_.Exception.Message)" 'ERROR'
            } finally {
                try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}
                $rs.Done = $true
            }
        })
        [void]$script:sofPS.AddParameters(@{ adminUrl = $adminUrl; oneDriveUrl = $oneDriveUrl; yourUpn = $yourUpn; rs = $rs; logFilePath = $logFilePath })
        $script:sofHandle = $script:sofPS.BeginInvoke()

        $rs2 = $rs

        $script:sofTimer = New-Object System.Windows.Forms.Timer
        $script:sofTimer.Interval = 300
        $script:sofTimer.Add_Tick({
            if ($script:tickBusy) { return }
            $script:tickBusy = $true
            try {
                if ($form.IsDisposed) { $script:sofTimer.Stop(); return }

                while ($rs2.LogQueue.Count -gt 0) {
                    $raw   = $rs2.LogQueue.Dequeue()
                    $level = if ($raw -match '^\[(\w+)\]') { $matches[1] } else { 'INFO' }
                    $msg   = $raw -replace '^\[\w+\] ', ''
                    Set-CtlText $lblStatus $msg
                    _RawLog "[$level] $msg"
                }
                while ($rs2.Rows.Count -gt 0) {
                    $row = $rs2.Rows[0]
                    $rs2.Rows.RemoveAt(0)
                    [void]$dt.Rows.Add($row.Status, $row.Name, $row.Location, [double]$row.SizeKB, $row.Date, $row.ModifiedOrDeletedBy)
                }
                if ($rs2.Done) {
                    $script:sofTimer.Stop()
                    if (-not $progress.IsDisposed) { $progress.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous; $progress.Value = 100 }
                    if ($rs2.FatalError) {
                        Set-CtlText $lblStatus "Failed — $($rs2.FatalError)"
                        _RawLog "=== Search failed: $($rs2.FatalError) ==="
                    } else {
                        Set-CtlText $lblStatus "Done — $($dt.Rows.Count) row(s). Use Filter to search, or Export CSV."
                        _RawLog "=== Search complete — $($dt.Rows.Count) row(s) ==="
                        if (-not $txtFilter.IsDisposed) { $txtFilter.Enabled = $true }
                        if (-not $btnExport.IsDisposed) { $btnExport.Enabled = ($dt.Rows.Count -gt 0) }
                    }
                    if (-not $btnRun.IsDisposed)   { $btnRun.Enabled = $true }
                    if (-not $btnClose.IsDisposed) { $btnClose.Enabled = $true }
                    try { $script:sofRunspace.Close(); $script:sofRunspace.Dispose() } catch {}
                }
            } catch {
                _RawLog "Tick handler error: $($_.Exception.Message)"
                _RawLog $_.ScriptStackTrace
                try { $script:sofTimer.Stop() } catch {}
                try { if ($script:sofRunspace) { $script:sofRunspace.Close(); $script:sofRunspace.Dispose() } } catch {}
                if (-not $form.IsDisposed) {
                    Set-CtlText $lblStatus "Failed — $($_.Exception.Message)"
                    if (-not $btnRun.IsDisposed)   { $btnRun.Enabled = $true }
                    if (-not $btnClose.IsDisposed) { $btnClose.Enabled = $true }
                }
            } finally { $script:tickBusy = $false }
        }.GetNewClosure())
        $script:sofTimer.Start()
      } catch {
        _RawLog "Run-click error: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Failed to start search: $($_.Exception.Message)", 'Error', 'OK', 'Error') | Out-Null
        $btnRun.Enabled = $true; $btnExport.Enabled = $true; $btnClose.Enabled = $true
      }
    }.GetNewClosure())

    $form.Add_FormClosing({
        param($s, $e)
        if ($script:sofTimer -and $script:sofTimer.Enabled) {
            $r = [System.Windows.Forms.MessageBox]::Show('Search is still running. Close anyway?', 'In Progress', 'YesNo', 'Warning')
            if ($r -eq [System.Windows.Forms.DialogResult]::No) { $e.Cancel = $true; return }
        }
        if ($script:sofTimer)    { try { $script:sofTimer.Stop();    $script:sofTimer.Dispose() }    catch {} }
        if ($script:sofPS)       { try { $script:sofPS.Stop();       $script:sofPS.Dispose() }        catch {} }
        if ($script:sofRunspace) { try { $script:sofRunspace.Close(); $script:sofRunspace.Dispose() } catch {} }
    }.GetNewClosure())

    $form.Add_Shown({ $form.BringToFront(); $form.Activate() }.GetNewClosure())
    [System.Windows.Forms.Application]::Run($form)
}

_RawLog '=== Search-OneDriveFiles.ps1 starting ==='
try { Show-SearchOneDriveUI }
catch {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    [System.Windows.Forms.MessageBox]::Show("Failed to start: $($_.Exception.Message)", 'Launch Error', 'OK', 'Error') | Out-Null
}
