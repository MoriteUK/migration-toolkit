#Requires -Version 7.0
<#
.SYNOPSIS
    Update-UPN.ps1 — Change the UPN domain suffix for Entra ID users.
    Connects to the tenant, loads all verified domains into dropdowns,
    then finds and updates matching users.

.NOTES
    Requires : Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement,
               Microsoft.Graph.Users
    Scopes   : User.ReadWrite.All, Domain.Read.All
    On-prem synced users are shown in amber and skipped — change their UPN in AD instead.
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$OldDomain,

    [Parameter(Mandatory=$false)]
    [string]$NewDomain,

    [switch]$WhatIf
)

$script:RootDir = $PSScriptRoot

# Disable WAM broker BEFORE any Graph modules load
$env:AZURE_IDENTITY_DISABLE_BROKER = 'true'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$_logDir = Join-Path $script:RootDir 'logs'
if (-not (Test-Path $_logDir)) { New-Item -ItemType Directory -Path $_logDir -Force | Out-Null }
$script:LogFile = Join-Path $_logDir "update-upn-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

function _RawLog {
    param([string]$Msg)
    "$(Get-Date -Format 'HH:mm:ss.fff') $Msg" | Add-Content -Path $script:LogFile -Encoding UTF8 -ErrorAction SilentlyContinue
}

_RawLog "=== Update-UPN.ps1 started  PID=$PID  PSVersion=$($PSVersionTable.PSVersion) ==="

$libPath = Join-Path $script:RootDir 'lib.ps1'
$_libLoaded = $false
if (Test-Path $libPath) {
    try { . $libPath; $_libLoaded = $true; _RawLog 'lib.ps1 loaded OK' }
    catch { _RawLog "lib.ps1 LOAD ERROR: $($_.Exception.Message)" }
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
                'OK'    { [System.Drawing.Color]::FromArgb(65, 195, 110)  }
                'WARN'  { [System.Drawing.Color]::FromArgb(220, 165, 45)  }
                'ERROR' { [System.Drawing.Color]::FromArgb(225, 80, 80)   }
                default { [System.Drawing.Color]::FromArgb(120, 155, 220) }
            }
            $script:rtbLog.SelectionColor = $lc; $script:rtbLog.AppendText("[$Level] ")
            $script:rtbLog.SelectionColor = [System.Drawing.Color]::FromArgb(205, 212, 230)
            $script:rtbLog.AppendText("$Msg`n"); $script:rtbLog.ScrollToCaret()
        }
    }
}

function Show-UpdateUpnUI {

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'Update UPNs'
    $form.ClientSize      = [System.Drawing.Size]::new(680, 820)
    $form.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.BackColor       = $clrBg
    $form.Font            = $FontBody
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.MaximizeBox     = $false
    $_ico = Join-Path $PSScriptRoot 'FlyMigration.ico'
    if (Test-Path $_ico) { $form.Icon = [System.Drawing.Icon]::new($_ico) }

    # ── Header ────────────────────────────────────────────────────────────────
    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Size = [System.Drawing.Size]::new(680, 56); $hdr.Dock = [System.Windows.Forms.DockStyle]::Top
    $hdr.BackColor = $clrAccent; $form.Controls.Add($hdr)
    $_hdrX = Add-HeaderLogo $hdr 36
    $hdrLbl = New-Object System.Windows.Forms.Label
    $hdrLbl.Text = '  Update UPNs'; $hdrLbl.Font = $FontTitle
    $hdrLbl.ForeColor = [System.Drawing.Color]::White
    $hdrLbl.Location  = [System.Drawing.Point]::new($_hdrX, 0)
    $hdrLbl.Size      = [System.Drawing.Size]::new(580, 56)
    $hdrLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdr.Controls.Add($hdrLbl)

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
    $btnClose.Add_Click({
        Write-Log 'Closing form - disconnecting from Graph...'
        try {
            if (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue) {
                Disconnect-MgGraph -Confirm:$false -ErrorAction SilentlyContinue
                Write-Log 'Disconnected from Microsoft Graph' 'OK'
            }
        } catch {
            Write-Log "Disconnect warning: $($_.Exception.Message)" 'WARN'
        }
        $form.Close()
    }.GetNewClosure())
    $footer.Controls.Add($btnClose)
    $footer.Add_SizeChanged({ $btnClose.Left = $footer.Width - 100 }.GetNewClosure())

    # ── Card ──────────────────────────────────────────────────────────────────
    $card = New-Object System.Windows.Forms.Panel
    $card.Location  = [System.Drawing.Point]::new(12, 66)
    $card.Size      = [System.Drawing.Size]::new(656, 480)
    $card.BackColor = $clrPanel
    $form.Controls.Add($card)

    $lx = 14; $y = 14

    # ── Connect row ───────────────────────────────────────────────────────────
    $lblConnCap = New-Object System.Windows.Forms.Label
    $lblConnCap.Text = 'TENANT CONNECTION'; $lblConnCap.Font = $FontCap; $lblConnCap.ForeColor = $clrMuted
    $lblConnCap.Location = [System.Drawing.Point]::new($lx, $y); $lblConnCap.AutoSize = $true
    $card.Controls.Add($lblConnCap); $y += 18

    $btnConnect = New-Object System.Windows.Forms.Button
    $btnConnect.Text = 'Connect to Domain'; $btnConnect.Location = [System.Drawing.Point]::new($lx, $y)
    $btnConnect.Size = [System.Drawing.Size]::new(180, 28); $btnConnect.Font = $FontBold
    $btnConnect.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnConnect.FlatAppearance.BorderSize = 0
    $btnConnect.BackColor = [System.Drawing.Color]::FromArgb(225, 228, 238); $btnConnect.ForeColor = $clrText
    $btnConnect.Cursor = [System.Windows.Forms.Cursors]::Hand
    $card.Controls.Add($btnConnect)

    $lblConnStatus = New-Object System.Windows.Forms.Label
    $lblConnStatus.Text = 'Not connected'
    $lblConnStatus.ForeColor = $clrMuted; $lblConnStatus.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $lblConnStatus.Location = [System.Drawing.Point]::new(202, $y + 8); $lblConnStatus.AutoSize = $true
    $card.Controls.Add($lblConnStatus); $y += 38

    # Separator
    $sep1 = New-Object System.Windows.Forms.Label
    $sep1.Location = [System.Drawing.Point]::new($lx, $y); $sep1.Size = [System.Drawing.Size]::new(628, 1)
    $sep1.BackColor = $clrBorder; $card.Controls.Add($sep1); $y += 12

    # ── Source domain ─────────────────────────────────────────────────────────
    $lblSrcCap = New-Object System.Windows.Forms.Label
    $lblSrcCap.Text = 'SOURCE DOMAIN  —  find users with this UPN suffix'
    $lblSrcCap.Font = $FontCap; $lblSrcCap.ForeColor = $clrMuted
    $lblSrcCap.Location = [System.Drawing.Point]::new($lx, $y); $lblSrcCap.AutoSize = $true
    $card.Controls.Add($lblSrcCap); $y += 18

    $cbSrc = New-Object System.Windows.Forms.ComboBox
    $cbSrc.Location = [System.Drawing.Point]::new($lx, $y); $cbSrc.Size = [System.Drawing.Size]::new(440, 24)
    $cbSrc.Font = $FontBody; $cbSrc.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cbSrc.Enabled = $false
    $card.Controls.Add($cbSrc); $y += 34

    # ── Target domain ─────────────────────────────────────────────────────────
    $lblTgtCap = New-Object System.Windows.Forms.Label
    $lblTgtCap.Text = 'TARGET DOMAIN  —  UPNs will be changed to this suffix'
    $lblTgtCap.Font = $FontCap; $lblTgtCap.ForeColor = $clrMuted
    $lblTgtCap.Location = [System.Drawing.Point]::new($lx, $y); $lblTgtCap.AutoSize = $true
    $card.Controls.Add($lblTgtCap); $y += 18

    $cbTgt = New-Object System.Windows.Forms.ComboBox
    $cbTgt.Location = [System.Drawing.Point]::new($lx, $y); $cbTgt.Size = [System.Drawing.Size]::new(440, 24)
    $cbTgt.Font = $FontBody; $cbTgt.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cbTgt.Enabled = $false
    $card.Controls.Add($cbTgt); $y += 34

    # Separator
    $sep2 = New-Object System.Windows.Forms.Label
    $sep2.Location = [System.Drawing.Point]::new($lx, $y); $sep2.Size = [System.Drawing.Size]::new(628, 1)
    $sep2.BackColor = $clrBorder; $card.Controls.Add($sep2); $y += 12

    # ── Find row ──────────────────────────────────────────────────────────────
    $btnFind = New-Object System.Windows.Forms.Button
    $btnFind.Text = 'Find Users'; $btnFind.Location = [System.Drawing.Point]::new($lx, $y)
    $btnFind.Size = [System.Drawing.Size]::new(120, 28); $btnFind.Font = $FontBold
    $btnFind.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnFind.FlatAppearance.BorderSize = 0
    $btnFind.BackColor = [System.Drawing.Color]::FromArgb(225, 228, 238); $btnFind.ForeColor = $clrText
    $btnFind.Cursor = [System.Windows.Forms.Cursors]::Hand; $btnFind.Enabled = $false
    $card.Controls.Add($btnFind)

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = 'Export List'; $btnExport.Location = [System.Drawing.Point]::new($lx + 126, $y)
    $btnExport.Size = [System.Drawing.Size]::new(120, 28); $btnExport.Font = $FontBold
    $btnExport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnExport.FlatAppearance.BorderSize = 0
    $btnExport.BackColor = [System.Drawing.Color]::FromArgb(225, 228, 238); $btnExport.ForeColor = $clrText
    $btnExport.Cursor = [System.Windows.Forms.Cursors]::Hand; $btnExport.Enabled = $false
    $card.Controls.Add($btnExport)

    $lblFindStatus = New-Object System.Windows.Forms.Label
    $lblFindStatus.Text = 'Connect first, then select domains and click Find Objects.'
    $lblFindStatus.ForeColor = $clrMuted; $lblFindStatus.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $lblFindStatus.Location = [System.Drawing.Point]::new(258, $y + 8); $lblFindStatus.AutoSize = $true
    $card.Controls.Add($lblFindStatus); $y += 38

    # ── Objects list ───────────────────────────────────────────────────────────
    $lblListCap = New-Object System.Windows.Forms.Label
    $lblListCap.Text = 'OBJECTS TO AUDIT / UPDATE  (amber = on-premises synced users must be changed in AD)'
    $lblListCap.Font = $FontCap; $lblListCap.ForeColor = $clrMuted
    $lblListCap.Location = [System.Drawing.Point]::new($lx, $y); $lblListCap.AutoSize = $true
    $card.Controls.Add($lblListCap); $y += 18

    $lv = New-Object System.Windows.Forms.ListView
    $lv.Location      = [System.Drawing.Point]::new($lx, $y)
    $lv.Size          = [System.Drawing.Size]::new(628, 165)
    $lv.View          = [System.Windows.Forms.View]::Details
    $lv.FullRowSelect = $true
    $lv.GridLines     = $true
    $lv.Font          = $FontMono
    $lv.BorderStyle   = [System.Windows.Forms.BorderStyle]::FixedSingle
    $lv.CheckBoxes    = $true
    [void]$lv.Columns.Add('Type', 110)
    [void]$lv.Columns.Add('Display Name', 185)
    [void]$lv.Columns.Add('Current Address', 210)
    [void]$lv.Columns.Add('New Address', 210)
    $card.Controls.Add($lv); $y += 170

    # Separator
    $sep3 = New-Object System.Windows.Forms.Label
    $sep3.Location = [System.Drawing.Point]::new($lx, $y); $sep3.Size = [System.Drawing.Size]::new(628, 1)
    $sep3.BackColor = $clrBorder; $card.Controls.Add($sep3); $y += 12

    # ── WhatIf + Update ───────────────────────────────────────────────────────
    $chkWhatIf = New-Object System.Windows.Forms.CheckBox
    $chkWhatIf.Text = 'WhatIf  (preview only — no changes will be made)'
    $chkWhatIf.Location = [System.Drawing.Point]::new($lx, $y + 6)
    $chkWhatIf.AutoSize = $true; $chkWhatIf.ForeColor = $clrText
    $card.Controls.Add($chkWhatIf)

    $btnUpdate = New-Object System.Windows.Forms.Button
    $btnUpdate.Text = 'Update UPNs'; $btnUpdate.Location = [System.Drawing.Point]::new(484, $y)
    $btnUpdate.Size = [System.Drawing.Size]::new(158, 34); $btnUpdate.Font = $FontBold
    $btnUpdate.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnUpdate.FlatAppearance.BorderSize = 0
    $btnUpdate.BackColor = $clrRed; $btnUpdate.ForeColor = [System.Drawing.Color]::White
    $btnUpdate.Cursor = [System.Windows.Forms.Cursors]::Hand; $btnUpdate.Enabled = $false
    $card.Controls.Add($btnUpdate)

    # ── Progress + Log ────────────────────────────────────────────────────────
    $progress = New-Object System.Windows.Forms.ProgressBar
    $progress.Location = [System.Drawing.Point]::new(12, $card.Bottom + 8)
    $progress.Size     = [System.Drawing.Size]::new(656, 8)
    $progress.Style    = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $form.Controls.Add($progress)

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

    # ── Shared state — List is a ref type; both closures share the same object ─
    $foundUsers = [System.Collections.Generic.List[pscustomobject]]::new()

    # ── ListView Column Sorting ───────────────────────────────────────────────
    $script:sortColumn = -1
    $script:sortOrder = [System.Windows.Forms.SortOrder]::Ascending

    $lv.Add_ColumnClick({
        param($sender, $e)
        $columnIndex = $e.Column

        # Toggle sort order if clicking the same column, otherwise reset to ascending
        if ($script:sortColumn -eq $columnIndex) {
            if ($script:sortOrder -eq [System.Windows.Forms.SortOrder]::Ascending) {
                $script:sortOrder = [System.Windows.Forms.SortOrder]::Descending
            } else {
                $script:sortOrder = [System.Windows.Forms.SortOrder]::Ascending
            }
        } else {
            $script:sortColumn = $columnIndex
            $script:sortOrder = [System.Windows.Forms.SortOrder]::Ascending
        }

        # Create array of items with their checked state
        $items = @()
        foreach ($item in $lv.Items) {
            $items += [pscustomobject]@{
                Item = $item
                Checked = $item.Checked
                Values = @(
                    $item.Text
                    $item.SubItems[1].Text
                    $item.SubItems[2].Text
                    $item.SubItems[3].Text
                )
            }
        }

        # Sort the items
        $sortedItems = if ($script:sortOrder -eq [System.Windows.Forms.SortOrder]::Ascending) {
            $items | Sort-Object { $_.Values[$columnIndex] }
        } else {
            $items | Sort-Object { $_.Values[$columnIndex] } -Descending
        }

        # Rebuild the ListView
        $lv.BeginUpdate()
        $lv.Items.Clear()
        foreach ($obj in $sortedItems) {
            $obj.Item.Checked = $obj.Checked
            [void]$lv.Items.Add($obj.Item)
        }
        $lv.EndUpdate()
    }.GetNewClosure())

    # ── Connect & Load Domains ────────────────────────────────────────────────
    $btnConnect.Add_Click({
        $btnConnect.Enabled = $false; $cbSrc.Enabled = $false; $cbTgt.Enabled = $false
        $btnFind.Enabled = $false; $btnUpdate.Enabled = $false
        $lblConnStatus.Text = 'Connecting...'; $lblConnStatus.ForeColor = $clrMuted
        [System.Windows.Forms.Application]::DoEvents()
        Write-Log 'Connecting to Microsoft Graph...'

        try {
            foreach ($m in @('Microsoft.Graph.Authentication','Microsoft.Graph.Identity.DirectoryManagement','Microsoft.Graph.Users','Microsoft.Graph.Groups')) {
                if (-not (Get-Module -ListAvailable -Name $m -ErrorAction SilentlyContinue)) {
                    Write-Log "Module not installed: $m  (run: Install-Module $m -Scope CurrentUser)" 'WARN'
                } else {
                    Import-Module $m -ErrorAction SilentlyContinue
                }
            }

            # Clear any cached tokens from previous sessions
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}

            $scopes = @('User.ReadWrite.All', 'Domain.Read.All', 'Group.Read.All')

            Write-Log 'Starting device code authentication...'

            try {
                # Create a script file to run device code auth in a visible console
                $authScript = Join-Path $env:TEMP "mg-auth-$(Get-Date -Format 'yyyyMMddHHmmss').ps1"
                $authScriptContent = @"
`$env:AZURE_IDENTITY_DISABLE_BROKER = 'true'
Write-Host ''
Write-Host '═══════════════════════════════════════════════════════════' -ForegroundColor Cyan
Write-Host '                  MICROSOFT GRAPH SIGN-IN' -ForegroundColor Cyan
Write-Host '═══════════════════════════════════════════════════════════' -ForegroundColor Cyan
Write-Host ''
Write-Host 'Connecting to Microsoft Graph...' -ForegroundColor Yellow
Write-Host ''

try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Connect-MgGraph -Scopes 'User.ReadWrite.All','Domain.Read.All','Group.Read.All' -UseDeviceCode -TenantId 'organizations' -NoWelcome -ErrorAction Stop

    Write-Host ''
    Write-Host '═══════════════════════════════════════════════════════════' -ForegroundColor Green
    Write-Host '              AUTHENTICATION SUCCESSFUL!' -ForegroundColor Green
    Write-Host '═══════════════════════════════════════════════════════════' -ForegroundColor Green
    Write-Host ''
    Write-Host 'You can close this window now.' -ForegroundColor Green
    Write-Host ''

    # Keep window open for 3 seconds so user sees success message
    Start-Sleep -Seconds 3
    exit 0
} catch {
    Write-Host ''
    Write-Host '═══════════════════════════════════════════════════════════' -ForegroundColor Red
    Write-Host '              AUTHENTICATION FAILED!' -ForegroundColor Red
    Write-Host '═══════════════════════════════════════════════════════════' -ForegroundColor Red
    Write-Host ''
    Write-Host "Error: `$(`$_.Exception.Message)" -ForegroundColor Red
    Write-Host ''
    Write-Host 'Press any key to close this window...' -ForegroundColor Yellow
    `$null = `$Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    exit 1
}
"@
                Set-Content -Path $authScript -Value $authScriptContent -Encoding UTF8

                Write-Log "Launching authentication window..."

                # Launch PowerShell with visible console window
                $proc = Start-Process -FilePath 'pwsh.exe' -ArgumentList @(
                    '-NoProfile'
                    '-ExecutionPolicy', 'Bypass'
                    '-File', "`"$authScript`""
                ) -Wait -PassThru -WindowStyle Normal

                # Clean up temp script
                Remove-Item -Path $authScript -Force -ErrorAction SilentlyContinue

                if ($proc.ExitCode -eq 0) {
                    Write-Log 'Authentication succeeded.' 'OK'

                    # Verify the connection worked in this session too
                    $ctx = Get-MgContext -ErrorAction SilentlyContinue
                    if ($ctx) {
                        Write-Log "Verified: Connected to tenant $($ctx.TenantId)" 'OK'
                    } else {
                        Write-Log 'Authentication succeeded in console but connection not found in this session.' 'WARN'
                        Write-Log 'Attempting to reconnect in this session...'
                        # The auth was in another process - we need to auth in this one too
                        # But now the user is already authenticated, so this should use cached token
                        Connect-MgGraph -Scopes $scopes -TenantId 'organizations' -NoWelcome -ErrorAction Stop
                        Write-Log 'Reconnected successfully.' 'OK'
                    }
                } else {
                    throw 'Authentication failed in console window (exit code: {0})' -f $proc.ExitCode
                }

            } catch {
                Write-Log "Authentication failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'
                throw
            }

            $ctx = Get-MgContext -ErrorAction Stop
            Write-Log "Connected — TenantId: $($ctx.TenantId)  Account: $($ctx.Account)" 'OK'
        } catch {
            Write-Log "Connection failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'
            $lblConnStatus.Text = 'Connection failed — see log.'; $lblConnStatus.ForeColor = [System.Drawing.Color]::FromArgb(180,30,30)
            $btnConnect.Enabled = $true; return
        }

        try {
            Write-Log 'Loading tenant domains...'
            $domains = @(Get-MgDomain -All -ErrorAction Stop | Where-Object { $_.IsVerified } | Sort-Object Id)
            Write-Log "$($domains.Count) verified domain(s): $($domains.Id -join ', ')"

            $cbSrc.Items.Clear(); $cbTgt.Items.Clear()
            foreach ($d in $domains) { [void]$cbSrc.Items.Add($d.Id); [void]$cbTgt.Items.Add($d.Id) }

            # Auto-select: source = initial .onmicrosoft.com, target = default or first custom domain
            $initialDomain = $domains | Where-Object { $_.IsInitial } | Select-Object -First 1
            $defaultDomain  = $domains | Where-Object { $_.IsDefault -and -not $_.IsInitial } | Select-Object -First 1
            $firstCustom    = $domains | Where-Object { -not $_.IsInitial } | Select-Object -First 1

            if ($initialDomain)              { $cbSrc.SelectedItem = $initialDomain.Id }
            elseif ($cbSrc.Items.Count -gt 0){ $cbSrc.SelectedIndex = 0 }

            if ($defaultDomain)              { $cbTgt.SelectedItem = $defaultDomain.Id }
            elseif ($firstCustom)            { $cbTgt.SelectedItem = $firstCustom.Id }
            elseif ($cbTgt.Items.Count -gt 0){ $cbTgt.SelectedIndex = 0 }

            $cbSrc.Enabled = $true; $cbTgt.Enabled = $true; $btnFind.Enabled = $true
            $lblConnStatus.Text = "Connected  —  $($domains.Count) domain(s) loaded"
            $lblConnStatus.ForeColor = [System.Drawing.Color]::FromArgb(20, 140, 60)
            Write-Log "Source pre-selected: $($cbSrc.SelectedItem)   Target pre-selected: $($cbTgt.SelectedItem)" 'OK'
        } catch {
            Write-Log "Domain load failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'
            $lblConnStatus.Text = 'Connected but domain load failed — see log.'
            $lblConnStatus.ForeColor = [System.Drawing.Color]::FromArgb(180,30,30)
        }
        $btnConnect.Enabled = $true
    }.GetNewClosure())

    # ── Find Objects ───────────────────────────────────────────────────────────
    $btnFind.Add_Click({
        $src = "$($cbSrc.SelectedItem)"
        $tgt = "$($cbTgt.SelectedItem)"

        if (-not $src -or -not $tgt) {
            [System.Windows.Forms.MessageBox]::Show('Please select both source and target domains.','Missing Selection','OK','Warning') | Out-Null; return
        }
        if ($src -eq $tgt) {
            [System.Windows.Forms.MessageBox]::Show('Source and target domains are the same.','Invalid Selection','OK','Warning') | Out-Null; return
        }

        $btnFind.Enabled = $false; $btnUpdate.Enabled = $false
        Write-Log "Finding tenant objects matching '@$src'..."
        $lblFindStatus.Text = 'Searching...'; [System.Windows.Forms.Application]::DoEvents()

        try {
            $lv.Items.Clear(); $foundUsers.Clear()
            $escapedSrc = [regex]::Escape($src)
            $clrAmber = [System.Drawing.Color]::FromArgb(160, 100, 0)
            $synced = 0

            Write-Log 'Searching users...'
            $users = @(Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,OnPremisesSyncEnabled -ErrorAction Stop |
                Where-Object {
                    $_.UserPrincipalName -and $_.UserPrincipalName -like "*@$src"
                })
            Write-Log "User search returned $($users.Count) user(s)."

            foreach ($u in $users) {
                $currentAddress = if ($u.UserPrincipalName -and $u.UserPrincipalName -like "*@$src") {
                    $u.UserPrincipalName
                } else {
                    ''
                }

                $newAddress = if ($currentAddress) { $currentAddress -replace [regex]::Escape("@$src"), "@$tgt" } else { '' }
                $item = New-Object System.Windows.Forms.ListViewItem('User')
                [void]$item.SubItems.Add($u.DisplayName)
                [void]$item.SubItems.Add($currentAddress)
                [void]$item.SubItems.Add($newAddress)
                $isSynced = ($u.OnPremisesSyncEnabled -eq $true)
                if ($isSynced) { $item.ForeColor = $clrAmber; $synced++ }
                else { $item.Checked = $true }
                [void]$lv.Items.Add($item)
                $foundUsers.Add([pscustomobject]@{
                    Type             = 'User'
                    Id               = $u.Id
                    DisplayName      = $u.DisplayName
                    CurrentAddress   = $currentAddress
                    NewAddress       = $newAddress
                    OnPremisesSynced = $isSynced
                })
            }

            Write-Log 'Searching groups...'
            $groups = @(Get-MgGroup -All -Property Id,DisplayName,Mail,MailNickname,GroupTypes,SecurityEnabled,MailEnabled,ResourceProvisioningOptions -ErrorAction Stop |
                Where-Object {
                    $_.Mail -and $_.Mail -like "*@$src"
                })
            Write-Log "Group search returned $($groups.Count) group(s)."

            foreach ($g in $groups) {
                $currentAddress = if ($g.Mail -and $g.Mail -like "*@$src") {
                    $g.Mail
                } else {
                    ''
                }

                $newAddress = if ($currentAddress) { $currentAddress -replace [regex]::Escape("@$src"), "@$tgt" } else { '' }

                $groupType = if ($g.ResourceProvisioningOptions -and $g.ResourceProvisioningOptions -contains 'Team') {
                    'Team'
                } elseif ($g.GroupTypes -and $g.GroupTypes -contains 'Unified') {
                    'M365 Group'
                } elseif ($g.MailEnabled -and -not $g.SecurityEnabled) {
                    'Distribution List'
                } elseif ($g.MailEnabled -and $g.SecurityEnabled) {
                    'Mail-Enabled Security Group'
                } else {
                    'Group'
                }

                $item = New-Object System.Windows.Forms.ListViewItem($groupType)
                [void]$item.SubItems.Add($g.DisplayName)
                [void]$item.SubItems.Add($currentAddress)
                [void]$item.SubItems.Add($newAddress)
                $item.Checked = $true
                [void]$lv.Items.Add($item)
                $foundUsers.Add([pscustomobject]@{
                    Type             = $groupType
                    Id               = $g.Id
                    DisplayName      = $g.DisplayName
                    CurrentAddress   = $currentAddress
                    NewAddress       = $newAddress
                    OnPremisesSynced = $false
                })
            }

            $cloudCount = @($foundUsers | Where-Object { $_.Type -eq 'User' -and -not $_.OnPremisesSynced }).Count
            $totalObjects = $foundUsers.Count
            $userCount = @($foundUsers | Where-Object { $_.Type -eq 'User' }).Count
            $groupCount = $totalObjects - $userCount

            if ($totalObjects -gt 0) {
                $msg = "$($totalObjects) object(s) found"
                if ($userCount -gt 0) {
                    $msg += " ($userCount users"
                    if ($groupCount -gt 0) { $msg += ", $groupCount groups" }
                    $msg += ')'
                }
                if ($synced -gt 0) { $msg += "  ($synced synced/amber - skipped)" }
                $lblFindStatus.Text = $msg
                $btnUpdate.Enabled = ($cloudCount -gt 0)
                $btnExport.Enabled = $true
                Write-Log "$($totalObjects) object(s)  ($cloudCount cloud-updatable users  |  $synced on-prem synced users)" 'OK'
                if ($synced -gt 0) { Write-Log 'Amber users must have their UPN changed in on-premises Active Directory.' 'WARN' }
            } else {
                $lblFindStatus.Text = "No objects found with '@$src' domain"
                $btnExport.Enabled = $false
                Write-Log "No objects found with '@$src' domain." 'WARN'
            }
        } catch {
            Write-Log "Search failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'
            $lblFindStatus.Text = 'Search failed — see log.'
            $btnExport.Enabled = $false
        } finally {
            $btnFind.Enabled = $true
        }
    }.GetNewClosure())

    # ── Export List ───────────────────────────────────────────────────────────
    $btnExport.Add_Click({
        if ($foundUsers.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('No objects to export. Click Find Objects first.','No Data','OK','Warning') | Out-Null
            return
        }

        Write-Log 'Export clicked — prompting for file location...'

        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Title = 'Export Object List'
        $sfd.Filter = 'CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt|All Files (*.*)|*.*'
        $sfd.FilterIndex = 1
        $sfd.DefaultExt = 'csv'
        $sfd.FileName = "Tenant-Object-Update-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
        $sfd.InitialDirectory = [Environment]::GetFolderPath('Desktop')

        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $exportPath = $sfd.FileName
                Write-Log "Exporting $($foundUsers.Count) object(s) to: $exportPath"

                if ($exportPath -like '*.csv') {
                    # Export as CSV
                    $foundUsers | Select-Object Type,DisplayName,CurrentAddress,NewAddress,OnPremisesSynced |
                        Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
                    Write-Log "Exported as CSV: $exportPath" 'OK'
                } else {
                    # Export as plain text
                    $sb = [System.Text.StringBuilder]::new()
                    [void]$sb.AppendLine("Tenant Object Update List")
                    [void]$sb.AppendLine("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
                    [void]$sb.AppendLine("Source Domain: @$($cbSrc.SelectedItem)")
                    [void]$sb.AppendLine("Target Domain: @$($cbTgt.SelectedItem)")
                    [void]$sb.AppendLine("=" * 80)
                    [void]$sb.AppendLine("")
                    [void]$sb.AppendLine("{0,-20} {1,-30} {2,-40} {3,-40} {4}" -f 'Type','Display Name','Current Address','New Address','Synced')
                    [void]$sb.AppendLine("-" * 150)
                    foreach ($u in $foundUsers) {
                        $syncFlag = if ($u.OnPremisesSynced) { 'Yes' } else { 'No' }
                        [void]$sb.AppendLine("{0,-20} {1,-30} {2,-40} {3,-40} {4}" -f $u.Type,$u.DisplayName,$u.CurrentAddress,$u.NewAddress,$syncFlag)
                    }
                    [void]$sb.AppendLine("")
                    [void]$sb.AppendLine("Total Objects: $($foundUsers.Count)")
                    $onPremUserCount = @($foundUsers | Where-Object { $_.Type -eq 'User' -and $_.OnPremisesSynced }).Count
                    $cloudUserCount = @($foundUsers | Where-Object { $_.Type -eq 'User' -and -not $_.OnPremisesSynced }).Count
                    [void]$sb.AppendLine("  Cloud-updatable Users: $cloudUserCount")
                    [void]$sb.AppendLine("  On-Premises Synced Users: $onPremUserCount")

                    Set-Content -Path $exportPath -Value $sb.ToString() -Encoding UTF8
                    Write-Log "Exported as text: $exportPath" 'OK'
                }

                # Show success and offer to open
                $openResult = [System.Windows.Forms.MessageBox]::Show(
                    "Successfully exported $($foundUsers.Count) user(s) to:`n`n$exportPath`n`nOpen the file now?",
                    'Export Complete',
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )

                if ($openResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                    Start-Process $exportPath
                }

            } catch {
                Write-Log "Export failed: $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'
                [System.Windows.Forms.MessageBox]::Show("Export failed:`n`n$($_.Exception.Message)",'Export Error','OK','Error') | Out-Null
            }
        } else {
            Write-Log 'Export cancelled by user.'
        }
    }.GetNewClosure())

    # ── Update UPNs ───────────────────────────────────────────────────────────
    $btnUpdate.Add_Click({
        if ($foundUsers.Count -eq 0) { return }
        $src        = "$($cbSrc.SelectedItem)"
        $tgt        = "$($cbTgt.SelectedItem)"
        $whatIf     = $chkWhatIf.Checked

        # Get only checked items from the ListView
        $checkedItems = @($lv.CheckedItems)
        if ($checkedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                'No accounts selected. Please check the accounts you want to update.',
                'No Selection', 'OK', 'Warning') | Out-Null; return
        }

        # Match checked items back to foundUsers by display name and current address
        $cloudUsers = @(foreach ($item in $checkedItems) {
            $displayName = $item.SubItems[1].Text
            $currentAddr = $item.SubItems[2].Text
            $foundUsers | Where-Object {
                $_.DisplayName -eq $displayName -and
                $_.CurrentAddress -eq $currentAddr -and
                $_.Type -eq 'User' -and
                -not $_.OnPremisesSynced
            } | Select-Object -First 1
        })

        if ($cloudUsers.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                'No cloud users selected. Only cloud users can be updated (on-premises synced users must be changed in Active Directory).',
                'No Cloud Users', 'OK', 'Warning') | Out-Null; return
        }

        if (-not $whatIf) {
            $dlgC = New-Object System.Windows.Forms.Form
            $dlgC.Text = 'Confirm UPN Update'; $dlgC.ClientSize = [System.Drawing.Size]::new(480, 196)
            $dlgC.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
            $dlgC.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
            $dlgC.MaximizeBox = $false; $dlgC.MinimizeBox = $false; $dlgC.BackColor = $clrBg

            $lblCMsg = New-Object System.Windows.Forms.Label
            $lblCMsg.Text = "Permanently update UPNs for $($cloudUsers.Count) cloud user(s)?`n`n@$src  →  @$tgt`n`nType YES to confirm:"
            $lblCMsg.Location = [System.Drawing.Point]::new(16, 14); $lblCMsg.Size = [System.Drawing.Size]::new(448, 72)
            $lblCMsg.ForeColor = $clrText; $dlgC.Controls.Add($lblCMsg)

            $tbC = New-Object System.Windows.Forms.TextBox
            $tbC.Location = [System.Drawing.Point]::new(16, 92); $tbC.Size = [System.Drawing.Size]::new(448, 24)
            $tbC.Font = $FontBody; $tbC.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
            $dlgC.Controls.Add($tbC)

            $btnCP = New-Object System.Windows.Forms.Button
            $btnCP.Text = 'Proceed'; $btnCP.Location = [System.Drawing.Point]::new(280, 130)
            $btnCP.Size = [System.Drawing.Size]::new(90, 30); $btnCP.Font = $FontBold
            $btnCP.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnCP.FlatAppearance.BorderSize = 0
            $btnCP.BackColor = $clrRed; $btnCP.ForeColor = [System.Drawing.Color]::White
            $btnCP.DialogResult = [System.Windows.Forms.DialogResult]::OK; $dlgC.Controls.Add($btnCP)

            $btnCC = New-Object System.Windows.Forms.Button
            $btnCC.Text = 'Cancel'; $btnCC.Location = [System.Drawing.Point]::new(378, 130)
            $btnCC.Size = [System.Drawing.Size]::new(90, 30); $btnCC.Font = $FontBold
            $btnCC.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnCC.FlatAppearance.BorderSize = 0
            $btnCC.BackColor = [System.Drawing.Color]::FromArgb(225,228,238); $btnCC.ForeColor = $clrText
            $btnCC.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $dlgC.Controls.Add($btnCC)
            $dlgC.AcceptButton = $btnCP; $dlgC.CancelButton = $btnCC

            $dlgResult = $dlgC.ShowDialog()
            if ($dlgResult -ne [System.Windows.Forms.DialogResult]::OK -or $tbC.Text -ne 'YES') {
                Write-Log 'Update cancelled.' 'WARN'; return
            }
            Write-Log 'Confirmed — proceeding with live update.'
        }

        $btnUpdate.Enabled = $false; $btnFind.Enabled = $false
        $btnConnect.Enabled = $false; $btnClose.Enabled = $false
        $progress.Value = 0
        Write-Log "=== UPN Update started$(if ($whatIf) {' [WHATIF]'}) ==="
        Write-Log "@$src  →  @$tgt  ($($cloudUsers.Count) cloud user(s))"

        $ok = 0; $fail = 0; $total = $cloudUsers.Count

        for ($i = 0; $i -lt $total; $i++) {
            $u = $cloudUsers[$i]
            try {
                if ($whatIf) {
                    Write-Log "WhatIf: $($u.CurrentAddress)  →  $($u.NewAddress)  [$($u.DisplayName)]" 'WARN'
                } else {
                    Update-MgUser -UserId $u.Id -UserPrincipalName $u.NewAddress -ErrorAction Stop
                    Write-Log "UPDATED  $($u.CurrentAddress)  →  $($u.NewAddress)  [$($u.DisplayName)]" 'OK'
                }
                $ok++
            } catch {
                Write-Log "FAILED  $($u.CurrentAddress): $($_.Exception.Message.Split([Environment]::NewLine)[0])" 'ERROR'
                $fail++
            }
            $progress.Value = [int](($i + 1) / $total * 100)
            [System.Windows.Forms.Application]::DoEvents()
        }

        $status = if ($fail -eq 0) { 'OK' } else { 'WARN' }
        Write-Log "=== Completed: updated $ok  |  failed $fail ===" $status
        $btnUpdate.Enabled = ($foundUsers.Where({ -not $_.OnPremisesSynced }).Count -gt 0)
        $btnFind.Enabled = $true; $btnConnect.Enabled = $true; $btnClose.Enabled = $true
    }.GetNewClosure())

    $form.Add_Shown({
        $form.BringToFront(); $form.Activate()
        Write-Log '=== Update UPNs ready ==='
        Write-Log 'Click "Connect & Load Domains" to sign in to the target tenant.'
        Write-Log 'The dropdowns will populate from the domains verified in that tenant.'
    }.GetNewClosure())

    [System.Windows.Forms.Application]::Run($form)

    # Final cleanup: attempt to disconnect from any remaining connected tenants
    Write-Log 'Final cleanup - disconnecting from any remaining sessions'
    try {
        if (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue) {
            Disconnect-MgGraph -Confirm:$false -ErrorAction SilentlyContinue
            Write-Log 'Final disconnect: Disconnect-MgGraph' 'OK'
        }
    } catch { Write-Log "Final disconnect-MgGraph: $($_.Exception.Message)" 'WARN' }
    try { if (Get-Command Disconnect-ExchangeOnline -ErrorAction SilentlyContinue) { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue; Write-Log 'Final disconnect: Disconnect-ExchangeOnline' 'OK' } } catch {}
    try { if (Get-Command Disconnect-PnPOnline -ErrorAction SilentlyContinue)    { Disconnect-PnPOnline -ErrorAction SilentlyContinue; Write-Log 'Final disconnect: Disconnect-PnPOnline' 'OK' } } catch {}
    try { if (Get-Command Disconnect-SPOService -ErrorAction SilentlyContinue)    { Disconnect-SPOService -ErrorAction SilentlyContinue; Write-Log 'Final disconnect: Disconnect-SPOService' 'OK' } } catch {}
    try { if (Get-Command Disconnect-MsolService -ErrorAction SilentlyContinue)   { Disconnect-MsolService -ErrorAction SilentlyContinue; Write-Log 'Final disconnect: Disconnect-MsolService' 'OK' } } catch {}
}

if ($OldDomain -and $NewDomain) {
    # ── Headless mode ──────────────────────────────────────────────────────────
    $old = $OldDomain.TrimStart('@')
    $new = $NewDomain.TrimStart('@')
    Write-Host "=== Update-UPN  @$old  →  @$new$(if ($WhatIf) { '  [WhatIf]' }) ==="
    Write-Host 'Loading Microsoft Graph modules...'

    foreach ($mod in @('Microsoft.Graph.Authentication','Microsoft.Graph.Users')) {
        if (-not (Get-Module -ListAvailable -Name $mod -ErrorAction SilentlyContinue)) {
            Write-Host "ERROR: Module not installed: $mod"
            Write-Host 'Install with: Install-Module Microsoft.Graph -Scope CurrentUser'
            exit 1
        }
        Import-Module $mod -ErrorAction Stop
    }

    Write-Host 'Connecting to Microsoft Graph (device code)...'
    Connect-MgGraph -Scopes 'User.ReadWrite.All' -UseDeviceCode -TenantId 'organizations' -NoWelcome -ErrorAction Stop
    Write-Host 'Connected.'

    Write-Host "Searching for users with UPN suffix @$old ..."
    $users = @(Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,OnPremisesSyncEnabled -ErrorAction Stop |
        Where-Object { $_.UserPrincipalName -like "*@$old" })
    Write-Host "Found $($users.Count) user(s)."

    $cloudUsers = @($users | Where-Object { $_.OnPremisesSyncEnabled -ne $true })
    $syncedCount = $users.Count - $cloudUsers.Count
    if ($syncedCount -gt 0) {
        Write-Host "Skipping $syncedCount on-premises synced user(s) — update their UPN in Active Directory instead."
    }

    if ($cloudUsers.Count -eq 0) {
        Write-Host 'No cloud-only users to update.'
    } else {
        Write-Host "Updating $($cloudUsers.Count) cloud user(s)..."
        $ok = 0; $fail = 0
        foreach ($u in $cloudUsers) {
            $current = $u.UserPrincipalName
            $updated = $current -replace ([regex]::Escape("@$old") + '$'), "@$new"
            try {
                if ($WhatIf) {
                    Write-Host "  WhatIf: $current  →  $updated  [$($u.DisplayName)]"
                } else {
                    Update-MgUser -UserId $u.Id -UserPrincipalName $updated -ErrorAction Stop
                    Write-Host "  UPDATED: $current  →  $updated  [$($u.DisplayName)]"
                }
                $ok++
            } catch {
                Write-Host "  FAILED: $current — $($_.Exception.Message.Split([Environment]::NewLine)[0])"
                $fail++
            }
        }
        Write-Host ""
        Write-Host "=== Completed: updated $ok  |  failed $fail ==="
    }

    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
} else {
    try {
        Show-UpdateUpnUI
    } catch {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        [System.Windows.Forms.MessageBox]::Show("Failed to start: $($_.Exception.Message)", 'Launch Error', 'OK', 'Error') | Out-Null
    }
}
