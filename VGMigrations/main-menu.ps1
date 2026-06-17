#Requires -Version 7.0
# Migration Toolkit — Top-Level Launcher

# Load WinForms early so we can show error dialogs if lib.ps1 is missing
Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue

# ── File logging — defined before dot-sourcing so startup errors are captured ─
$_logDir = Join-Path $PSScriptRoot 'logs'
if (-not (Test-Path $_logDir)) { New-Item -ItemType Directory -Path $_logDir -Force | Out-Null }
$script:LogFile = Join-Path $_logDir "main-menu-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
function Write-Log {
    param([string]$Msg, [string]$Level = 'INFO')
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff')  [$($Level.PadRight(5))]  $Msg"
    Add-Content -Path $script:LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
}
Write-Log "=== main-menu.ps1 started  PID=$PID  PSVersion=$($PSVersionTable.PSVersion) ==="
Write-Log "Script root: $PSScriptRoot"

$_startupError = $null
try   { . "$PSScriptRoot\lib.ps1";      Write-Log 'lib.ps1 loaded OK' }
catch { $_startupError = "lib.ps1 failed to load: $($_.Exception.Message)"; Write-Log $_startupError 'ERROR' }

if ($_startupError) {
    [System.Windows.Forms.MessageBox]::Show(
        "$_startupError`n`nLog: $script:LogFile",
        'Startup Error', 'OK', 'Error') | Out-Null
    exit 1
}

try   { . "$PSScriptRoot\settings.ps1"; Write-Log 'settings.ps1 loaded OK' }
catch { Write-Log "settings.ps1 failed to load: $($_.Exception.Message)" 'WARN' }

# ── Discovery sub-menu ────────────────────────────────────────────────────────
function Show-DiscoverySubMenu {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = 'Discovery Tools'
    $dlg.ClientSize      = [System.Drawing.Size]::new(480, 280)
    $dlg.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $dlg.BackColor       = $clrBg
    $dlg.Font            = $FontBody
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $dlg.MaximizeBox     = $false
    $_ico = Join-Path $PSScriptRoot 'FlyMigration.ico'
    if (Test-Path $_ico) { $dlg.Icon = [System.Drawing.Icon]::new($_ico) }

    # ── Header ────────────────────────────────────────────────────────────────
    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Size = [System.Drawing.Size]::new(480, 56); $hdr.Dock = [System.Windows.Forms.DockStyle]::Top
    $hdr.BackColor = $clrAccent; $dlg.Controls.Add($hdr)
    $_hdrX = Add-HeaderLogo $hdr 36
    $hdrLbl = New-Object System.Windows.Forms.Label
    $hdrLbl.Text      = '  Discovery Tools'
    $hdrLbl.Font      = $FontTitle
    $hdrLbl.ForeColor = [System.Drawing.Color]::White
    $hdrLbl.Location  = [System.Drawing.Point]::new($_hdrX, 0)
    $hdrLbl.Size      = [System.Drawing.Size]::new(380, 56)
    $hdrLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdr.Controls.Add($hdrLbl)

    $btnGear = New-Object System.Windows.Forms.Button
    $btnGear.BackColor = $clrAccent; $btnGear.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnGear.FlatAppearance.BorderSize = 0
    $btnGear.FlatAppearance.MouseOverBackColor = $clrAccentHover
    $btnGear.Size = [System.Drawing.Size]::new(38, 38); $btnGear.Location = [System.Drawing.Point]::new(434, 9)
    $btnGear.Cursor = [System.Windows.Forms.Cursors]::Hand; $btnGear.Add_Click({ Show-SettingsDialog })
    if ($script:GearBitmap) { $btnGear.Image = $script:GearBitmap; $btnGear.ImageAlign = [System.Drawing.ContentAlignment]::MiddleCenter }
    else { $btnGear.Text = [char]0x2699; $btnGear.Font = New-Object System.Drawing.Font('Segoe UI Symbol', 16); $btnGear.ForeColor = [System.Drawing.Color]::White }
    $hdr.Controls.Add($btnGear)

    $bW = 400; $bH = 90; $bX = 40; $y = 82

    # ── M365 Discovery ────────────────────────────────────────────────────────
    $btn1 = New-Object System.Windows.Forms.Button
    $btn1.Text      = 'M365 Discovery'
    $btn1.Font      = $FontTile
    $btn1.Location  = [System.Drawing.Point]::new($bX, $y)
    $btn1.Size      = [System.Drawing.Size]::new($bW, $bH)
    $btn1.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btn1.FlatAppearance.BorderSize = 0
    $btn1.BackColor = $clrAccent
    $btn1.ForeColor = [System.Drawing.Color]::White
    $btn1.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $discScript = Join-Path $PSScriptRoot 'discovery-menu.ps1'
    $btn1.Add_Click({
        Write-Log "M365 Discovery clicked  path=$discScript"
        try   { [FlyConsole.NativeMethods]::AllowSetForegroundWindow(-1) | Out-Null } catch {}
        Write-Log "Launching discovery-menu: $discScript"
        try {
            Start-HiddenProcess 'pwsh.exe' "-NoProfile -ExecutionPolicy Bypass -File `"$discScript`""
            Write-Log 'Launch returned'
            $dlg.Close()
        } catch {
            Write-Log "Discovery launch FAILED: $_" 'ERROR'
            [System.Windows.Forms.MessageBox]::Show("Failed to launch:`n$_", 'Launch Error', 'OK', 'Error') | Out-Null
        }
    }.GetNewClosure())
    $dlg.Controls.Add($btn1)
    $y += $bH + 6

    $sub1 = New-Object System.Windows.Forms.Label
    $sub1.Text      = 'M365 tenant assessment — mailboxes, sites, OneDrive, groups, devices'
    $sub1.Font      = $FontSub
    $sub1.ForeColor = $clrMuted
    $sub1.Location  = [System.Drawing.Point]::new($bX + 4, $y)
    $sub1.AutoSize  = $true
    $dlg.Controls.Add($sub1)

    $footer = New-Object System.Windows.Forms.Panel
    $footer.Height = 46; $footer.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $footer.BackColor = $clrFooter
    $dlg.Controls.Add($footer)
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'; $btnClose.Size = [System.Drawing.Size]::new(90, 30)
    $btnClose.Location = [System.Drawing.Point]::new(374, 8)
    $btnClose.BackColor = $clrCloseRed
    $btnClose.ForeColor = [System.Drawing.Color]::White; $btnClose.Font = $FontBold
    $btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnClose.FlatAppearance.BorderSize = 0
    $btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnClose.Add_Click({ $dlg.Close() }.GetNewClosure())
    $footer.Controls.Add($btnClose)
    $footer.Add_SizeChanged({ $btnClose.Left = $footer.Width - 100 }.GetNewClosure())

    $dlg.ShowDialog() | Out-Null
}

# ── Misc Scripts sub-menu ────────────────────────────────────────────────────
function Show-MiscSubMenu {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = 'Misc Scripts'
    $dlg.ClientSize      = [System.Drawing.Size]::new(480, 500)
    $dlg.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $dlg.BackColor       = $clrBg
    $dlg.Font            = $FontBody
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $dlg.MaximizeBox     = $false
    $_ico = Join-Path $PSScriptRoot 'FlyMigration.ico'
    if (Test-Path $_ico) { $dlg.Icon = [System.Drawing.Icon]::new($_ico) }

    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Size = [System.Drawing.Size]::new(480, 56); $hdr.Dock = [System.Windows.Forms.DockStyle]::Top
    $hdr.BackColor = $clrAccent; $dlg.Controls.Add($hdr)
    $_hdrX = Add-HeaderLogo $hdr 36
    $hdrLbl = New-Object System.Windows.Forms.Label
    $hdrLbl.Text      = '  Misc Scripts'
    $hdrLbl.Font      = $FontTitle
    $hdrLbl.ForeColor = [System.Drawing.Color]::White
    $hdrLbl.Location  = [System.Drawing.Point]::new($_hdrX, 0)
    $hdrLbl.Size      = [System.Drawing.Size]::new(380, 56)
    $hdrLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdr.Controls.Add($hdrLbl)

    $btnGear = New-Object System.Windows.Forms.Button
    $btnGear.BackColor = $clrAccent; $btnGear.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnGear.FlatAppearance.BorderSize = 0
    $btnGear.FlatAppearance.MouseOverBackColor = $clrAccentHover
    $btnGear.Size = [System.Drawing.Size]::new(38, 38); $btnGear.Location = [System.Drawing.Point]::new(434, 9)
    $btnGear.Cursor = [System.Windows.Forms.Cursors]::Hand; $btnGear.Add_Click({ Show-SettingsDialog })
    if ($script:GearBitmap) { $btnGear.Image = $script:GearBitmap; $btnGear.ImageAlign = [System.Drawing.ContentAlignment]::MiddleCenter }
    else { $btnGear.Text = [char]0x2699; $btnGear.Font = New-Object System.Drawing.Font('Segoe UI Symbol', 16); $btnGear.ForeColor = [System.Drawing.Color]::White }
    $hdr.Controls.Add($btnGear)

    $bW = 400; $bH = 90; $bX = 40; $y = 82

    # ── Provision OneDrives ───────────────────────────────────────────────────
    $miscScript1 = Join-Path $PSScriptRoot 'provision-onedrives.ps1'
    $miscBtn1 = New-Object System.Windows.Forms.Button
    $miscBtn1.Text      = 'Provision OneDrives'
    $miscBtn1.Font      = New-Object System.Drawing.Font('Segoe UI Semibold', 14)
    $miscBtn1.Location  = [System.Drawing.Point]::new($bX, $y)
    $miscBtn1.Size      = [System.Drawing.Size]::new($bW, $bH)
    $miscBtn1.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $miscBtn1.FlatAppearance.BorderSize = 0
    $miscBtn1.BackColor = $clrAccent
    $miscBtn1.ForeColor = [System.Drawing.Color]::White
    $miscBtn1.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $miscBtn1.Add_Click({
        Write-Log "Provision OneDrives clicked  path=$miscScript1"
        if (-not (Test-Path $miscScript1)) {
            [System.Windows.Forms.MessageBox]::Show("Script not found:`n$miscScript1", 'Not Found', 'OK', 'Warning') | Out-Null; return
        }
        try   { [FlyConsole.NativeMethods]::AllowSetForegroundWindow(-1) | Out-Null } catch {}
        try {
            Start-HiddenProcess 'pwsh.exe' "-NoProfile -ExecutionPolicy Bypass -File `"$miscScript1`""
            Write-Log 'Provision OneDrives launched'
        } catch {
            Write-Log "Provision OneDrives launch FAILED: $_" 'ERROR'
            [System.Windows.Forms.MessageBox]::Show("Failed to launch:`n$_", 'Launch Error', 'OK', 'Error') | Out-Null
        }
    }.GetNewClosure())
    $dlg.Controls.Add($miscBtn1)
    $y += $bH + 6

    $miscSub1 = New-Object System.Windows.Forms.Label
    $miscSub1.Text      = 'Pre-provision OneDrive for Business sites from a mapping file'
    $miscSub1.Font      = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $miscSub1.ForeColor = $clrMuted
    $miscSub1.Location  = [System.Drawing.Point]::new($bX + 4, $y)
    $miscSub1.AutoSize  = $true
    $dlg.Controls.Add($miscSub1)
    $y += 26

    # ── Set Teams Owners ──────────────────────────────────────────────────────
    $miscScript2 = Join-Path $PSScriptRoot 'Set-TeamsOwners.ps1'
    $miscBtn2 = New-Object System.Windows.Forms.Button
    $miscBtn2.Text      = 'Set Teams Owners'
    $miscBtn2.Font      = New-Object System.Drawing.Font('Segoe UI Semibold', 14)
    $miscBtn2.Location  = [System.Drawing.Point]::new($bX, $y)
    $miscBtn2.Size      = [System.Drawing.Size]::new($bW, $bH)
    $miscBtn2.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $miscBtn2.FlatAppearance.BorderSize = 0
    $miscBtn2.BackColor = $clrAccent
    $miscBtn2.ForeColor = [System.Drawing.Color]::White
    $miscBtn2.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $miscBtn2.Add_Click({
        Write-Log "Set Teams Owners clicked  path=$miscScript2"
        if (-not (Test-Path $miscScript2)) {
            [System.Windows.Forms.MessageBox]::Show("Script not found:`n$miscScript2", 'Not Found', 'OK', 'Warning') | Out-Null; return
        }
        try   { [FlyConsole.NativeMethods]::AllowSetForegroundWindow(-1) | Out-Null } catch {}
        try {
            Start-HiddenProcess 'pwsh.exe' "-NoProfile -ExecutionPolicy Bypass -File `"$miscScript2`""
            Write-Log 'Set Teams Owners launched'
        } catch {
            Write-Log "Set Teams Owners launch FAILED: $_" 'ERROR'
            [System.Windows.Forms.MessageBox]::Show("Failed to launch:`n$_", 'Launch Error', 'OK', 'Error') | Out-Null
        }
    }.GetNewClosure())
    $dlg.Controls.Add($miscBtn2)
    $y += $bH + 6

    $miscSub2 = New-Object System.Windows.Forms.Label
    $miscSub2.Text      = 'Add a user as owner to Teams and M365 Groups from a CSV'
    $miscSub2.Font      = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $miscSub2.ForeColor = $clrMuted
    $miscSub2.Location  = [System.Drawing.Point]::new($bX + 4, $y)
    $miscSub2.AutoSize  = $true
    $dlg.Controls.Add($miscSub2)
    $y += 26

    # ── Get Domain Devices ────────────────────────────────────────────────────
    $miscScript3 = Join-Path $PSScriptRoot 'Get-DomainDevices.ps1'
    $miscBtn3 = New-Object System.Windows.Forms.Button
    $miscBtn3.Text      = 'Get Domain Devices'
    $miscBtn3.Font      = New-Object System.Drawing.Font('Segoe UI Semibold', 14)
    $miscBtn3.Location  = [System.Drawing.Point]::new($bX, $y)
    $miscBtn3.Size      = [System.Drawing.Size]::new($bW, $bH)
    $miscBtn3.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $miscBtn3.FlatAppearance.BorderSize = 0
    $miscBtn3.BackColor = $clrAccent
    $miscBtn3.ForeColor = [System.Drawing.Color]::White
    $miscBtn3.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $miscBtn3.Add_Click({
        Write-Log "Get Domain Devices clicked  path=$miscScript3"
        if (-not (Test-Path $miscScript3)) {
            [System.Windows.Forms.MessageBox]::Show("Script not found:`n$miscScript3", 'Not Found', 'OK', 'Warning') | Out-Null; return
        }
        try   { [FlyConsole.NativeMethods]::AllowSetForegroundWindow(-1) | Out-Null } catch {}
        try {
            Start-HiddenProcess 'pwsh.exe' "-NoProfile -ExecutionPolicy Bypass -File `"$miscScript3`""
            Write-Log 'Get Domain Devices launched'
        } catch {
            Write-Log "Get Domain Devices launch FAILED: $_" 'ERROR'
            [System.Windows.Forms.MessageBox]::Show("Failed to launch:`n$_", 'Launch Error', 'OK', 'Error') | Out-Null
        }
    }.GetNewClosure())
    $dlg.Controls.Add($miscBtn3)
    $y += $bH + 6

    $miscSub3 = New-Object System.Windows.Forms.Label
    $miscSub3.Text      = 'Export Entra registered devices for users in a specific domain'
    $miscSub3.Font      = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $miscSub3.ForeColor = $clrMuted
    $miscSub3.Location  = [System.Drawing.Point]::new($bX + 4, $y)
    $miscSub3.AutoSize  = $true
    $dlg.Controls.Add($miscSub3)
    $y += 26

    $dlg.ClientSize = [System.Drawing.Size]::new(480, ($y + 56))

    $footer = New-Object System.Windows.Forms.Panel
    $footer.Height = 46; $footer.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $footer.BackColor = $clrFooter
    $dlg.Controls.Add($footer)
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'; $btnClose.Size = [System.Drawing.Size]::new(90, 30)
    $btnClose.Location = [System.Drawing.Point]::new(374, 8)
    $btnClose.BackColor = $clrCloseRed
    $btnClose.ForeColor = [System.Drawing.Color]::White; $btnClose.Font = $FontBold
    $btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnClose.FlatAppearance.BorderSize = 0
    $btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnClose.Add_Click({ $dlg.Close() }.GetNewClosure())
    $footer.Controls.Add($btnClose)
    $footer.Add_SizeChanged({ $btnClose.Left = $footer.Width - 100 }.GetNewClosure())

    $dlg.ShowDialog() | Out-Null
}

# ── Domain Removal sub-menu ───────────────────────────────────────────────────
function Show-DomainRemovalSubMenu {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = 'Domain Removal'
    $dlg.ClientSize      = [System.Drawing.Size]::new(480, 534)
    $dlg.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $dlg.BackColor       = $clrBg
    $dlg.Font            = $FontBody
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $dlg.MaximizeBox     = $false
    $_ico = Join-Path $PSScriptRoot 'FlyMigration.ico'
    if (Test-Path $_ico) { $dlg.Icon = [System.Drawing.Icon]::new($_ico) }

    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Size = [System.Drawing.Size]::new(480, 56); $hdr.Dock = [System.Windows.Forms.DockStyle]::Top
    $hdr.BackColor = $clrAccent; $dlg.Controls.Add($hdr)
    $_hdrX = Add-HeaderLogo $hdr 36
    $hdrLbl = New-Object System.Windows.Forms.Label
    $hdrLbl.Text      = '  Domain Removal'
    $hdrLbl.Font      = $FontTitle
    $hdrLbl.ForeColor = [System.Drawing.Color]::White
    $hdrLbl.Location  = [System.Drawing.Point]::new($_hdrX, 0)
    $hdrLbl.Size      = [System.Drawing.Size]::new(380, 56)
    $hdrLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdr.Controls.Add($hdrLbl)

    $btnGear = New-Object System.Windows.Forms.Button
    $btnGear.BackColor = $clrAccent; $btnGear.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnGear.FlatAppearance.BorderSize = 0
    $btnGear.FlatAppearance.MouseOverBackColor = $clrAccentHover
    $btnGear.Size = [System.Drawing.Size]::new(38, 38); $btnGear.Location = [System.Drawing.Point]::new(434, 9)
    $btnGear.Cursor = [System.Windows.Forms.Cursors]::Hand; $btnGear.Add_Click({ Show-SettingsDialog })
    if ($script:GearBitmap) { $btnGear.Image = $script:GearBitmap; $btnGear.ImageAlign = [System.Drawing.ContentAlignment]::MiddleCenter }
    else { $btnGear.Text = [char]0x2699; $btnGear.Font = New-Object System.Drawing.Font('Segoe UI Symbol', 16); $btnGear.ForeColor = [System.Drawing.Color]::White }
    $hdr.Controls.Add($btnGear)

    $bW = 400; $bH = 90; $bX = 40; $y = 82

    # ── Domain Removal Workflow ───────────────────────────────────────────────
    $workflowScript = Join-Path $PSScriptRoot 'Domain-Removal-Workflow.ps1'
    $btnWorkflow = New-Object System.Windows.Forms.Button
    $btnWorkflow.Text      = 'Domain Removal Workflow'
    $btnWorkflow.Font      = New-Object System.Drawing.Font('Segoe UI Semibold', 14)
    $btnWorkflow.Location  = [System.Drawing.Point]::new($bX, $y)
    $btnWorkflow.Size      = [System.Drawing.Size]::new($bW, $bH)
    $btnWorkflow.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnWorkflow.FlatAppearance.BorderSize = 0
    $btnWorkflow.BackColor = [System.Drawing.Color]::FromArgb(195, 30, 30)
    $btnWorkflow.ForeColor = [System.Drawing.Color]::White
    $btnWorkflow.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $btnWorkflow.Add_Click({
        Write-Log "Domain Removal Workflow clicked  path=$workflowScript"
        if (-not (Test-Path $workflowScript)) {
            [System.Windows.Forms.MessageBox]::Show("Script not found:`n$workflowScript", 'Not Found', 'OK', 'Warning') | Out-Null; return
        }
        try { [FlyConsole.NativeMethods]::AllowSetForegroundWindow(-1) | Out-Null } catch {}
        try {
            Start-HiddenProcess 'pwsh.exe' "-NoProfile -ExecutionPolicy Bypass -File `"$workflowScript`""
            Write-Log 'Domain Removal Workflow launched'
        } catch {
            Write-Log "Domain Removal Workflow launch FAILED: $_" 'ERROR'
            [System.Windows.Forms.MessageBox]::Show("Failed to launch:`n$_", 'Launch Error', 'OK', 'Error') | Out-Null
        }
    }.GetNewClosure())
    $dlg.Controls.Add($btnWorkflow)
    $y += $bH + 6

    $subWorkflow = New-Object System.Windows.Forms.Label
    $subWorkflow.Text      = '3-step workflow: Update on-prem UPN → AD Sync → Remove domain'
    $subWorkflow.Font      = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $subWorkflow.ForeColor = $clrMuted
    $subWorkflow.Location  = [System.Drawing.Point]::new($bX + 4, $y)
    $subWorkflow.AutoSize  = $true
    $dlg.Controls.Add($subWorkflow)
    $y += 26

    # ── Remove Domain ─────────────────────────────────────────────────────────
    $script2 = Join-Path $PSScriptRoot 'remove-domain.ps1'
    $btn2 = New-Object System.Windows.Forms.Button
    $btn2.Text      = 'Remove Domain'
    $btn2.Font      = New-Object System.Drawing.Font('Segoe UI Semibold', 14)
    $btn2.Location  = [System.Drawing.Point]::new($bX, $y)
    $btn2.Size      = [System.Drawing.Size]::new($bW, $bH)
    $btn2.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btn2.FlatAppearance.BorderSize = 0
    $btn2.BackColor = $clrAccent
    $btn2.ForeColor = [System.Drawing.Color]::White
    $btn2.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $btn2.Add_Click({
        Write-Log "Remove Domain clicked  path=$script2"
        if (-not (Test-Path $script2)) {
            [System.Windows.Forms.MessageBox]::Show("Script not found:`n$script2", 'Not Found', 'OK', 'Warning') | Out-Null; return
        }
        try   { [FlyConsole.NativeMethods]::AllowSetForegroundWindow(-1) | Out-Null } catch {}
        try {
            Start-HiddenProcess 'pwsh.exe' "-NoProfile -ExecutionPolicy Bypass -File `"$script2`""
            Write-Log 'Remove Domain launched'
        } catch {
            Write-Log "Remove Domain launch FAILED: $_" 'ERROR'
            [System.Windows.Forms.MessageBox]::Show("Failed to launch:`n$_", 'Launch Error', 'OK', 'Error') | Out-Null
        }
    }.GetNewClosure())
    $dlg.Controls.Add($btn2)
    $y += $bH + 6

    $sub2 = New-Object System.Windows.Forms.Label
    $sub2.Text      = 'Remove a verified domain and all associated M365 objects'
    $sub2.Font      = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $sub2.ForeColor = $clrMuted
    $sub2.Location  = [System.Drawing.Point]::new($bX + 4, $y)
    $sub2.AutoSize  = $true
    $dlg.Controls.Add($sub2)
    $y += 26

    # ── Update On-Premise UPNs ────────────────────────────────────────────────
    $domScript3 = Join-Path $PSScriptRoot 'Update-OnPremUPN.ps1'
    $domBtn3 = New-Object System.Windows.Forms.Button
    $domBtn3.Text      = 'Update On-Premise UPNs'
    $domBtn3.Font      = New-Object System.Drawing.Font('Segoe UI Semibold', 14)
    $domBtn3.Location  = [System.Drawing.Point]::new($bX, $y)
    $domBtn3.Size      = [System.Drawing.Size]::new($bW, $bH)
    $domBtn3.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $domBtn3.FlatAppearance.BorderSize = 0
    $domBtn3.BackColor = $clrAccent
    $domBtn3.ForeColor = [System.Drawing.Color]::White
    $domBtn3.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $domBtn3.Add_Click({
        Write-Log "Update On-Premise UPNs clicked  path=$domScript3"
        if (-not (Test-Path $domScript3)) {
            [System.Windows.Forms.MessageBox]::Show("Script not found:`n$domScript3", 'Not Found', 'OK', 'Warning') | Out-Null; return
        }
        try { [FlyConsole.NativeMethods]::AllowSetForegroundWindow(-1) | Out-Null } catch {}
        try {
            Start-HiddenProcess 'pwsh.exe' "-NoProfile -ExecutionPolicy Bypass -File `"$domScript3`""
            Write-Log 'Update On-Premise UPNs launched'
        } catch {
            Write-Log "Update On-Premise UPNs launch FAILED: $_" 'ERROR'
            [System.Windows.Forms.MessageBox]::Show("Failed to launch:`n$_", 'Launch Error', 'OK', 'Error') | Out-Null
        }
    }.GetNewClosure())
    $dlg.Controls.Add($domBtn3)
    $y += $bH + 6

    $domSub3 = New-Object System.Windows.Forms.Label
    $domSub3.Text      = 'Update UPN, email, and aliases in on-premise Active Directory'
    $domSub3.Font      = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $domSub3.ForeColor = $clrMuted
    $domSub3.Location  = [System.Drawing.Point]::new($bX + 4, $y)
    $domSub3.AutoSize  = $true
    $dlg.Controls.Add($domSub3)
    $y += 26

    # ── Update Cloud UPNs ─────────────────────────────────────────────────────
    $domScript3b = Join-Path $PSScriptRoot 'Update-UPN.ps1'
    $domBtn3b = New-Object System.Windows.Forms.Button
    $domBtn3b.Text      = 'Update Cloud UPNs'
    $domBtn3b.Font      = New-Object System.Drawing.Font('Segoe UI Semibold', 14)
    $domBtn3b.Location  = [System.Drawing.Point]::new($bX, $y)
    $domBtn3b.Size      = [System.Drawing.Size]::new($bW, $bH)
    $domBtn3b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $domBtn3b.FlatAppearance.BorderSize = 0
    $domBtn3b.BackColor = $clrAccent
    $domBtn3b.ForeColor = [System.Drawing.Color]::White
    $domBtn3b.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $domBtn3b.Add_Click({
        Write-Log "Update Cloud UPNs clicked  path=$domScript3b"
        if (-not (Test-Path $domScript3b)) {
            [System.Windows.Forms.MessageBox]::Show("Script not found:`n$domScript3b", 'Not Found', 'OK', 'Warning') | Out-Null; return
        }
        try { [FlyConsole.NativeMethods]::AllowSetForegroundWindow(-1) | Out-Null } catch {}
        try {
            Start-HiddenProcess 'pwsh.exe' "-NoProfile -ExecutionPolicy Bypass -File `"$domScript3b`""
            Write-Log 'Update Cloud UPNs launched'
        } catch {
            Write-Log "Update Cloud UPNs launch FAILED: $_" 'ERROR'
            [System.Windows.Forms.MessageBox]::Show("Failed to launch:`n$_", 'Launch Error', 'OK', 'Error') | Out-Null
        }
    }.GetNewClosure())
    $dlg.Controls.Add($domBtn3b)
    $y += $bH + 6

    $domSub3b = New-Object System.Windows.Forms.Label
    $domSub3b.Text      = 'Change UPN domain suffix for cloud users via Microsoft Graph'
    $domSub3b.Font      = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $domSub3b.ForeColor = $clrMuted
    $domSub3b.Location  = [System.Drawing.Point]::new($bX + 4, $y)
    $domSub3b.AutoSize  = $true
    $dlg.Controls.Add($domSub3b)
    $y += 26

    # ── Run AD Sync ───────────────────────────────────────────────────────────
    $domBtn3c = New-Object System.Windows.Forms.Button
    $domBtn3c.Text      = 'Run AD Sync'
    $domBtn3c.Font      = New-Object System.Drawing.Font('Segoe UI Semibold', 14)
    $domBtn3c.Location  = [System.Drawing.Point]::new($bX, $y)
    $domBtn3c.Size      = [System.Drawing.Size]::new($bW, $bH)
    $domBtn3c.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $domBtn3c.FlatAppearance.BorderSize = 0
    $domBtn3c.BackColor = [System.Drawing.Color]::FromArgb(0, 130, 70)
    $domBtn3c.ForeColor = [System.Drawing.Color]::White
    $domBtn3c.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $domBtn3c.Add_Click({
        Write-Log "Run AD Sync clicked"
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Trigger Azure AD Connect sync on VOL-ane-aad1?`n`nThis will start a delta sync cycle.",
            'Confirm AD Sync',
            'YesNo',
            'Question'
        )
        if ($result -eq 'Yes') {
            try {
                Write-Log "Starting AD Sync on VOL-ane-aad1..."
                $syncScript = {
                    Import-Module ADSync -ErrorAction Stop
                    Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Stop
                }
                $session = New-PSSession -ComputerName 'VOL-ane-aad1' -ErrorAction Stop
                $syncResult = Invoke-Command -Session $session -ScriptBlock $syncScript -ErrorAction Stop
                Remove-PSSession $session
                Write-Log "AD Sync completed: $($syncResult.Result)" 'OK'
                [System.Windows.Forms.MessageBox]::Show(
                    "AD Sync initiated successfully.`n`nResult: $($syncResult.Result)",
                    'AD Sync',
                    'OK',
                    'Information'
                ) | Out-Null
            } catch {
                Write-Log "AD Sync failed: $_" 'ERROR'
                [System.Windows.Forms.MessageBox]::Show(
                    "Failed to run AD Sync:`n`n$_",
                    'AD Sync Error',
                    'OK',
                    'Error'
                ) | Out-Null
            }
        }
    }.GetNewClosure())
    $dlg.Controls.Add($domBtn3c)
    $y += $bH + 6

    $domSub3c = New-Object System.Windows.Forms.Label
    $domSub3c.Text      = 'Trigger Azure AD Connect delta sync on VOL-ane-aad1 server'
    $domSub3c.Font      = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $domSub3c.ForeColor = $clrMuted
    $domSub3c.Location  = [System.Drawing.Point]::new($bX + 4, $y)
    $domSub3c.AutoSize  = $true
    $dlg.Controls.Add($domSub3c)
    $y += 26

    # ── Hide from Address Book ────────────────────────────────────────────────
    $domScript4 = Join-Path $PSScriptRoot 'Hide-AddressBook.ps1'
    $domBtn4 = New-Object System.Windows.Forms.Button
    $domBtn4.Text      = 'Hide from Address Book'
    $domBtn4.Font      = New-Object System.Drawing.Font('Segoe UI Semibold', 14)
    $domBtn4.Location  = [System.Drawing.Point]::new($bX, $y)
    $domBtn4.Size      = [System.Drawing.Size]::new($bW, $bH)
    $domBtn4.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $domBtn4.FlatAppearance.BorderSize = 0
    $domBtn4.BackColor = $clrAccent
    $domBtn4.ForeColor = [System.Drawing.Color]::White
    $domBtn4.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $domBtn4.Add_Click({
        Write-Log "Hide from Address Book clicked  path=$domScript4"
        if (-not (Test-Path $domScript4)) {
            [System.Windows.Forms.MessageBox]::Show("Script not found:`n$domScript4", 'Not Found', 'OK', 'Warning') | Out-Null; return
        }
        try { [FlyConsole.NativeMethods]::AllowSetForegroundWindow(-1) | Out-Null } catch {}
        try {
            Start-HiddenProcess 'pwsh.exe' "-NoProfile -ExecutionPolicy Bypass -File `"$domScript4`""
            Write-Log 'Hide from Address Book launched'
        } catch {
            Write-Log "Hide from Address Book launch FAILED: $_" 'ERROR'
            [System.Windows.Forms.MessageBox]::Show("Failed to launch:`n$_", 'Launch Error', 'OK', 'Error') | Out-Null
        }
    }.GetNewClosure())
    $dlg.Controls.Add($domBtn4)
    $y += $bH + 6

    $domSub4 = New-Object System.Windows.Forms.Label
    $domSub4.Text      = 'Bulk hide Exchange Online recipients from the Global Address List'
    $domSub4.Font      = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $domSub4.ForeColor = $clrMuted
    $domSub4.Location  = [System.Drawing.Point]::new($bX + 4, $y)
    $domSub4.AutoSize  = $true
    $dlg.Controls.Add($domSub4)
    $y += 26

    $dlg.ClientSize = [System.Drawing.Size]::new(480, ($y + 56))

    $footer = New-Object System.Windows.Forms.Panel
    $footer.Height = 46; $footer.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $footer.BackColor = $clrFooter
    $dlg.Controls.Add($footer)
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'; $btnClose.Size = [System.Drawing.Size]::new(90, 30)
    $btnClose.Location = [System.Drawing.Point]::new(374, 8)
    $btnClose.BackColor = $clrCloseRed
    $btnClose.ForeColor = [System.Drawing.Color]::White; $btnClose.Font = $FontBold
    $btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnClose.FlatAppearance.BorderSize = 0
    $btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnClose.Add_Click({ $dlg.Close() }.GetNewClosure())
    $footer.Controls.Add($btnClose)
    $footer.Add_SizeChanged({ $btnClose.Left = $footer.Width - 100 }.GetNewClosure())

    $dlg.ShowDialog() | Out-Null
}

# ── Main launcher ─────────────────────────────────────────────────────────────
function Show-Launcher {
    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'Migration Toolkit - Dashboard'
    $form.ClientSize      = [System.Drawing.Size]::new(1000, 700)
    $form.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.BackColor       = $clrBg
    $form.Font            = $FontBody
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $form.MaximizeBox     = $true
    $form.WindowState     = [System.Windows.Forms.FormWindowState]::Maximized
    $_ico = Join-Path $PSScriptRoot 'FlyMigration.ico'
    if (Test-Path $_ico) { $form.Icon = [System.Drawing.Icon]::new($_ico) }

    # ── Header ────────────────────────────────────────────────────────────────
    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Size = [System.Drawing.Size]::new(1000, 80)
    $hdr.Dock = [System.Windows.Forms.DockStyle]::Top
    $hdr.BackColor = $clrAccent
    $form.Controls.Add($hdr)
    $_hdrX = Add-HeaderLogo $hdr 32
    $hdrLbl = New-Object System.Windows.Forms.Label
    $hdrLbl.Text      = '  Migration Toolkit - Dashboard'
    $hdrLbl.Font      = New-Object System.Drawing.Font('Segoe UI Light', 22)
    $hdrLbl.ForeColor = [System.Drawing.Color]::White
    $hdrLbl.Location  = [System.Drawing.Point]::new($_hdrX, 0)
    $hdrLbl.Size      = [System.Drawing.Size]::new(800, 80)
    $hdrLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdr.Controls.Add($hdrLbl)

    $btnGear = New-Object System.Windows.Forms.Button
    $btnGear.BackColor = $clrAccent; $btnGear.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnGear.FlatAppearance.BorderSize = 0
    $btnGear.FlatAppearance.MouseOverBackColor = $clrAccentHover
    $btnGear.Size = [System.Drawing.Size]::new(42, 42); $btnGear.Location = [System.Drawing.Point]::new(940, 19)
    $btnGear.Cursor = [System.Windows.Forms.Cursors]::Hand; $btnGear.Add_Click({ Show-SettingsDialog })
    if ($script:GearBitmap) { $btnGear.Image = $script:GearBitmap; $btnGear.ImageAlign = [System.Drawing.ContentAlignment]::MiddleCenter }
    else { $btnGear.Text = [char]0x2699; $btnGear.Font = New-Object System.Drawing.Font('Segoe UI Symbol', 20); $btnGear.ForeColor = [System.Drawing.Color]::White }
    $hdr.Controls.Add($btnGear)

    # Stat card helper function
    function MkStatCard { param([int]$X,[int]$Y,[int]$W,[int]$H,[string]$Label,[string]$Value,[string]$Status)
        $card = New-Object System.Windows.Forms.Panel
        $card.Location = [System.Drawing.Point]::new($X,$Y)
        $card.Size = [System.Drawing.Size]::new($W,$H)
        $card.BackColor = [System.Drawing.Color]::White
        $card.BorderStyle = [System.Windows.Forms.BorderStyle]::None
        $radius = 8
        $regionPath = New-Object System.Drawing.Drawing2D.GraphicsPath
        $regionRect = [System.Drawing.Rectangle]::new(0, 0, $W, $H)
        $regionPath.AddArc($regionRect.X, $regionRect.Y, $radius * 2, $radius * 2, 180, 90)
        $regionPath.AddArc($regionRect.Right - $radius * 2, $regionRect.Y, $radius * 2, $radius * 2, 270, 90)
        $regionPath.AddArc($regionRect.Right - $radius * 2, $regionRect.Bottom - $radius * 2, $radius * 2, $radius * 2, 0, 90)
        $regionPath.AddArc($regionRect.X, $regionRect.Bottom - $radius * 2, $radius * 2, $radius * 2, 90, 90)
        $regionPath.CloseFigure()
        $card.Region = New-Object System.Drawing.Region($regionPath)
        $card.Add_Paint({
            param($s, $e)
            $g = $e.Graphics
            $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
            $r = 8
            $rect = [System.Drawing.Rectangle]::new(0, 0, $s.Width - 1, $s.Height - 1)
            $path = New-Object System.Drawing.Drawing2D.GraphicsPath
            $path.AddArc($rect.X, $rect.Y, $r * 2, $r * 2, 180, 90)
            $path.AddArc($rect.Right - $r * 2, $rect.Y, $r * 2, $r * 2, 270, 90)
            $path.AddArc($rect.Right - $r * 2, $rect.Bottom - $r * 2, $r * 2, $r * 2, 0, 90)
            $path.AddArc($rect.X, $rect.Bottom - $r * 2, $r * 2, $r * 2, 90, 90)
            $path.CloseFigure()
            $pen = New-Object System.Drawing.Pen($clrBorder, 1)
            $g.DrawPath($pen, $path)
            $pen.Dispose()
            $path.Dispose()
        }.GetNewClosure())
        $lblLabel = New-Object System.Windows.Forms.Label
        $lblLabel.Text = $Label.ToUpper()
        $lblLabel.Location = [System.Drawing.Point]::new(16, 16)
        $lblLabel.Size = [System.Drawing.Size]::new($W - 32, 18)
        $lblLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 8.5)
        $lblLabel.ForeColor = $clrText
        $card.Controls.Add($lblLabel)
        $lblValue = New-Object System.Windows.Forms.Label
        $lblValue.Text = $Value
        $lblValue.Location = [System.Drawing.Point]::new(16, 36)
        $lblValue.AutoSize = $true
        $lblValue.Font = New-Object System.Drawing.Font('Segoe UI Light', 28)
        $lblValue.ForeColor = $clrText
        $card.Controls.Add($lblValue)
        if ($Status) {
            $lblStatus = New-Object System.Windows.Forms.Label
            $lblStatus.Text = $Status
            $lblStatus.Location = [System.Drawing.Point]::new(16, 80)
            $lblStatus.AutoSize = $true
            $lblStatus.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 8)
            $lblStatus.ForeColor = $clrMuted
            $card.Controls.Add($lblStatus)
        }
        $form.Controls.Add($card)
        return $card
    }

    $margin = 32; $gap = 20; $y = 110
    $statW = 220; $statH = 110
    $stat1 = MkStatCard $margin $y $statW $statH 'Active Projects' '3' 'Running'
    $stat2 = MkStatCard ($margin + $statW + $gap) $y $statW $statH 'Users Migrated' '1,247' 'This month'
    $stat3 = MkStatCard ($margin + ($statW + $gap) * 2) $y $statW $statH 'Success Rate' '94%' 'Last 30 days'
    $stat4 = MkStatCard ($margin + ($statW + $gap) * 3) $y $statW $statH 'Total Data' '2.4 TB' 'Transferred'
    $y += $statH + $gap + 20

    # Progress by Workload section
    $lblProgress = New-Object System.Windows.Forms.Label
    $lblProgress.Text = 'Progress by Workload'
    $lblProgress.Location = [System.Drawing.Point]::new($margin, $y)
    $lblProgress.AutoSize = $true
    $lblProgress.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 12)
    $lblProgress.ForeColor = $clrText
    $form.Controls.Add($lblProgress)
    $y += 35

    # Helper function for workload progress cards
    function MkWorkloadCard { param([int]$X,[int]$Y,[int]$W,[string]$Icon,[string]$Name,[int]$Current,[int]$Total)
        $card = New-Object System.Windows.Forms.Panel
        $card.Location = [System.Drawing.Point]::new($X,$Y)
        $card.Size = [System.Drawing.Size]::new($W, 90)
        $card.BackColor = [System.Drawing.Color]::White
        $card.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

        # Icon + Name
        $lblIcon = New-Object System.Windows.Forms.Label
        $lblIcon.Text = $Icon
        $lblIcon.Font = New-Object System.Drawing.Font('Segoe UI', 18)
        $lblIcon.Location = [System.Drawing.Point]::new(16, 16)
        $lblIcon.AutoSize = $true
        $card.Controls.Add($lblIcon)

        $lblName = New-Object System.Windows.Forms.Label
        $lblName.Text = $Name
        $lblName.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 11)
        $lblName.ForeColor = $clrText
        $lblName.Location = [System.Drawing.Point]::new(52, 20)
        $lblName.AutoSize = $true
        $card.Controls.Add($lblName)

        # Progress bar background
        $progBg = New-Object System.Windows.Forms.Panel
        $progBg.Location = [System.Drawing.Point]::new(16, 52)
        $progBg.Size = [System.Drawing.Size]::new($W - 32, 8)
        $progBg.BackColor = [System.Drawing.Color]::FromArgb(232, 236, 243)
        $card.Controls.Add($progBg)

        # Progress bar fill
        $percent = [Math]::Round(($Current / $Total) * 100)
        $fillWidth = [int](($W - 32) * ($Current / $Total))
        $progFill = New-Object System.Windows.Forms.Panel
        $progFill.Location = [System.Drawing.Point]::new(0, 0)
        $progFill.Size = [System.Drawing.Size]::new($fillWidth, 8)
        $progFill.BackColor = $clrAccent
        $progBg.Controls.Add($progFill)

        # Progress text
        $lblProg = New-Object System.Windows.Forms.Label
        $lblProg.Text = "$Current / $Total $(if($Name -eq 'SharePoint'){'sites'}elseif($Name -eq 'Teams'){'teams'}elseif($Name -eq 'Exchange Online'){'mailboxes'}else{'accounts'})"
        $lblProg.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
        $lblProg.ForeColor = $clrMuted
        $lblProg.Location = [System.Drawing.Point]::new(16, 66)
        $lblProg.AutoSize = $true
        $card.Controls.Add($lblProg)

        $form.Controls.Add($card)
    }

    # Workload cards - 2 per row
    $wCardW = 465
    MkWorkloadCard $margin $y $wCardW '📧' 'Exchange Online' 750 1000
    MkWorkloadCard ($margin + $wCardW + $gap) $y $wCardW '📁' 'OneDrive' 820 1000
    $y += 90 + $gap
    MkWorkloadCard $margin $y $wCardW '📑' 'SharePoint' 45 100
    MkWorkloadCard ($margin + $wCardW + $gap) $y $wCardW '👥' 'Teams' 90 100
    $y += 90 + $gap + 10

    $lblActions = New-Object System.Windows.Forms.Label
    $lblActions.Text = 'Quick Actions'
    $lblActions.Location = [System.Drawing.Point]::new($margin, $y)
    $lblActions.AutoSize = $true
    $lblActions.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 12)
    $lblActions.ForeColor = $clrText
    $form.Controls.Add($lblActions)
    $y += 35

    $bW = 305; $bH = 110; $bX = $margin

    # Action tiles in 3-column grid
    # Row 1
    $btnDisc = New-Object System.Windows.Forms.Button
    $btnDisc.Text      = 'Discovery'
    $btnDisc.Font      = New-Object System.Drawing.Font('Segoe UI Semibold', 13)
    $btnDisc.Location  = [System.Drawing.Point]::new($bX, $y)
    $btnDisc.Size      = [System.Drawing.Size]::new($bW, $bH)
    $btnDisc.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnDisc.FlatAppearance.BorderSize = 0
    $btnDisc.BackColor = $clrAccent
    $btnDisc.ForeColor = [System.Drawing.Color]::White
    $btnDisc.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $btnDisc.Add_Click({
        Write-Log 'Discovery tile clicked — opening sub-menu'
        Show-DiscoverySubMenu
        Write-Log 'Discovery sub-menu closed'
    }.GetNewClosure())
    $form.Controls.Add($btnDisc)

    $menuScript = Join-Path $PSScriptRoot 'menu.ps1'
    $btnAve = New-Object System.Windows.Forms.Button
    $btnAve.Text      = 'AvePoint Fly'
    $btnAve.Font      = New-Object System.Drawing.Font('Segoe UI Semibold', 13)
    $btnAve.Location  = [System.Drawing.Point]::new($bX + $bW + $gap, $y)
    $btnAve.Size      = [System.Drawing.Size]::new($bW, $bH)
    $btnAve.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnAve.FlatAppearance.BorderSize = 0
    $btnAve.BackColor = $clrAccent
    $btnAve.ForeColor = [System.Drawing.Color]::White
    $btnAve.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $btnAve.Add_Click({
        Write-Log "AvePoint Fly clicked  path=$menuScript"
        try   { [FlyConsole.NativeMethods]::AllowSetForegroundWindow(-1) | Out-Null } catch {}
        Write-Log "Launching menu: $menuScript"
        try {
            Start-HiddenProcess 'pwsh.exe' "-NoProfile -ExecutionPolicy Bypass -File `"$menuScript`""
            Write-Log 'Launch returned'
        } catch {
            Write-Log "AvePoint Fly launch FAILED: $_" 'ERROR'
            [System.Windows.Forms.MessageBox]::Show("Failed to launch:`n$_", 'Launch Error', 'OK', 'Error') | Out-Null
        }
    }.GetNewClosure())
    $form.Controls.Add($btnAve)

    $btnMisc = New-Object System.Windows.Forms.Button
    $btnMisc.Text      = 'Misc Scripts'
    $btnMisc.Font      = New-Object System.Drawing.Font('Segoe UI Semibold', 13)
    $btnMisc.Location  = [System.Drawing.Point]::new($bX + ($bW + $gap) * 2, $y)
    $btnMisc.Size      = [System.Drawing.Size]::new($bW, $bH)
    $btnMisc.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnMisc.FlatAppearance.BorderSize = 0
    $btnMisc.BackColor = $clrAccent
    $btnMisc.ForeColor = [System.Drawing.Color]::White
    $btnMisc.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $btnMisc.Add_Click({
        Write-Log 'Misc Scripts tile clicked — opening sub-menu'
        Show-MiscSubMenu
        Write-Log 'Misc Scripts sub-menu closed'
    }.GetNewClosure())
    $form.Controls.Add($btnMisc)
    $y += $bH + $gap

    # Row 2
    $btnDom = New-Object System.Windows.Forms.Button
    $btnDom.Text      = 'Domain Removal'
    $btnDom.Font      = New-Object System.Drawing.Font('Segoe UI Semibold', 13)
    $btnDom.Location  = [System.Drawing.Point]::new($bX, $y)
    $btnDom.Size      = [System.Drawing.Size]::new($bW, $bH)
    $btnDom.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnDom.FlatAppearance.BorderSize = 0
    $btnDom.BackColor = $clrAccent
    $btnDom.ForeColor = [System.Drawing.Color]::White
    $btnDom.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $btnDom.Add_Click({
        Write-Log 'Domain Removal tile clicked — opening sub-menu'
        Show-DomainRemovalSubMenu
        Write-Log 'Domain Removal sub-menu closed'
    }.GetNewClosure())
    $form.Controls.Add($btnDom)

    $footer = New-Object System.Windows.Forms.Panel
    $footer.Height = 56; $footer.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $footer.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)
    $form.Controls.Add($footer)

    $lblVersion = New-Object System.Windows.Forms.Label
    $lblVersion.Text = "Version $($script:ToolVersion)"
    $lblVersion.Location = [System.Drawing.Point]::new(32, 18)
    $lblVersion.AutoSize = $true
    $lblVersion.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $lblVersion.ForeColor = $clrMuted
    $footer.Controls.Add($lblVersion)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'; $btnClose.Size = [System.Drawing.Size]::new(100, 36)
    $btnClose.Location = [System.Drawing.Point]::new(880, 10)
    $btnClose.BackColor = $clrCloseRed
    $btnClose.ForeColor = [System.Drawing.Color]::White; $btnClose.Font = $FontBold
    $btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnClose.FlatAppearance.BorderSize = 0
    $btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnClose.Add_Click({ $form.Close() }.GetNewClosure())
    $footer.Controls.Add($btnClose)
    $footer.Add_SizeChanged({ $btnClose.Left = $footer.Width - 106 }.GetNewClosure())

    $form.Add_FormClosed({ Write-Log '=== Launcher closed ===' }.GetNewClosure())

    Write-Log 'Launcher form ready — entering Application::Run'
    try {
        [System.Windows.Forms.Application]::Run($form)
    } catch {
        Write-Log "Application::Run crashed: $($_.Exception.Message)" 'ERROR'
        Write-Log "  $($_.ScriptStackTrace)" 'ERROR'
        [System.Windows.Forms.MessageBox]::Show(
            "Fatal error launching the main window:`n$($_.Exception.Message)`n`nLog: $script:LogFile",
            'Fatal Error', 'OK', 'Error') | Out-Null
    }
}

# ── Check for Updates ─────────────────────────────────────────────────────────
$CheckUpdatesScript = Join-Path $PSScriptRoot 'Check-Updates.ps1'
if (Test-Path $CheckUpdatesScript) {
    try {
        Write-Log 'Checking for updates...'
        # Run update check silently in background (won't block startup)
        $null = Start-Job -ScriptBlock {
            param($ScriptPath)
            & $ScriptPath -Silent
        } -ArgumentList $CheckUpdatesScript

        # Don't wait for update check - let it run in background
        Write-Log 'Update check started in background'
    } catch {
        Write-Log "Update check failed to start: $($_.Exception.Message)" 'WARN'
    }
}

Write-Log 'Calling Show-Launcher'

# Must be set before any controls are created (i.e. before Show-Launcher instantiates the Form)
[System.Windows.Forms.Application]::SetUnhandledExceptionMode(
    [System.Windows.Forms.UnhandledExceptionMode]::CatchException)
[System.Windows.Forms.Application]::add_ThreadException({
    param($s, $e)
    Write-Log "UNHANDLED UI EXCEPTION: $($e.Exception.Message)" 'ERROR'
    Write-Log "  $($e.Exception.StackTrace -replace [Environment]::NewLine,' | ')" 'ERROR'
    [System.Windows.Forms.MessageBox]::Show(
        "An unexpected error occurred:`n$($e.Exception.Message)`n`nSee log for details:`n$script:LogFile",
        'Error', 'OK', 'Error') | Out-Null
})

try {
    Show-Launcher
} catch {
    Write-Log "Show-Launcher crashed: $($_.Exception.Message)" 'ERROR'
    [System.Windows.Forms.MessageBox]::Show(
        "Failed to start:`n$($_.Exception.Message)`n`nLog: $script:LogFile",
        'Startup Failure', 'OK', 'Error') | Out-Null
}
Write-Log '=== main-menu.ps1 exiting ==='
