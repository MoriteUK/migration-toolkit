#Requires -Version 7.0
<#
.SYNOPSIS
    Domain-Removal-Workflow.ps1 — Complete domain removal workflow in three steps.

.DESCRIPTION
    Step 1: Update on-premise UPN to @ourvolaris.onmicrosoft.com
    Step 2: Run Azure AD Connect sync on VOL-ane-aad1
    Step 3: Remove domain from Microsoft 365

.NOTES
    Requires: ActiveDirectory, Microsoft.Graph modules
    Run from a machine with access to on-premise AD and VOL-ane-aad1
#>

$ErrorActionPreference = 'Stop'
$script:RootDir = $PSScriptRoot

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ── Logging ───────────────────────────────────────────────────────────────────
$_logDir = Join-Path $script:RootDir 'logs'
if (-not (Test-Path $_logDir)) { New-Item -ItemType Directory -Path $_logDir -Force | Out-Null }
$script:LogFile = Join-Path $_logDir "domain-removal-workflow-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

function Write-Log {
    param([string]$Msg, [string]$Level = 'INFO')
    $ts = Get-Date -Format 'HH:mm:ss'
    "$ts [$Level] $Msg" | Add-Content -Path $script:LogFile -Encoding UTF8 -ErrorAction SilentlyContinue

    if ($script:rtbLog -and -not $script:rtbLog.IsDisposed) {
        $script:rtbLog.SelectionStart = $script:rtbLog.TextLength
        $script:rtbLog.SelectionLength = 0
        $script:rtbLog.SelectionColor = [System.Drawing.Color]::FromArgb(80, 95, 120)
        $script:rtbLog.AppendText("$ts ")

        $lc = switch ($Level) {
            'OK'    { [System.Drawing.Color]::FromArgb(65, 195, 110)  }
            'WARN'  { [System.Drawing.Color]::FromArgb(220, 165, 45)  }
            'ERROR' { [System.Drawing.Color]::FromArgb(225, 80, 80)   }
            default { [System.Drawing.Color]::FromArgb(120, 155, 220) }
        }
        $script:rtbLog.SelectionColor = $lc
        $script:rtbLog.AppendText("[$Level] ")
        $script:rtbLog.SelectionColor = [System.Drawing.Color]::FromArgb(205, 212, 230)
        $script:rtbLog.AppendText("$Msg`n")
        $script:rtbLog.ScrollToCaret()
    }
}

Write-Log "=== Domain-Removal-Workflow.ps1 started ==="

# ── UI Theme ──────────────────────────────────────────────────────────────────
$libPath = Join-Path $script:RootDir 'lib.ps1'
if (Test-Path $libPath) {
    try { . $libPath }
    catch { Write-Log "lib.ps1 load error: $_" 'WARN' }
}

if (-not (Get-Variable -Name 'clrBg' -Scope Script -ErrorAction SilentlyContinue)) {
    $clrBg     = [System.Drawing.Color]::FromArgb(240, 242, 247)
    $clrPanel  = [System.Drawing.Color]::White
    $clrAccent = [System.Drawing.Color]::FromArgb(0, 100, 180)
    $clrText   = [System.Drawing.Color]::FromArgb(28, 28, 32)
    $clrMuted  = [System.Drawing.Color]::FromArgb(100, 108, 120)
    $clrBorder = [System.Drawing.Color]::FromArgb(210, 215, 228)
    $clrLogBg  = [System.Drawing.Color]::FromArgb(26, 27, 38)
    $clrGreen  = [System.Drawing.Color]::FromArgb(0, 130, 70)
    $clrRed    = [System.Drawing.Color]::FromArgb(195, 30, 30)
    $FontBody  = New-Object System.Drawing.Font('Segoe UI', 9)
    $FontBold  = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
    $FontCap   = New-Object System.Drawing.Font('Segoe UI Semibold', 7.5)
    $FontMono  = New-Object System.Drawing.Font('Consolas', 8.5)
    $FontTitle = New-Object System.Drawing.Font('Segoe UI Semibold', 14)
    $FontLarge = New-Object System.Drawing.Font('Segoe UI Semibold', 12)
}

# ── Global State ──────────────────────────────────────────────────────────────
$script:TargetDomain = '@ourvolaris.onmicrosoft.com'
$script:AllUsers = [System.Collections.Generic.List[pscustomobject]]::new()
$script:CurrentStep = 1

# ── Main UI ───────────────────────────────────────────────────────────────────
function Show-DomainRemovalWorkflowUI {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Domain Removal Workflow'
    $form.ClientSize = [System.Drawing.Size]::new(1100, 900)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.BackColor = $clrBg
    $form.Font = $FontBody
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.MaximizeBox = $false
    $_ico = Join-Path $script:RootDir 'FlyMigration.ico'
    if (Test-Path $_ico) { $form.Icon = [System.Drawing.Icon]::new($_ico) }

    # ── Header ────────────────────────────────────────────────────────────────
    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Size = [System.Drawing.Size]::new(1100, 56)
    $hdr.Dock = [System.Windows.Forms.DockStyle]::Top
    $hdr.BackColor = $clrAccent
    $form.Controls.Add($hdr)

    $hdrLbl = New-Object System.Windows.Forms.Label
    $hdrLbl.Text = '  Domain Removal Workflow'
    $hdrLbl.Font = $FontTitle
    $hdrLbl.ForeColor = [System.Drawing.Color]::White
    $hdrLbl.Location = [System.Drawing.Point]::new(12, 0)
    $hdrLbl.Size = [System.Drawing.Size]::new(900, 56)
    $hdrLbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdr.Controls.Add($hdrLbl)

    # ── Footer ────────────────────────────────────────────────────────────────
    $footer = New-Object System.Windows.Forms.Panel
    $footer.Height = 46
    $footer.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $footer.BackColor = [System.Drawing.Color]::FromArgb(20, 24, 38)
    $form.Controls.Add($footer)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'
    $btnClose.Size = [System.Drawing.Size]::new(90, 30)
    $btnClose.Location = [System.Drawing.Point]::new(994, 8)
    $btnClose.BackColor = [System.Drawing.Color]::FromArgb(200, 55, 55)
    $btnClose.ForeColor = [System.Drawing.Color]::White
    $btnClose.Font = $FontBold
    $btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnClose.FlatAppearance.BorderSize = 0
    $btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnClose.Add_Click({ $form.Close() })
    $footer.Controls.Add($btnClose)

    # ── Workflow Steps Panel ──────────────────────────────────────────────────
    $stepsPanel = New-Object System.Windows.Forms.Panel
    $stepsPanel.Location = [System.Drawing.Point]::new(12, 66)
    $stepsPanel.Size = [System.Drawing.Size]::new(1076, 80)
    $stepsPanel.BackColor = $clrPanel
    $form.Controls.Add($stepsPanel)

    # Step indicators
    $stepWidth = 350
    $stepX = 12

    # Step 1
    $step1Panel = New-Object System.Windows.Forms.Panel
    $step1Panel.Location = [System.Drawing.Point]::new($stepX, 10)
    $step1Panel.Size = [System.Drawing.Size]::new($stepWidth, 60)
    $step1Panel.BackColor = $clrAccent
    $stepsPanel.Controls.Add($step1Panel)

    $step1Num = New-Object System.Windows.Forms.Label
    $step1Num.Text = '1'
    $step1Num.Font = New-Object System.Drawing.Font('Segoe UI', 20, [System.Drawing.FontStyle]::Bold)
    $step1Num.ForeColor = [System.Drawing.Color]::White
    $step1Num.Location = [System.Drawing.Point]::new(10, 10)
    $step1Num.Size = [System.Drawing.Size]::new(40, 40)
    $step1Num.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $step1Panel.Controls.Add($step1Num)

    $step1Lbl = New-Object System.Windows.Forms.Label
    $step1Lbl.Text = "Update On-Prem UPN`n→ @ourvolaris.onmicrosoft.com"
    $step1Lbl.Font = $FontBody
    $step1Lbl.ForeColor = [System.Drawing.Color]::White
    $step1Lbl.Location = [System.Drawing.Point]::new(55, 10)
    $step1Lbl.Size = [System.Drawing.Size]::new(285, 40)
    $step1Panel.Controls.Add($step1Lbl)

    $stepX += $stepWidth + 12

    # Step 2
    $step2Panel = New-Object System.Windows.Forms.Panel
    $step2Panel.Location = [System.Drawing.Point]::new($stepX, 10)
    $step2Panel.Size = [System.Drawing.Size]::new($stepWidth, 60)
    $step2Panel.BackColor = $clrMuted
    $stepsPanel.Controls.Add($step2Panel)

    $step2Num = New-Object System.Windows.Forms.Label
    $step2Num.Text = '2'
    $step2Num.Font = New-Object System.Drawing.Font('Segoe UI', 20, [System.Drawing.FontStyle]::Bold)
    $step2Num.ForeColor = [System.Drawing.Color]::White
    $step2Num.Location = [System.Drawing.Point]::new(10, 10)
    $step2Num.Size = [System.Drawing.Size]::new(40, 40)
    $step2Num.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $step2Panel.Controls.Add($step2Num)

    $step2Lbl = New-Object System.Windows.Forms.Label
    $step2Lbl.Text = "Run Azure AD Sync`nVOL-ane-aad1 delta sync"
    $step2Lbl.Font = $FontBody
    $step2Lbl.ForeColor = [System.Drawing.Color]::White
    $step2Lbl.Location = [System.Drawing.Point]::new(55, 10)
    $step2Lbl.Size = [System.Drawing.Size]::new(285, 40)
    $step2Panel.Controls.Add($step2Lbl)

    $stepX += $stepWidth + 12

    # Step 3
    $step3Panel = New-Object System.Windows.Forms.Panel
    $step3Panel.Location = [System.Drawing.Point]::new($stepX, 10)
    $step3Panel.Size = [System.Drawing.Size]::new($stepWidth, 60)
    $step3Panel.BackColor = $clrMuted
    $stepsPanel.Controls.Add($step3Panel)

    $step3Num = New-Object System.Windows.Forms.Label
    $step3Num.Text = '3'
    $step3Num.Font = New-Object System.Drawing.Font('Segoe UI', 20, [System.Drawing.FontStyle]::Bold)
    $step3Num.ForeColor = [System.Drawing.Color]::White
    $step3Num.Location = [System.Drawing.Point]::new(10, 10)
    $step3Num.Size = [System.Drawing.Size]::new(40, 40)
    $step3Num.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $step3Panel.Controls.Add($step3Num)

    $step3Lbl = New-Object System.Windows.Forms.Label
    $step3Lbl.Text = "Remove Domain from M365`nDelete verified domain"
    $step3Lbl.Font = $FontBody
    $step3Lbl.ForeColor = [System.Drawing.Color]::White
    $step3Lbl.Location = [System.Drawing.Point]::new(55, 10)
    $step3Lbl.Size = [System.Drawing.Size]::new(285, 40)
    $step3Panel.Controls.Add($step3Lbl)

    # ── Main Content Panel ────────────────────────────────────────────────────
    $card = New-Object System.Windows.Forms.Panel
    $card.Location = [System.Drawing.Point]::new(12, 156)
    $card.Size = [System.Drawing.Size]::new(1076, 550)
    $card.BackColor = $clrPanel
    $form.Controls.Add($card)

    $lx = 14
    $y = 12

    # ══════════════════════════════════════════════════════════════════════════
    # STEP 1: Update On-Prem UPN
    # ══════════════════════════════════════════════════════════════════════════
    $step1Content = New-Object System.Windows.Forms.Panel
    $step1Content.Location = [System.Drawing.Point]::new(0, 0)
    $step1Content.Size = [System.Drawing.Size]::new(1076, 550)
    $step1Content.BackColor = $clrPanel
    $card.Controls.Add($step1Content)

    $y = 12

    # CSV Folder
    $lblFolderCap = New-Object System.Windows.Forms.Label
    $lblFolderCap.Text = 'CSV FOLDER (Discovery output)'
    $lblFolderCap.Font = $FontCap
    $lblFolderCap.ForeColor = $clrMuted
    $lblFolderCap.Location = [System.Drawing.Point]::new($lx, $y)
    $lblFolderCap.AutoSize = $true
    $step1Content.Controls.Add($lblFolderCap)
    $y += 18

    $tbFolder = New-Object System.Windows.Forms.TextBox
    $tbFolder.Location = [System.Drawing.Point]::new($lx, $y)
    $tbFolder.Size = [System.Drawing.Size]::new(930, 24)
    $tbFolder.Font = $FontBody
    $tbFolder.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $step1Content.Controls.Add($tbFolder)

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = 'Browse...'
    $btnBrowse.Location = [System.Drawing.Point]::new(952, $y - 2)
    $btnBrowse.Size = [System.Drawing.Size]::new(104, 28)
    $btnBrowse.Font = $FontBold
    $btnBrowse.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnBrowse.FlatAppearance.BorderSize = 0
    $btnBrowse.BackColor = [System.Drawing.Color]::FromArgb(225, 228, 238)
    $btnBrowse.ForeColor = $clrText
    $btnBrowse.Cursor = [System.Windows.Forms.Cursors]::Hand
    $step1Content.Controls.Add($btnBrowse)
    $y += 32

    $btnScan = New-Object System.Windows.Forms.Button
    $btnScan.Text = 'Load CSV Files'
    $btnScan.Location = [System.Drawing.Point]::new($lx, $y)
    $btnScan.Size = [System.Drawing.Size]::new(130, 28)
    $btnScan.Font = $FontBold
    $btnScan.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnScan.FlatAppearance.BorderSize = 0
    $btnScan.BackColor = [System.Drawing.Color]::FromArgb(225, 228, 238)
    $btnScan.ForeColor = $clrText
    $btnScan.Cursor = [System.Windows.Forms.Cursors]::Hand
    $step1Content.Controls.Add($btnScan)

    $lblScanStatus = New-Object System.Windows.Forms.Label
    $lblScanStatus.Text = 'Select CSV folder containing UserPrincipalName data'
    $lblScanStatus.ForeColor = $clrMuted
    $lblScanStatus.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $lblScanStatus.Location = [System.Drawing.Point]::new(152, $y + 8)
    $lblScanStatus.AutoSize = $true
    $step1Content.Controls.Add($lblScanStatus)
    $y += 38

    $sep1 = New-Object System.Windows.Forms.Label
    $sep1.Location = [System.Drawing.Point]::new($lx, $y)
    $sep1.Size = [System.Drawing.Size]::new(1048, 1)
    $sep1.BackColor = $clrBorder
    $step1Content.Controls.Add($sep1)
    $y += 10

    # Domain info
    $lblDomainInfo = New-Object System.Windows.Forms.Label
    $lblDomainInfo.Text = 'SOURCE DOMAIN TO REMOVE'
    $lblDomainInfo.Font = $FontCap
    $lblDomainInfo.ForeColor = $clrMuted
    $lblDomainInfo.Location = [System.Drawing.Point]::new($lx, $y)
    $lblDomainInfo.AutoSize = $true
    $step1Content.Controls.Add($lblDomainInfo)
    $y += 18

    $tbSourceDomain = New-Object System.Windows.Forms.TextBox
    $tbSourceDomain.Location = [System.Drawing.Point]::new($lx, $y)
    $tbSourceDomain.Size = [System.Drawing.Size]::new(300, 24)
    $tbSourceDomain.Font = $FontBody
    $tbSourceDomain.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $tbSourceDomain.Text = '@olddomain.com'
    $step1Content.Controls.Add($tbSourceDomain)

    $lblArrow = New-Object System.Windows.Forms.Label
    $lblArrow.Text = '→'
    $lblArrow.Font = New-Object System.Drawing.Font('Segoe UI', 16)
    $lblArrow.ForeColor = $clrMuted
    $lblArrow.Location = [System.Drawing.Point]::new(322, $y - 2)
    $lblArrow.Size = [System.Drawing.Size]::new(30, 28)
    $lblArrow.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $step1Content.Controls.Add($lblArrow)

    $lblTarget = New-Object System.Windows.Forms.Label
    $lblTarget.Text = '@ourvolaris.onmicrosoft.com'
    $lblTarget.Font = $FontLarge
    $lblTarget.ForeColor = $clrAccent
    $lblTarget.Location = [System.Drawing.Point]::new(358, $y + 2)
    $lblTarget.AutoSize = $true
    $step1Content.Controls.Add($lblTarget)
    $y += 38

    $sep2 = New-Object System.Windows.Forms.Label
    $sep2.Location = [System.Drawing.Point]::new($lx, $y)
    $sep2.Size = [System.Drawing.Size]::new(1048, 1)
    $sep2.BackColor = $clrBorder
    $step1Content.Controls.Add($sep2)
    $y += 10

    # User List
    $lblUsersCap = New-Object System.Windows.Forms.Label
    $lblUsersCap.Text = 'USERS TO UPDATE'
    $lblUsersCap.Font = $FontCap
    $lblUsersCap.ForeColor = $clrMuted
    $lblUsersCap.Location = [System.Drawing.Point]::new($lx, $y)
    $lblUsersCap.AutoSize = $true
    $step1Content.Controls.Add($lblUsersCap)
    $y += 18

    $lvUsers = New-Object System.Windows.Forms.ListView
    $lvUsers.Location = [System.Drawing.Point]::new($lx, $y)
    $lvUsers.Size = [System.Drawing.Size]::new(1048, 240)
    $lvUsers.View = [System.Windows.Forms.View]::Details
    $lvUsers.FullRowSelect = $true
    $lvUsers.GridLines = $true
    $lvUsers.CheckBoxes = $true
    $lvUsers.Font = $FontBody
    $lvUsers.BackColor = [System.Drawing.Color]::FromArgb(250, 251, 253)
    $lvUsers.ForeColor = $clrText
    $lvUsers.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    [void]$lvUsers.Columns.Add('Display Name', 250)
    [void]$lvUsers.Columns.Add('Current UPN', 300)
    [void]$lvUsers.Columns.Add('New UPN', 350)
    [void]$lvUsers.Columns.Add('Status', 140)
    $step1Content.Controls.Add($lvUsers)
    $y += 245

    $lblUserCount = New-Object System.Windows.Forms.Label
    $lblUserCount.Text = '0 users loaded'
    $lblUserCount.ForeColor = $clrMuted
    $lblUserCount.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
    $lblUserCount.Location = [System.Drawing.Point]::new($lx, $y)
    $lblUserCount.AutoSize = $true
    $step1Content.Controls.Add($lblUserCount)

    $btnSelectAll = New-Object System.Windows.Forms.Button
    $btnSelectAll.Text = 'Select All'
    $btnSelectAll.Location = [System.Drawing.Point]::new(846, $y - 2)
    $btnSelectAll.Size = [System.Drawing.Size]::new(100, 26)
    $btnSelectAll.Font = $FontBold
    $btnSelectAll.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnSelectAll.FlatAppearance.BorderSize = 0
    $btnSelectAll.BackColor = [System.Drawing.Color]::FromArgb(225, 228, 238)
    $btnSelectAll.ForeColor = $clrText
    $btnSelectAll.Cursor = [System.Windows.Forms.Cursors]::Hand
    $step1Content.Controls.Add($btnSelectAll)

    $btnDeselectAll = New-Object System.Windows.Forms.Button
    $btnDeselectAll.Text = 'Deselect All'
    $btnDeselectAll.Location = [System.Drawing.Point]::new(956, $y - 2)
    $btnDeselectAll.Size = [System.Drawing.Size]::new(100, 26)
    $btnDeselectAll.Font = $FontBold
    $btnDeselectAll.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnDeselectAll.FlatAppearance.BorderSize = 0
    $btnDeselectAll.BackColor = [System.Drawing.Color]::FromArgb(225, 228, 238)
    $btnDeselectAll.ForeColor = $clrText
    $btnDeselectAll.Cursor = [System.Windows.Forms.Cursors]::Hand
    $step1Content.Controls.Add($btnDeselectAll)
    $y += 34

    $sep3 = New-Object System.Windows.Forms.Label
    $sep3.Location = [System.Drawing.Point]::new($lx, $y)
    $sep3.Size = [System.Drawing.Size]::new(1048, 1)
    $sep3.BackColor = $clrBorder
    $step1Content.Controls.Add($sep3)
    $y += 12

    # Run Step 1
    $btnStep1 = New-Object System.Windows.Forms.Button
    $btnStep1.Text = 'Step 1: Update On-Prem UPNs →'
    $btnStep1.Location = [System.Drawing.Point]::new(796, $y)
    $btnStep1.Size = [System.Drawing.Size]::new(260, 40)
    $btnStep1.Font = $FontLarge
    $btnStep1.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnStep1.FlatAppearance.BorderSize = 0
    $btnStep1.BackColor = $clrAccent
    $btnStep1.ForeColor = [System.Drawing.Color]::White
    $btnStep1.Cursor = [System.Windows.Forms.Cursors]::Hand
    $step1Content.Controls.Add($btnStep1)

    # ══════════════════════════════════════════════════════════════════════════
    # STEP 2: AD Sync
    # ══════════════════════════════════════════════════════════════════════════
    $step2Content = New-Object System.Windows.Forms.Panel
    $step2Content.Location = [System.Drawing.Point]::new(0, 0)
    $step2Content.Size = [System.Drawing.Size]::new(1076, 550)
    $step2Content.BackColor = $clrPanel
    $step2Content.Visible = $false
    $card.Controls.Add($step2Content)

    $y = 40

    $lblStep2Title = New-Object System.Windows.Forms.Label
    $lblStep2Title.Text = 'Azure AD Connect Sync'
    $lblStep2Title.Font = New-Object System.Drawing.Font('Segoe UI', 18, [System.Drawing.FontStyle]::Bold)
    $lblStep2Title.ForeColor = $clrAccent
    $lblStep2Title.Location = [System.Drawing.Point]::new($lx, $y)
    $lblStep2Title.AutoSize = $true
    $step2Content.Controls.Add($lblStep2Title)
    $y += 50

    $lblStep2Info = New-Object System.Windows.Forms.Label
    $lblStep2Info.Text = "On-premise UPN changes have been applied.`n`nNow we need to sync these changes to Microsoft 365 (Entra ID).`n`nThis will trigger a delta synchronization on VOL-ane-aad1."
    $lblStep2Info.Font = $FontBody
    $lblStep2Info.ForeColor = $clrText
    $lblStep2Info.Location = [System.Drawing.Point]::new($lx, $y)
    $lblStep2Info.Size = [System.Drawing.Size]::new(1000, 80)
    $step2Content.Controls.Add($lblStep2Info)
    $y += 100

    $lblStep2Status = New-Object System.Windows.Forms.Label
    $lblStep2Status.Text = 'Ready to run sync'
    $lblStep2Status.Font = $FontLarge
    $lblStep2Status.ForeColor = $clrMuted
    $lblStep2Status.Location = [System.Drawing.Point]::new($lx, $y)
    $lblStep2Status.Size = [System.Drawing.Size]::new(1000, 30)
    $step2Content.Controls.Add($lblStep2Status)
    $y += 50

    $btnStep2 = New-Object System.Windows.Forms.Button
    $btnStep2.Text = 'Step 2: Run Azure AD Sync →'
    $btnStep2.Location = [System.Drawing.Point]::new(($step2Content.Width - 280) / 2, $y)
    $btnStep2.Size = [System.Drawing.Size]::new(280, 50)
    $btnStep2.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 13)
    $btnStep2.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnStep2.FlatAppearance.BorderSize = 0
    $btnStep2.BackColor = $clrGreen
    $btnStep2.ForeColor = [System.Drawing.Color]::White
    $btnStep2.Cursor = [System.Windows.Forms.Cursors]::Hand
    $step2Content.Controls.Add($btnStep2)
    $y += 70

    $btnSkipStep2 = New-Object System.Windows.Forms.Button
    $btnSkipStep2.Text = 'Skip (already synced manually) →'
    $btnSkipStep2.Location = [System.Drawing.Point]::new(($step2Content.Width - 260) / 2, $y)
    $btnSkipStep2.Size = [System.Drawing.Size]::new(260, 32)
    $btnSkipStep2.Font = $FontBody
    $btnSkipStep2.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnSkipStep2.FlatAppearance.BorderSize = 0
    $btnSkipStep2.BackColor = [System.Drawing.Color]::FromArgb(225, 228, 238)
    $btnSkipStep2.ForeColor = $clrText
    $btnSkipStep2.Cursor = [System.Windows.Forms.Cursors]::Hand
    $step2Content.Controls.Add($btnSkipStep2)

    # ══════════════════════════════════════════════════════════════════════════
    # STEP 3: Remove Domain
    # ══════════════════════════════════════════════════════════════════════════
    $step3Content = New-Object System.Windows.Forms.Panel
    $step3Content.Location = [System.Drawing.Point]::new(0, 0)
    $step3Content.Size = [System.Drawing.Size]::new(1076, 550)
    $step3Content.BackColor = $clrPanel
    $step3Content.Visible = $false
    $card.Controls.Add($step3Content)

    $y = 40

    $lblStep3Title = New-Object System.Windows.Forms.Label
    $lblStep3Title.Text = 'Remove Domain from Microsoft 365'
    $lblStep3Title.Font = New-Object System.Drawing.Font('Segoe UI', 18, [System.Drawing.FontStyle]::Bold)
    $lblStep3Title.ForeColor = $clrAccent
    $lblStep3Title.Location = [System.Drawing.Point]::new($lx, $y)
    $lblStep3Title.AutoSize = $true
    $step3Content.Controls.Add($lblStep3Title)
    $y += 50

    $lblStep3Info = New-Object System.Windows.Forms.Label
    $lblStep3Info.Text = "UPNs have been updated and synced to Microsoft 365.`n`nNow we can safely remove the domain from your tenant.`n`nDomain to remove:"
    $lblStep3Info.Font = $FontBody
    $lblStep3Info.ForeColor = $clrText
    $lblStep3Info.Location = [System.Drawing.Point]::new($lx, $y)
    $lblStep3Info.Size = [System.Drawing.Size]::new(1000, 80)
    $step3Content.Controls.Add($lblStep3Info)
    $y += 100

    $lblDomainToRemove = New-Object System.Windows.Forms.Label
    $lblDomainToRemove.Text = '@olddomain.com'
    $lblDomainToRemove.Font = New-Object System.Drawing.Font('Segoe UI', 20, [System.Drawing.FontStyle]::Bold)
    $lblDomainToRemove.ForeColor = $clrRed
    $lblDomainToRemove.Location = [System.Drawing.Point]::new($lx, $y)
    $lblDomainToRemove.Size = [System.Drawing.Size]::new(1000, 40)
    $lblDomainToRemove.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $step3Content.Controls.Add($lblDomainToRemove)
    $y += 60

    $lblStep3Status = New-Object System.Windows.Forms.Label
    $lblStep3Status.Text = 'Ready to remove domain'
    $lblStep3Status.Font = $FontLarge
    $lblStep3Status.ForeColor = $clrMuted
    $lblStep3Status.Location = [System.Drawing.Point]::new($lx, $y)
    $lblStep3Status.Size = [System.Drawing.Size]::new(1000, 30)
    $lblStep3Status.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $step3Content.Controls.Add($lblStep3Status)
    $y += 50

    $btnStep3 = New-Object System.Windows.Forms.Button
    $btnStep3.Text = 'Step 3: Remove Domain'
    $btnStep3.Location = [System.Drawing.Point]::new(($step3Content.Width - 280) / 2, $y)
    $btnStep3.Size = [System.Drawing.Size]::new(280, 50)
    $btnStep3.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 13)
    $btnStep3.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnStep3.FlatAppearance.BorderSize = 0
    $btnStep3.BackColor = $clrRed
    $btnStep3.ForeColor = [System.Drawing.Color]::White
    $btnStep3.Cursor = [System.Windows.Forms.Cursors]::Hand
    $step3Content.Controls.Add($btnStep3)

    # ── Progress bar ──────────────────────────────────────────────────────────
    $progress = New-Object System.Windows.Forms.ProgressBar
    $progress.Location = [System.Drawing.Point]::new(12, $card.Bottom + 8)
    $progress.Size = [System.Drawing.Size]::new(1076, 8)
    $progress.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $form.Controls.Add($progress)

    # ── Log RTB ───────────────────────────────────────────────────────────────
    $script:rtbLog = New-Object System.Windows.Forms.RichTextBox
    $script:rtbLog.Location = [System.Drawing.Point]::new(12, $progress.Bottom + 8)
    $script:rtbLog.Size = [System.Drawing.Size]::new(1076, 120)
    $script:rtbLog.BackColor = $clrLogBg
    $script:rtbLog.ForeColor = [System.Drawing.Color]::FromArgb(200, 210, 230)
    $script:rtbLog.Font = $FontMono
    $script:rtbLog.ReadOnly = $true
    $script:rtbLog.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $script:rtbLog.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    $form.Controls.Add($script:rtbLog)

    # ── Helper: Update Step Indicators ───────────────────────────────────────
    $updateStepIndicators = {
        param([int]$CurrentStep)

        # Reset all to muted
        $step1Panel.BackColor = $clrMuted
        $step2Panel.BackColor = $clrMuted
        $step3Panel.BackColor = $clrMuted

        # Highlight current
        switch ($CurrentStep) {
            1 { $step1Panel.BackColor = $clrAccent }
            2 { $step2Panel.BackColor = $clrGreen }
            3 { $step3Panel.BackColor = $clrRed }
        }

        # Show/hide content panels
        $step1Content.Visible = ($CurrentStep -eq 1)
        $step2Content.Visible = ($CurrentStep -eq 2)
        $step3Content.Visible = ($CurrentStep -eq 3)
    }

    # ── Event Handlers ────────────────────────────────────────────────────────

    # Browse
    $btnBrowse.Add_Click({
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description = 'Select the folder containing CSV files'
        $fbd.ShowNewFolderButton = $false
        if ($tbFolder.Text -and (Test-Path $tbFolder.Text)) {
            $fbd.SelectedPath = $tbFolder.Text
        }
        if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $tbFolder.Text = $fbd.SelectedPath
            Write-Log "Folder selected: $($fbd.SelectedPath)"
        }
    })

    # Load CSVs
    $btnScan.Add_Click({
        $folder = $tbFolder.Text.Trim().Trim('"')
        if (-not $folder -or -not (Test-Path $folder)) {
            [System.Windows.Forms.MessageBox]::Show('Please select a valid folder.', 'Invalid Folder', 'OK', 'Warning') | Out-Null
            return
        }

        # Check for ActiveDirectory module
        if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
            [System.Windows.Forms.MessageBox]::Show(
                "ActiveDirectory module not found.`n`nThis step requires the Active Directory PowerShell module.`nInstall RSAT or run from a Domain Controller.",
                'Module Required',
                'OK',
                'Error'
            ) | Out-Null
            return
        }

        Write-Log "=== Loading CSV files from: $folder ==="
        $lblScanStatus.Text = 'Loading CSV files...'
        [System.Windows.Forms.Application]::DoEvents()

        $csvFiles = Get-ChildItem -Path $folder -Filter *.csv -File
        if ($csvFiles.Count -eq 0) {
            Write-Log "No CSV files found in: $folder" 'WARN'
            $lblScanStatus.Text = 'No CSV files found'
            return
        }

        Write-Log "Found $($csvFiles.Count) CSV file(s)"
        $lvUsers.Items.Clear()
        $script:AllUsers.Clear()

        foreach ($csvFile in $csvFiles) {
            Write-Log "  Loading: $($csvFile.Name)"
            try {
                $rows = Import-Csv -Path $csvFile.FullName -Encoding UTF8

                if (-not ($rows[0].PSObject.Properties.Name -contains 'UserPrincipalName')) {
                    Write-Log "    Skipped: No UserPrincipalName column" 'WARN'
                    continue
                }

                foreach ($row in $rows) {
                    if ($row.UserPrincipalName) {
                        $script:AllUsers.Add([pscustomobject]@{
                            UserPrincipalName = $row.UserPrincipalName
                            DisplayName = if ($row.DisplayName) { $row.DisplayName } else { $row.UserPrincipalName.Split('@')[0] }
                            CSVFile = $csvFile.Name
                        })
                    }
                }
            } catch {
                Write-Log "    Error reading $($csvFile.Name): $_" 'ERROR'
            }
        }

        Write-Log "Loaded $($script:AllUsers.Count) user(s) from CSV files"

        # Populate ListView with new UPN preview
        $src = $tbSourceDomain.Text.Trim()
        if (-not $src.StartsWith('@')) { $src = '@' + $src }

        foreach ($user in $script:AllUsers) {
            $currentUPN = $user.UserPrincipalName
            $newUPN = if ($currentUPN -like "*$src") {
                $currentUPN -replace [regex]::Escape($src) + '$', $script:TargetDomain
            } else {
                $currentUPN
            }

            $lvi = New-Object System.Windows.Forms.ListViewItem($user.DisplayName)
            $lvi.Checked = ($currentUPN -like "*$src")
            [void]$lvi.SubItems.Add($currentUPN)
            [void]$lvi.SubItems.Add($newUPN)
            [void]$lvi.SubItems.Add('Ready')
            $lvi.Tag = $user
            [void]$lvUsers.Items.Add($lvi)
        }

        $checkedCount = ($lvUsers.Items | Where-Object { $_.Checked }).Count
        $lblUserCount.Text = "$($script:AllUsers.Count) user(s) loaded, $checkedCount selected"
        $lblScanStatus.Text = "Loaded $($script:AllUsers.Count) user(s) from $($csvFiles.Count) CSV file(s)"
    })

    # Update New UPN when source domain changes
    $tbSourceDomain.Add_TextChanged({
        $src = $tbSourceDomain.Text.Trim()
        if (-not $src) { return }
        if (-not $src.StartsWith('@')) { $src = '@' + $src }

        foreach ($lvi in $lvUsers.Items) {
            $currentUPN = $lvi.SubItems[1].Text
            if ($currentUPN -like "*$src") {
                $newUPN = $currentUPN -replace [regex]::Escape($src) + '$', $script:TargetDomain
                $lvi.SubItems[2].Text = $newUPN
                $lvi.Checked = $true
            } else {
                $lvi.SubItems[2].Text = $currentUPN
                $lvi.Checked = $false
            }
        }

        $lblDomainToRemove.Text = $src
    })

    # Select/Deselect All
    $btnSelectAll.Add_Click({
        foreach ($lvi in $lvUsers.Items) { $lvi.Checked = $true }
        $checked = ($lvUsers.Items | Where-Object { $_.Checked }).Count
        $lblUserCount.Text = "$($lvUsers.Items.Count) user(s) loaded, $checked selected"
    })

    $btnDeselectAll.Add_Click({
        foreach ($lvi in $lvUsers.Items) { $lvi.Checked = $false }
        $lblUserCount.Text = "$($lvUsers.Items.Count) user(s) loaded, 0 selected"
    })

    $lvUsers.Add_ItemChecked({
        $checked = ($lvUsers.Items | Where-Object { $_.Checked }).Count
        $lblUserCount.Text = "$($lvUsers.Items.Count) user(s) loaded, $checked selected"
    })

    # Step 1: Update On-Prem UPNs
    $btnStep1.Add_Click({
        $selectedUsers = @($lvUsers.Items | Where-Object { $_.Checked })
        if ($selectedUsers.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('Please select at least one user.', 'No Users Selected', 'OK', 'Warning') | Out-Null
            return
        }

        $src = $tbSourceDomain.Text.Trim()
        if (-not $src.StartsWith('@')) { $src = '@' + $src }

        # Check ActiveDirectory module
        if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
            [System.Windows.Forms.MessageBox]::Show(
                "ActiveDirectory module not found.`n`nPlease run this from a machine with RSAT tools or a Domain Controller.",
                'Module Required',
                'OK',
                'Error'
            ) | Out-Null
            return
        }

        Import-Module ActiveDirectory -ErrorAction Stop

        $result = [System.Windows.Forms.MessageBox]::Show(
            "Update $($selectedUsers.Count) user(s) to $($script:TargetDomain)?`n`nThis will update on-premise Active Directory.`n`nSource: $src`nTarget: $($script:TargetDomain)",
            'Confirm Step 1',
            'YesNo',
            'Question'
        )

        if ($result -ne 'Yes') { return }

        Write-Log "=== STEP 1: Update On-Premise UPNs ===" 'OK'
        Write-Log "Source: $src → Target: $($script:TargetDomain)"
        Write-Log "Users to update: $($selectedUsers.Count)"

        $btnStep1.Enabled = $false
        $btnScan.Enabled = $false
        $btnBrowse.Enabled = $false
        $progress.Value = 0

        $success = 0
        $failed = 0
        $skipped = 0

        for ($i = 0; $i -lt $selectedUsers.Count; $i++) {
            $lvi = $selectedUsers[$i]
            $currentUPN = $lvi.SubItems[1].Text

            $pct = [int](($i / $selectedUsers.Count) * 100)
            try { $progress.Value = $pct } catch { }

            if ($currentUPN -notlike "*$src") {
                Write-Log "SKIP: $currentUPN" 'WARN'
                $lvi.SubItems[3].Text = 'Skipped'
                $skipped++
                continue
            }

            $newUPN = $currentUPN -replace [regex]::Escape($src) + '$', $script:TargetDomain

            try {
                $adUser = Get-ADUser -Filter "UserPrincipalName -eq '$currentUPN'" -Properties proxyAddresses, mail -ErrorAction Stop

                if (-not $adUser) {
                    Write-Log "NOT FOUND: $currentUPN" 'ERROR'
                    $lvi.SubItems[3].Text = 'Not Found'
                    $failed++
                    continue
                }

                # Update UPN
                Set-ADUser -Identity $adUser -UserPrincipalName $newUPN -ErrorAction Stop

                # Update mail
                Set-ADUser -Identity $adUser -EmailAddress $newUPN -ErrorAction Stop

                # Update proxyAddresses
                if ($adUser.proxyAddresses) {
                    $newProxyAddresses = @()
                    foreach ($proxy in $adUser.proxyAddresses) {
                        if ($proxy -like "*$src") {
                            $newProxy = $proxy -replace [regex]::Escape($src) + '$', $script:TargetDomain
                            $newProxyAddresses += $newProxy
                        } else {
                            $newProxyAddresses += $proxy
                        }
                    }
                    Set-ADUser -Identity $adUser -Replace @{proxyAddresses = $newProxyAddresses} -ErrorAction Stop
                }

                Write-Log "SUCCESS: $currentUPN → $newUPN" 'OK'
                $lvi.SubItems[3].Text = 'Updated'
                $lvi.SubItems[1].Text = $newUPN
                $success++
            } catch {
                Write-Log "FAILED: $currentUPN - $_" 'ERROR'
                $lvi.SubItems[3].Text = 'Failed'
                $failed++
            }

            [System.Windows.Forms.Application]::DoEvents()
        }

        try { $progress.Value = 100 } catch { }

        Write-Log "=== Step 1 Complete ===" 'OK'
        Write-Log "Success: $success  |  Failed: $failed  |  Skipped: $skipped"

        $btnStep1.Enabled = $true
        $btnScan.Enabled = $true
        $btnBrowse.Enabled = $true

        if ($success -gt 0) {
            $result = [System.Windows.Forms.MessageBox]::Show(
                "Step 1 complete.`n`nSuccess: $success`nFailed: $failed`nSkipped: $skipped`n`nProceed to Step 2 (Azure AD Sync)?",
                'Step 1 Complete',
                'YesNo',
                'Information'
            )

            if ($result -eq 'Yes') {
                $script:CurrentStep = 2
                & $updateStepIndicators 2
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "No users were updated.`n`nFailed: $failed`nSkipped: $skipped",
                'Step 1 Complete',
                'OK',
                'Warning'
            ) | Out-Null
        }
    })

    # Step 2: Run AD Sync
    $btnStep2.Add_Click({
        Write-Log "=== STEP 2: Azure AD Sync ===" 'OK'

        $result = [System.Windows.Forms.MessageBox]::Show(
            "Trigger Azure AD Connect sync on VOL-ane-aad1?`n`nThis will start a delta sync cycle.",
            'Confirm Step 2',
            'YesNo',
            'Question'
        )

        if ($result -ne 'Yes') { return }

        try {
            $btnStep2.Enabled = $false
            $btnStep2.Text = 'Syncing...'
            $lblStep2Status.Text = 'Connecting to VOL-ane-aad1...'
            $lblStep2Status.ForeColor = $clrAccent
            [System.Windows.Forms.Application]::DoEvents()

            Write-Log "Connecting to VOL-ane-aad1..."
            $syncScript = {
                Import-Module ADSync -ErrorAction Stop
                Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Stop
            }

            $session = New-PSSession -ComputerName 'VOL-ane-aad1' -ErrorAction Stop
            Write-Log "Remote session established. Invoking sync..."
            $lblStep2Status.Text = 'Running delta sync...'
            [System.Windows.Forms.Application]::DoEvents()

            $syncResult = Invoke-Command -Session $session -ScriptBlock $syncScript -ErrorAction Stop
            Remove-PSSession $session

            Write-Log "AD Sync completed: $($syncResult.Result)" 'OK'
            $lblStep2Status.Text = "Sync complete: $($syncResult.Result)"
            $lblStep2Status.ForeColor = $clrGreen

            $result = [System.Windows.Forms.MessageBox]::Show(
                "Step 2 complete.`n`nResult: $($syncResult.Result)`n`nChanges are now syncing to Microsoft 365.`n`nProceed to Step 3 (Remove Domain)?",
                'Step 2 Complete',
                'YesNo',
                'Information'
            )

            if ($result -eq 'Yes') {
                $script:CurrentStep = 3
                & $updateStepIndicators 3
            }
        } catch {
            Write-Log "AD Sync failed: $_" 'ERROR'
            $lblStep2Status.Text = 'Sync failed - see log'
            $lblStep2Status.ForeColor = $clrRed
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to run AD Sync:`n`n$_",
                'Step 2 Error',
                'OK',
                'Error'
            ) | Out-Null
        } finally {
            $btnStep2.Enabled = $true
            $btnStep2.Text = 'Step 2: Run Azure AD Sync →'
        }
    })

    # Skip Step 2
    $btnSkipStep2.Add_Click({
        Write-Log "Step 2 skipped by user" 'WARN'
        $script:CurrentStep = 3
        & $updateStepIndicators 3
    })

    # Step 3: Remove Domain
    $btnStep3.Add_Click({
        $domainToRemove = $tbSourceDomain.Text.Trim()
        if (-not $domainToRemove.StartsWith('@')) { $domainToRemove = '@' + $domainToRemove }
        $domainToRemove = $domainToRemove.TrimStart('@')

        Write-Log "=== STEP 3: Remove Domain ===" 'OK'
        Write-Log "Domain: $domainToRemove"

        # Check Microsoft.Graph module
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Identity.DirectoryManagement)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Microsoft.Graph.Identity.DirectoryManagement module not found.`n`nInstall with:`nInstall-Module Microsoft.Graph.Identity.DirectoryManagement",
                'Module Required',
                'OK',
                'Error'
            ) | Out-Null
            return
        }

        $result = [System.Windows.Forms.MessageBox]::Show(
            "REMOVE DOMAIN: $domainToRemove`n`nThis will permanently delete the domain from your Microsoft 365 tenant.`n`nAre you sure?",
            'Confirm Step 3',
            'YesNo',
            'Warning'
        )

        if ($result -ne 'Yes') { return }

        try {
            $btnStep3.Enabled = $false
            $btnStep3.Text = 'Removing...'
            $lblStep3Status.Text = 'Connecting to Microsoft Graph...'
            $lblStep3Status.ForeColor = $clrAccent
            [System.Windows.Forms.Application]::DoEvents()

            Write-Log "Connecting to Microsoft Graph..."
            Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
            Connect-MgGraph -Scopes 'Domain.ReadWrite.All' -NoWelcome -ErrorAction Stop

            $lblStep3Status.Text = "Removing domain: $domainToRemove..."
            [System.Windows.Forms.Application]::DoEvents()

            Write-Log "Removing domain: $domainToRemove"
            Remove-MgDomain -DomainId $domainToRemove -ErrorAction Stop

            Write-Log "Domain removed successfully!" 'OK'
            $lblStep3Status.Text = 'Domain removed successfully!'
            $lblStep3Status.ForeColor = $clrGreen

            Disconnect-MgGraph | Out-Null

            [System.Windows.Forms.MessageBox]::Show(
                "Domain removal complete!`n`n$domainToRemove has been removed from Microsoft 365.`n`nWorkflow finished.",
                'Workflow Complete',
                'OK',
                'Information'
            ) | Out-Null
        } catch {
            Write-Log "Domain removal failed: $_" 'ERROR'
            $lblStep3Status.Text = 'Domain removal failed - see log'
            $lblStep3Status.ForeColor = $clrRed
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to remove domain:`n`n$_`n`nEnsure all users/groups have been migrated and no objects reference this domain.",
                'Step 3 Error',
                'OK',
                'Error'
            ) | Out-Null
        } finally {
            $btnStep3.Enabled = $true
            $btnStep3.Text = 'Step 3: Remove Domain'
        }
    })

    Write-Log "UI initialized. Ready to start workflow."
    $form.ShowDialog() | Out-Null
}

# ═════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═════════════════════════════════════════════════════════════════════════════
Show-DomainRemovalWorkflowUI
