Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ── P/INVOKE TYPES — compiled once, cached to disk; subsequent launches skip the C# compiler ─
$_nativeCachePath = Join-Path $env:LOCALAPPDATA 'FlyMigration\NativeHelpers.dll'
$_nativeLoaded    = $false

if (Test-Path $_nativeCachePath) {
    try { [System.Reflection.Assembly]::LoadFrom($_nativeCachePath) | Out-Null; $_nativeLoaded = $true }
    catch { Remove-Item $_nativeCachePath -Force -ErrorAction SilentlyContinue }
}

if (-not $_nativeLoaded) {
    $_nativeDef = @'
using System; using System.Runtime.InteropServices;
namespace Win32 { public class DpiHelper {
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware(); } }
namespace FlyConsole { public class NativeMethods {
    [DllImport("user32.dll")]   public static extern bool  ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]   public static extern bool  AllowSetForegroundWindow(int dwProcessId); } }
namespace FlyMigration { public class Taskbar {
    [DllImport("shell32.dll")]
    public static extern int SetCurrentProcessExplicitAppUserModelID(
        [MarshalAs(UnmanagedType.LPWStr)] string AppID); } }
'@
    $_cacheDir = Split-Path $_nativeCachePath
    if (-not (Test-Path $_cacheDir)) { New-Item -ItemType Directory -Path $_cacheDir -Force | Out-Null }
    try   { Add-Type -TypeDefinition $_nativeDef -OutputAssembly $_nativeCachePath -ErrorAction Stop }
    catch { try { Add-Type -TypeDefinition $_nativeDef -ErrorAction SilentlyContinue } catch {} }
}

try { [Win32.DpiHelper]::SetProcessDPIAware() | Out-Null } catch {}
[System.Windows.Forms.Application]::EnableVisualStyles()

try {
    $_consoleHwnd = [FlyConsole.NativeMethods]::GetConsoleWindow()
    if ($_consoleHwnd -ne [IntPtr]::Zero) {
        [FlyConsole.NativeMethods]::ShowWindow($_consoleHwnd, 0) | Out-Null  # SW_HIDE
    }
} catch {}

try { [FlyMigration.Taskbar]::SetCurrentProcessExplicitAppUserModelID('AvePoint.FlyMigration') | Out-Null } catch {}

# ── HIDDEN PROCESS LAUNCHER ───────────────────────────────────────────────────
# CREATE_NO_WINDOW prevents any console window being allocated for the child
# powershell.exe process. UseShellExecute=false with default (Normal) WindowStyle
# means STARTUPINFO.dwFlags does NOT include STARTF_USESHOWWINDOW, so the OS
# never applies an SW_HIDE override to the WinForms form's first ShowWindow call.
function Start-HiddenProcess {
    param([string]$Exe, [string]$Arguments)
    if (-not [System.IO.Path]::IsPathRooted($Exe)) {
        $cmd = Get-Command $Exe -ErrorAction SilentlyContinue
        if ($cmd) { $Exe = $cmd.Source }
    }
    $psi                  = New-Object System.Diagnostics.ProcessStartInfo($Exe)
    $psi.Arguments        = $Arguments
    $psi.UseShellExecute  = $false
    $psi.CreateNoWindow   = $true
    [System.Diagnostics.Process]::Start($psi) | Out-Null
}

# ── SHARED COLOURS & FONTS ────────────────────────────────────────────────────
# Design System: Fly Migration Toolkit — matched to design-system/colors_and_type.css

# Base palette
$clrBg     = [System.Drawing.Color]::FromArgb(240, 242, 247)  # --fly-bg #f0f2f7
$clrPanel  = [System.Drawing.Color]::White                     # --fly-panel
$clrAccent = [System.Drawing.Color]::FromArgb(0, 100, 180)    # --fly-accent #0064b4 AvePoint blue
$clrAccentHover = [System.Drawing.Color]::FromArgb(0, 78, 152) # --fly-accent-hover #004e98
$clrAccentTint  = [System.Drawing.Color]::FromArgb(220, 230, 248) # --fly-accent-tint #dce6f8

$clrText   = [System.Drawing.Color]::FromArgb(28, 28, 32)     # --fly-text #1c1c20
$clrMuted  = [System.Drawing.Color]::FromArgb(100, 108, 120)  # --fly-muted #646c78
$clrBorder = [System.Drawing.Color]::FromArgb(210, 215, 228)  # --fly-border #d2d7e4
$clrGrey   = [System.Drawing.Color]::FromArgb(175, 182, 195)  # --fly-grey #afb6c3

# Dark surfaces
$clrLogBg  = [System.Drawing.Color]::FromArgb(26, 27, 38)     # --fly-log-bg #1a1b26
$clrFooter = [System.Drawing.Color]::FromArgb(20, 24, 38)     # --fly-footer #141826
$clrFooterAlt = [System.Drawing.Color]::FromArgb(28, 32, 48)  # --fly-footer-alt #1c2030

# Status colors
$clrGreen  = [System.Drawing.Color]::FromArgb(18, 155, 60)    # --fly-green #129b3c
$clrAmber  = [System.Drawing.Color]::FromArgb(195, 135, 0)    # --fly-amber #c38700
$clrRed    = [System.Drawing.Color]::FromArgb(195, 30, 30)    # --fly-red #c31e1e
$clrCloseRed = [System.Drawing.Color]::FromArgb(200, 55, 55)  # --fly-close-red #c83737
$clrBannerWarn = [System.Drawing.Color]::FromArgb(255, 243, 205) # --fly-banner-warn #fff3cd

# Log console text colors
$clrLogTimestamp = [System.Drawing.Color]::FromArgb(80, 95, 120)   # --fly-log-ts #505f78
$clrLogInfo      = [System.Drawing.Color]::FromArgb(120, 155, 220) # --fly-log-info #789bdc
$clrLogOK        = [System.Drawing.Color]::FromArgb(65, 195, 110)  # --fly-log-ok #41c36e
$clrLogWarn      = [System.Drawing.Color]::FromArgb(220, 165, 45)  # --fly-log-warn #dca52d
$clrLogError     = [System.Drawing.Color]::FromArgb(225, 80, 80)   # --fly-log-error #e15050
$clrLogBody      = [System.Drawing.Color]::FromArgb(205, 212, 230) # --fly-log-body #cdd4e6
$clrGridText     = [System.Drawing.Color]::FromArgb(190, 210, 255) # --fly-grid-text #bed2ff
$clrGridLine     = [System.Drawing.Color]::FromArgb(45, 55, 75)    # --fly-grid-line #2d374b

# Dark grid row tints
$clrRowFailBg    = [System.Drawing.Color]::FromArgb(70, 25, 25)    # --fly-row-fail-bg #461919
$clrRowFailFg    = [System.Drawing.Color]::FromArgb(240, 125, 125) # --fly-row-fail-fg #f07d7d
$clrRowWarnBg    = [System.Drawing.Color]::FromArgb(65, 48, 12)    # --fly-row-warn-bg #41300c
$clrRowWarnFg    = [System.Drawing.Color]::FromArgb(235, 195, 80)  # --fly-row-warn-fg #ebc350

# Typography — Segoe UI system font, sizes matched to design system
$FontBody  = New-Object System.Drawing.Font("Segoe UI", 9)          # --text-body 13px (9pt)
$FontBold  = New-Object System.Drawing.Font("Segoe UI Semibold", 9) # --text-bold 13px semibold
$FontCap   = New-Object System.Drawing.Font("Segoe UI Semibold", 7.5) # --text-cap 10px (7.5pt)
$FontMono  = New-Object System.Drawing.Font("Consolas", 8.5)        # --text-mono 12px (8.5pt)
$FontTitle = New-Object System.Drawing.Font("Segoe UI Semibold", 14) # --text-title 19px (14pt)
$FontSub   = New-Object System.Drawing.Font("Segoe UI", 8.5)        # --text-sub 12px tile subtitle
$FontTile  = New-Object System.Drawing.Font("Segoe UI Semibold", 14) # --text-tile 19px large nav tiles

$AnchorTL  = [System.Windows.Forms.AnchorStyles]::Top  -bor [System.Windows.Forms.AnchorStyles]::Left
$AnchorTLR = [System.Windows.Forms.AnchorStyles]::Top  -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$AnchorTR  = [System.Windows.Forms.AnchorStyles]::Top  -bor [System.Windows.Forms.AnchorStyles]::Right

# ── GEAR BUTTON ICON ──────────────────────────────────────────────────────────
# Draw a white circle with the banner-blue gear inside — matches the circular
# settings icon style and sits correctly on the accent-coloured banner.
$script:GearBitmap = $null
try {
    $gSz  = 28   # bitmap size; button is 38x38 so this gives ~5px padding all round
    $gBmp = New-Object System.Drawing.Bitmap($gSz, $gSz, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $gGfx = [System.Drawing.Graphics]::FromImage($gBmp)
    $gGfx.SmoothingMode    = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $gGfx.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $gGfx.Clear([System.Drawing.Color]::Transparent)

    # Filled white circle
    $gCircleBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
    $gGfx.FillEllipse($gCircleBrush, 0, 0, $gSz - 1, $gSz - 1)
    $gCircleBrush.Dispose()

    # Banner-blue gear glyph centred inside the circle
    $gGearFont  = New-Object System.Drawing.Font("Segoe UI Symbol", 14, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)
    $gGearBrush = New-Object System.Drawing.SolidBrush($clrAccent)
    $gSF = New-Object System.Drawing.StringFormat
    $gSF.Alignment     = [System.Drawing.StringAlignment]::Center
    $gSF.LineAlignment = [System.Drawing.StringAlignment]::Center
    $gGfx.DrawString([char]0x2699, $gGearFont, $gGearBrush, [System.Drawing.RectangleF]::new(0, 1, $gSz, $gSz), $gSF)
    $gGearFont.Dispose(); $gGearBrush.Dispose(); $gSF.Dispose()

    $gGfx.Dispose()
    $script:GearBitmap = $gBmp
} catch {}

# ── SHARED CONFIG ─────────────────────────────────────────────────────────────
$script:SharedConfigPath = Join-Path $env:LOCALAPPDATA "FlyMigration\shared-config.json"

function Read-SharedConfig {
    if (-not (Test-Path $script:SharedConfigPath)) { return [pscustomobject]@{} }
    try { return Get-Content $script:SharedConfigPath -Raw | ConvertFrom-Json }
    catch { return [pscustomobject]@{} }
}

function Update-SharedConfig {
    param([hashtable]$Values)
    $dir = Split-Path $script:SharedConfigPath
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $cfg = Read-SharedConfig
    foreach ($k in $Values.Keys) {
        if ($cfg.PSObject.Properties[$k]) { $cfg.$k = $Values[$k] }
        else { $cfg | Add-Member -NotePropertyName $k -NotePropertyValue $Values[$k] }
    }
    $cfg | ConvertTo-Json | Set-Content $script:SharedConfigPath -Encoding UTF8
}

# ── LOGO HELPER ───────────────────────────────────────────────────────────────
function Add-HeaderLogo {
    param($Header, [int]$LogoH = 34)
    $icoPath = Join-Path $PSScriptRoot "FlyMigration.ico"
    $pngPath = Join-Path $PSScriptRoot "ourvolaris.png"

    $img = $null
    if (Test-Path $icoPath) {
        try {
            # Read into a MemoryStream so the file isn't locked, then load at native
            # size (no size arg) to avoid the garbage-pixel bug when the ICO contains
            # a PNG-compressed 256x256 frame and a specific small size is requested.
            $bytes = [System.IO.File]::ReadAllBytes($icoPath)
            $ms    = New-Object System.IO.MemoryStream(,$bytes)
            $icon  = New-Object System.Drawing.Icon($ms)
            $raw   = $icon.ToBitmap()
            $icon.Dispose(); $ms.Dispose()
            # Scale to exactly LogoH x LogoH so the PictureBox never stretches it.
            $img = New-Object System.Drawing.Bitmap($LogoH, $LogoH, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
            $gfx = [System.Drawing.Graphics]::FromImage($img)
            $gfx.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $gfx.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
            $gfx.Clear([System.Drawing.Color]::Transparent)
            $gfx.DrawImage($raw, 0, 0, $LogoH, $LogoH)
            $gfx.Dispose(); $raw.Dispose()
        } catch { $img = $null }
    }
    if (-not $img -and (Test-Path $pngPath)) {
        try {
            $raw  = [System.Drawing.Image]::FromFile($pngPath)
            $img  = New-Object System.Drawing.Bitmap($LogoH, $LogoH, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
            $gfx  = [System.Drawing.Graphics]::FromImage($img)
            $gfx.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $gfx.Clear([System.Drawing.Color]::Transparent)
            $gfx.DrawImage($raw, 0, 0, $LogoH, $LogoH)
            $gfx.Dispose(); $raw.Dispose()
        } catch { $img = $null }
    }
    if (-not $img) { return 8 }

    $pb = New-Object System.Windows.Forms.PictureBox
    $pb.Image     = $img
    $pb.SizeMode  = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    $pb.Size      = [System.Drawing.Size]::new($LogoH, $LogoH)
    $pb.BackColor = $clrAccent
    $pb.Anchor    = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
    $pb.Location  = [System.Drawing.Point]::new(8, [int](($Header.Height - $LogoH) / 2))
    $Header.Controls.Add($pb)
    return ($LogoH + 16)   # x-offset caller should use for the title label
}

# ── WINFORMS HELPERS (used by Migration Runner) ───────────────────────────────
function New-CardPanel {
    param($Title = "")
    $outer = New-Object System.Windows.Forms.Panel
    $outer.BackColor = $clrBorder
    $outer.Dock      = [System.Windows.Forms.DockStyle]::Fill
    $outer.Padding   = New-Object System.Windows.Forms.Padding(1)
    $inner = New-Object System.Windows.Forms.Panel
    $inner.BackColor = $clrPanel
    $inner.Dock      = [System.Windows.Forms.DockStyle]::Fill
    $outer.Controls.Add($inner)
    $bar = New-Object System.Windows.Forms.Panel
    $bar.BackColor = $clrAccent
    $bar.Width     = 4
    $bar.Dock      = [System.Windows.Forms.DockStyle]::Left
    $inner.Controls.Add($bar)
    if ($Title) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text      = $Title
        $lbl.Font      = $FontCap
        $lbl.ForeColor = $clrMuted
        $lbl.Location  = [System.Drawing.Point]::new(16, 9)
        $lbl.AutoSize  = $true
        $inner.Controls.Add($lbl)
    }
    return $inner
}

function New-Lbl {
    param($Parent, [string]$Text, [int]$X, [int]$Y, [bool]$Bold = $false)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text; $l.Location = [System.Drawing.Point]::new($X, $Y)
    $l.AutoSize = $true; $l.Anchor = $AnchorTL
    $l.Font = if ($Bold) { $FontBold } else { $FontBody }
    $l.ForeColor = $clrText
    $Parent.Controls.Add($l)
    return $l
}

function New-TB {
    param($Parent, [int]$X, [int]$Y, [int]$W, [int]$RightMargin = -1,
          [bool]$Password = $false, [string]$Default = "")
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location    = [System.Drawing.Point]::new($X, $Y)
    $tb.Size        = [System.Drawing.Size]::new($W, 24)
    $tb.Font        = $FontBody
    $tb.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $tb.Text        = $Default
    if ($Password) { $tb.UseSystemPasswordChar = $true }
    if ($RightMargin -ge 0) {
        $tb.Anchor = $AnchorTLR
        $tb.Width  = $Parent.Width - $X - $RightMargin
    } else {
        $tb.Anchor = $AnchorTL
    }
    $Parent.Controls.Add($tb)
    return $tb
}

function New-Btn {
    param($Parent, [string]$Text, [int]$X, [int]$Y, [int]$W = 140, [int]$H = 30,
          [bool]$Primary = $true, [bool]$AnchorRight = $false)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text; $b.Location = [System.Drawing.Point]::new($X, $Y)
    $b.Size = [System.Drawing.Size]::new($W, $H); $b.Font = $FontBold
    $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $b.FlatAppearance.BorderSize = 0
    $b.BackColor = if ($Primary) { $clrAccent } else { [System.Drawing.Color]::FromArgb(225, 228, 238) }
    $b.ForeColor = if ($Primary) { [System.Drawing.Color]::White } else { $clrText }
    $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $b.Anchor    = if ($AnchorRight) { $AnchorTR } else { $AnchorTL }
    $Parent.Controls.Add($b)
    return $b
}

function New-Dot {
    param($Parent, [int]$X, [int]$Y)
    $d = New-Object System.Windows.Forms.Label
    $d.Text = [char]0x25CF; $d.Font = New-Object System.Drawing.Font("Segoe UI", 13)
    $d.ForeColor = $clrGrey; $d.Location = [System.Drawing.Point]::new($X, $Y)
    $d.AutoSize = $true; $d.Anchor = $AnchorTR
    $Parent.Controls.Add($d)
    return $d
}

function New-HSep {
    param($Parent, [int]$X, [int]$Y, [int]$RightMargin = 10)
    $s = New-Object System.Windows.Forms.Label
    $s.Location  = [System.Drawing.Point]::new($X, $Y)
    $s.Height    = 1
    $s.Width     = $Parent.Width - $X - $RightMargin
    $s.BackColor = $clrBorder
    $s.Anchor    = $AnchorTLR
    $Parent.Controls.Add($s)
}

# ── SHARED WRITE-LOG (App Reg overrides locally; Migration Runner uses this) ──
function Write-Log {
    param([string]$Msg, [string]$Level = "INFO")
    $ts = Get-Date -Format "HH:mm:ss"
    if ($script:rtbLog -and -not $script:rtbLog.IsDisposed) {
        $script:rtbLog.SelectionStart  = $script:rtbLog.TextLength
        $script:rtbLog.SelectionLength = 0
        $script:rtbLog.SelectionColor  = [System.Drawing.Color]::FromArgb(80, 95, 120)
        $script:rtbLog.AppendText("$ts ")
        $levelColor = switch ($Level) {
            "OK"    { [System.Drawing.Color]::FromArgb(65, 195, 110) }
            "WARN"  { [System.Drawing.Color]::FromArgb(220, 165, 45) }
            "ERROR" { [System.Drawing.Color]::FromArgb(225, 80, 80) }
            default { [System.Drawing.Color]::FromArgb(120, 155, 220) }
        }
        $script:rtbLog.SelectionColor = $levelColor
        $script:rtbLog.AppendText("[$Level] ")
        $script:rtbLog.SelectionColor = [System.Drawing.Color]::FromArgb(205, 212, 230)
        $script:rtbLog.AppendText("$Msg`n")
        $script:rtbLog.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
    if ($script:LogFile) { "$ts [$Level] $Msg" | Add-Content -Path $script:LogFile -Encoding UTF8 }
}

# ── WORKLOAD DEFINITIONS ───────────────────────────────────────────────────────
# Single authoritative source for AvePoint Fly cmdlet names per workload.
# runner.ps1 and monitor.ps1 both reference this via $script:FlyWorkloadDefs.
$script:FlyWorkloadDefs = [ordered]@{
    SharePoint   = @{ Import = 'Import-FlySharePointMappings';  Start = 'Start-FlySharePointMigration';  PreScan = 'Start-FlySharePointPreScan';  Verify = 'Start-FlySharePointVerification';  Status = 'Export-FlySharePointMappingStatus';  Report = 'Export-FlySharePointMigrationReport';  PolicyType = 'SharePoint' }
    Exchange     = @{ Import = 'Import-FlyExchangeMappings';    Start = 'Start-FlyExchangeMigration';    PreScan = 'Start-FlyExchangePreScan';    Verify = 'Start-FlyExchangeVerification';    Status = 'Export-FlyExchangeMappingStatus';    Report = 'Export-FlyExchangeMigrationReport';    PolicyType = 'Exchange'   }
    OneDrive     = @{ Import = 'Import-FlyOneDriveMappings';    Start = 'Start-FlyOneDriveMigration';    PreScan = 'Start-FlyOneDrivePreScan';    Verify = 'Start-FlyOneDriveVerification';    Status = 'Export-FlyOneDriveMappingStatus';    Report = 'Export-FlyOneDriveMigrationReport';    PolicyType = 'OneDrive'   }
    Teams        = @{ Import = 'Import-FlyTeamsMappings';       Start = 'Start-FlyTeamsMigration';       PreScan = 'Start-FlyTeamsPreScan';       Verify = 'Start-FlyTeamsVerification';       Status = 'Export-FlyTeamsMappingStatus';       Report = 'Export-FlyTeamsMigrationReport';       PolicyType = 'Teams'      }
    TeamChat     = @{ Import = 'Import-FlyTeamChatMappings';    Start = 'Start-FlyTeamChatMigration';    PreScan = '';                            Verify = 'Start-FlyTeamChatVerification';    Status = 'Export-FlyTeamChatMappingStatus';    Report = 'Export-FlyTeamChatMigrationReport';    PolicyType = 'TeamChat'   }
    Groups       = @{ Import = 'Import-FlyM365GroupMappings';   Start = 'Start-FlyM365GroupMigration';   PreScan = 'Start-FlyM365GroupPreScan';   Verify = 'Start-FlyM365GroupVerification';   Status = 'Export-FlyM365GroupMappingStatus';   Report = 'Export-FlyM365GroupMigrationReport';   PolicyType = 'M365Group'  }
}

# ── FLY CONNECTION LOOKUP & AUTO-CREATE ───────────────────────────────────────
# Workload codes → AOS portal labels used by fly-connector.js
$script:WorkloadToAosLabel = @{
    Exchange   = 'Exchange Online'
    SharePoint = 'SharePoint Online'
    OneDrive   = 'OneDrive'
    Teams      = 'Microsoft Teams'
    TeamChat   = 'Microsoft Teams Chat'
    Groups     = 'Microsoft 365 Groups'
}

# Creates Fly destination connections for all workloads by driving the AOS portal
# via fly-connector.js (Playwright). Requires Node.js + a saved AOS session.
function Invoke-FlyConnectorCreate {
    param(
        [string]$TenantName,    # customer prefix, e.g. "Fara"
        [string]$TenantSearch   # domain prefix, e.g. "mbufara"
    )

    $connectorJs = Join-Path $PSScriptRoot 'fly-connector.js'
    $authFile    = Join-Path $PSScriptRoot 'auth\storageState.json'

    if (-not (Test-Path $connectorJs)) {
        throw "fly-connector.js not found at '$PSScriptRoot'. Cannot auto-create connections."
    }
    if (-not (Test-Path $authFile)) {
        throw ("No AOS browser session found.`n" +
               "In the toolkit, go to Settings → Config and click 'Sign in to AOS', " +
               "complete the Microsoft SSO in the browser, then try again.")
    }

    $nodeExe = (Get-Command node.exe -ErrorAction SilentlyContinue)?.Source
    if (-not $nodeExe) {
        throw "Node.js not found on PATH. Install Node.js 18+ from nodejs.org to enable automatic connection creation."
    }

    Write-Host "`nCreating destination connections for '$TenantName' (all workloads)..." -ForegroundColor Cyan

    # Build one task per workload
    $tasks = $script:WorkloadToAosLabel.GetEnumerator() | ForEach-Object {
        @{
            id              = [guid]::NewGuid().ToString('N').Substring(0, 8)
            tenantName      = $TenantName
            tenantSearch    = $TenantSearch
            workloadLabel  = $_.Value
            connectionName = "$TenantName - $($_.Value)"
        } | ConvertTo-Json -Compress
    }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = $nodeExe
    $psi.WorkingDirectory       = $PSScriptRoot
    $psi.RedirectStandardInput  = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute        = $false
    $psi.CreateNoWindow         = $true
    $psi.ArgumentList.Add($connectorJs)
    $psi.ArgumentList.Add('--mode=create')
    $psi.ArgumentList.Add('--headless')
    $psi.ArgumentList.Add("--display-name=$TenantName")

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi
    $null = $proc.Start()

    foreach ($line in $tasks) { $proc.StandardInput.WriteLine($line) }
    $proc.StandardInput.Close()

    $created = 0; $failed = 0; $skipped = 0
    while (-not $proc.StandardOutput.EndOfStream) {
        $raw = $proc.StandardOutput.ReadLine()
        if (-not $raw) { continue }
        try {
            $obj = $raw | ConvertFrom-Json -ErrorAction Stop
            if ($obj.event -and $obj.message) {
                Write-Host "  [AOS] $($obj.message)" -ForegroundColor Gray
            } elseif ($obj.id -and $obj.status) {
                $c = switch ($obj.status) { 'CREATED'{'Green'} 'SKIPPED'{'Yellow'} 'FAILED'{'Red'} default{'Gray'} }
                Write-Host "  $($obj.status)  $($obj.message)" -ForegroundColor $c
                switch ($obj.status) { 'CREATED'{$created++} 'FAILED'{$failed++} 'SKIPPED'{$skipped++} }
            }
        } catch { Write-Host "  $raw" -ForegroundColor Gray }
    }
    $proc.WaitForExit(120000)

    if ($failed -gt 0) {
        Write-Warning "$failed connection(s) failed to create. Check the connector logs for details."
    } else {
        Write-Host "Connections ready: $created created, $skipped already existed" -ForegroundColor Green
    }
    return ($failed -eq 0)
}

# Finds the destination connection for a given project+workload.
# Returns the name of the migration policy to use for a given workload.
# Prefers the default policy; falls back to the first available policy.
# migrationModuleType integers match connectionType: Exchange=0, Teams=1, SharePoint=2, OneDrive=3, Groups=4, TeamChat=5
# Requires: Fly.Client module loaded and Connect-Fly already called.
function Find-FlyPolicy {
    param(
        [string]$Workload,  # e.g. "Exchange"
        [string]$ApiUrl
    )

    $typeMap = @{ Exchange=0; Teams=1; SharePoint=2; OneDrive=3; Groups=4; TeamChat=5 }
    $flyModule = Get-Module Fly.Client
    if (-not $flyModule) { throw "Fly.Client module is not loaded" }
    $cfg   = & $flyModule { $Script:Configuration }
    $token = $cfg['AccessToken']
    $base  = ($cfg['BaseUrl'] ?? $ApiUrl).TrimEnd('/')
    if (-not $token) { throw "No bearer token — Connect-Fly may have failed" }

    $moduleType = $typeMap[$Workload]
    Write-Host "Looking up $Workload policies..." -ForegroundColor Cyan

    $response = Invoke-RestMethod `
        -Uri     "$base/policies/summaries?migrationModuleType=$moduleType&top=100" `
        -Headers @{ Authorization="Bearer $token"; Accept="application/json" } `
        -Method  Get -ErrorAction Stop

    # Response may be double-nested { data: { data: [...] } } or single { data: [...] }
    $policies = if ($response.data.data) { $response.data.data }
                elseif ($response.data -is [array]) { $response.data }
                else { $null }

    if (-not $policies -or $policies.Count -eq 0) {
        throw "No $Workload policies found in Fly. Create at least one policy in the Fly portal before running migrations."
    }

    # Prefer the marked default; otherwise take the first one
    $policy = $policies | Where-Object { $_.isDefault -eq $true } | Select-Object -First 1
    if (-not $policy) { $policy = $policies | Select-Object -First 1 }

    Write-Host "Using policy: $($policy.name)" -ForegroundColor Green
    return $policy.name
}

# If no connections exist for the customer and CustomerDomain is provided,
# auto-creates them via fly-connector.js first.
# Requires: Fly.Client module loaded and Connect-Fly already called.
function Find-FlyDestinationConnection {
    param(
        [string]$ProjectName,         # e.g. "Fara - Exchange"
        [string]$Workload,            # e.g. "Exchange"
        [string]$ApiUrl,              # e.g. "https://graph.avepointonlineservices.com/fly"
        [string]$CustomerDomain = ""  # tenant search code, e.g. "mbufara" — triggers auto-create if set
    )

    # connectionType integers used by the Fly API
    $typeMap = @{ Exchange=0; Teams=1; SharePoint=2; OneDrive=3; Groups=4; TeamChat=5 }

    $flyModule = Get-Module Fly.Client
    if (-not $flyModule) { throw "Fly.Client module is not loaded" }
    $cfg   = & $flyModule { $Script:Configuration }
    $token = $cfg['AccessToken']
    $base  = ($cfg['BaseUrl'] ?? $ApiUrl).TrimEnd('/')
    if (-not $token) { throw "No bearer token — Connect-Fly may have failed" }

    $prefix      = ($ProjectName -split ' - ')[0]
    $prefixLower = $prefix.ToLower()
    $expectedType = $typeMap[$Workload]

    $headers = @{ Authorization="Bearer $token"; Accept="application/json" }

    function Get-ConnList {
        $r = Invoke-RestMethod -Uri "$base/connections/summaries?top=500" -Headers $headers -Method Get -ErrorAction Stop
        return $r.data.data
    }

    Write-Host "Looking up destination connection for '$prefix'..." -ForegroundColor Cyan

    $connList = Get-ConnList

    # Check if any connections exist for this customer prefix
    $existing = $connList | Where-Object { $_.name -and $_.name.ToLower() -like "$prefixLower*" }

    if (-not $existing -and $CustomerDomain) {
        Write-Host "No connections found for '$prefix' — creating via AOS portal..." -ForegroundColor Yellow
        $ok = Invoke-FlyConnectorCreate -TenantName $prefix -TenantSearch $CustomerDomain
        if ($ok) {
            Write-Host "Connections created — re-querying..." -ForegroundColor Cyan
            Start-Sleep -Seconds 3   # brief pause for Fly to index new connections
            $connList = Get-ConnList
        }
    }

    # Best match: name starts with prefix AND correct workload type
    $conn = $connList | Where-Object {
        $_.name -and $_.name.ToLower() -like "$prefixLower*" -and $_.connectionType -eq $expectedType
    } | Select-Object -First 1

    # Fallback: any connection whose name starts with the customer prefix
    if (-not $conn) {
        $conn = $connList | Where-Object {
            $_.name -and $_.name.ToLower() -like "$prefixLower*"
        } | Select-Object -First 1
    }

    if (-not $conn) {
        $available = ($connList | Where-Object { $_.name } |
                      Select-Object -ExpandProperty name | Sort-Object -Unique) -join "`n  "
        $hint = if ($CustomerDomain) {
            "The connector ran but the connection may not have been created successfully. Check the connector logs."
        } else {
            "Add the customer domain in Settings → Customer, then retry."
        }
        throw "No destination connection found for '$prefix' ($Workload).`n$hint`n`nAvailable connections:`n  $available"
    }

    Write-Host "Using connection: $($conn.name)" -ForegroundColor Green
    return $conn.name
}

# ── LOG MANAGEMENT ────────────────────────────────────────────────────────────
function Move-OldLogs {
    <#
    .SYNOPSIS
        Archives old log files to keep logs folder clean
    .PARAMETER DaysOld
        Logs older than this many days will be moved to 'old' subfolder (default: 7)
    #>
    param([int]$DaysOld = 7)

    $logPath = Join-Path $PSScriptRoot "logs"
    if (-not (Test-Path $logPath)) { return }

    # Create old folder if needed
    $oldFolder = Join-Path $logPath "old"
    if (-not (Test-Path $oldFolder)) {
        New-Item -ItemType Directory -Path $oldFolder -Force | Out-Null
    }

    # Move old logs
    $cutoffDate = (Get-Date).AddDays(-$DaysOld)
    Get-ChildItem -Path $logPath -File | Where-Object {
        $_.LastWriteTime -lt $cutoffDate
    } | ForEach-Object {
        $dest = Join-Path $oldFolder $_.Name
        if (Test-Path $dest) {
            $timestamp = $_.LastWriteTime.ToString('yyyyMMdd-HHmmss')
            $dest = Join-Path $oldFolder "$([System.IO.Path]::GetFileNameWithoutExtension($_.Name))-$timestamp$($_.Extension)"
        }
        Move-Item $_.FullName $dest -Force -ErrorAction SilentlyContinue
    }

    # Clean up very old logs (90+ days)
    Get-ChildItem -Path $oldFolder -File | Where-Object {
        $_.LastWriteTime -lt (Get-Date).AddDays(-90)
    } | Remove-Item -Force -ErrorAction SilentlyContinue
}
