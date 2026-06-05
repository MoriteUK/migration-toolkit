#Requires -Version 7.0
# AvePoint Fly Migration Toolkit - Entry Point

. "$PSScriptRoot\lib.ps1"
. "$PSScriptRoot\settings.ps1"
. "$PSScriptRoot\appregistration.ps1"
. "$PSScriptRoot\aossetup.ps1"
. "$PSScriptRoot\runner.ps1"
. "$PSScriptRoot\monitor.ps1"
. "$PSScriptRoot\reports.ps1"

# Archive old logs on startup
Move-OldLogs -DaysOld 7

function Show-MainMenu {
    $MenuForm = New-Object System.Windows.Forms.Form
    $MenuForm.Text            = "AvePoint Fly - Dashboard"
    $MenuForm.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $MenuForm.ClientSize      = [System.Drawing.Size]::new(1000, 700)
    $MenuForm.BackColor       = $clrBg
    $MenuForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $MenuForm.MaximizeBox     = $true
    $MenuForm.WindowState     = [System.Windows.Forms.FormWindowState]::Maximized
    $MenuForm.Font            = $FontBody
    $_ico = Join-Path $PSScriptRoot 'FlyMigration.ico'; if (Test-Path $_ico) { $MenuForm.Icon = [System.Drawing.Icon]::new($_ico) }

    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Size = [System.Drawing.Size]::new(1000, 80); $hdr.Dock = [System.Windows.Forms.DockStyle]::Top; $hdr.BackColor = $clrAccent
    $MenuForm.Controls.Add($hdr)
    $_hdrX = Add-HeaderLogo $hdr 32
    $hdrTitle = New-Object System.Windows.Forms.Label
    $hdrTitle.Text      = "  AvePoint Fly - Dashboard"
    $hdrTitle.Font      = New-Object System.Drawing.Font('Segoe UI Light', 22)
    $hdrTitle.ForeColor = [System.Drawing.Color]::White
    $hdrTitle.Location  = [System.Drawing.Point]::new($_hdrX, 0)
    $hdrTitle.Size      = [System.Drawing.Size]::new(800, 80)
    $hdrTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $hdr.Controls.Add($hdrTitle)

    $btnGear = New-Object System.Windows.Forms.Button
    $btnGear.BackColor = $clrAccent
    $btnGear.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnGear.FlatAppearance.BorderSize = 0
    $btnGear.FlatAppearance.MouseOverBackColor = $clrAccentHover
    $btnGear.Size     = [System.Drawing.Size]::new(42, 42)
    $btnGear.Location = [System.Drawing.Point]::new(940, 19)
    $btnGear.Cursor   = [System.Windows.Forms.Cursors]::Hand
    $btnGear.Add_Click({ Show-SettingsDialog })
    if ($script:GearBitmap) {
        $btnGear.Image = $script:GearBitmap; $btnGear.ImageAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    } else {
        $btnGear.Text = [char]0x2699; $btnGear.Font = New-Object System.Drawing.Font("Segoe UI Symbol", 20)
        $btnGear.ForeColor = [System.Drawing.Color]::White
    }
    $hdr.Controls.Add($btnGear)

    # Stat card helper function for dashboard metrics
    function MkStatCard { param([int]$X,[int]$Y,[int]$W,[int]$H,[string]$Label,[string]$Value,[string]$Status)
        $card = New-Object System.Windows.Forms.Panel
        $card.Location = [System.Drawing.Point]::new($X,$Y)
        $card.Size = [System.Drawing.Size]::new($W,$H)
        $card.BackColor = [System.Drawing.Color]::White
        $card.BorderStyle = [System.Windows.Forms.BorderStyle]::None

        # Create rounded region
        $radius = 8
        $regionPath = New-Object System.Drawing.Drawing2D.GraphicsPath
        $regionRect = [System.Drawing.Rectangle]::new(0, 0, $W, $H)
        $regionPath.AddArc($regionRect.X, $regionRect.Y, $radius * 2, $radius * 2, 180, 90)
        $regionPath.AddArc($regionRect.Right - $radius * 2, $regionRect.Y, $radius * 2, $radius * 2, 270, 90)
        $regionPath.AddArc($regionRect.Right - $radius * 2, $regionRect.Bottom - $radius * 2, $radius * 2, $radius * 2, 0, 90)
        $regionPath.AddArc($regionRect.X, $regionRect.Bottom - $radius * 2, $radius * 2, $radius * 2, 90, 90)
        $regionPath.CloseFigure()
        $card.Region = New-Object System.Drawing.Region($regionPath)

        # Paint handler for border
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

        # Value (big number)
        $lblValue = New-Object System.Windows.Forms.Label
        $lblValue.Text = $Value
        $lblValue.Location = [System.Drawing.Point]::new(16, 16)
        $lblValue.AutoSize = $true
        $lblValue.Font = New-Object System.Drawing.Font('Segoe UI Light', 28)
        $lblValue.ForeColor = $clrText
        $card.Controls.Add($lblValue)

        # Label (description)
        $lblLabel = New-Object System.Windows.Forms.Label
        $lblLabel.Text = $Label
        $lblLabel.Location = [System.Drawing.Point]::new(16, 60)
        $lblLabel.Size = [System.Drawing.Size]::new($W - 32, 20)
        $lblLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
        $lblLabel.ForeColor = $clrMuted
        $card.Controls.Add($lblLabel)

        # Status badge
        if ($Status) {
            $lblStatus = New-Object System.Windows.Forms.Label
            $lblStatus.Text = $Status
            $lblStatus.Location = [System.Drawing.Point]::new(16, 84)
            $lblStatus.AutoSize = $true
            $lblStatus.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 8)
            $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
            $card.Controls.Add($lblStatus)
        }

        $MenuForm.Controls.Add($card)
        return $card
    }

    # Action card helper function with rounded corners and blue top edge
    function MkCard { param([int]$X,[int]$Y,[int]$W,[int]$H,[string]$Title,[string]$Subtitle)
        $card = New-Object System.Windows.Forms.Panel
        $card.Location = [System.Drawing.Point]::new($X,$Y)
        $card.Size = [System.Drawing.Size]::new($W,$H)
        $card.BackColor = [System.Drawing.Color]::White
        $card.BorderStyle = [System.Windows.Forms.BorderStyle]::None
        $card.Cursor = [System.Windows.Forms.Cursors]::Hand

        # Store original position for hover effect
        $card | Add-Member -NotePropertyName OriginalY -NotePropertyValue $Y

        # Create rounded region to clip the panel
        $radius = 12
        $regionPath = New-Object System.Drawing.Drawing2D.GraphicsPath
        $regionRect = [System.Drawing.Rectangle]::new(0, 0, $W, $H)
        $regionPath.AddArc($regionRect.X, $regionRect.Y, $radius * 2, $radius * 2, 180, 90)
        $regionPath.AddArc($regionRect.Right - $radius * 2, $regionRect.Y, $radius * 2, $radius * 2, 270, 90)
        $regionPath.AddArc($regionRect.Right - $radius * 2, $regionRect.Bottom - $radius * 2, $radius * 2, $radius * 2, 0, 90)
        $regionPath.AddArc($regionRect.X, $regionRect.Bottom - $radius * 2, $radius * 2, $radius * 2, 90, 90)
        $regionPath.CloseFigure()
        $card.Region = New-Object System.Drawing.Region($regionPath)

        # Add rounded corners with blue top edge
        $card.Add_Paint({
            param($s, $e)
            $g = $e.Graphics
            $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias

            # Create rounded rectangle path
            $radius = 12
            $rect = [System.Drawing.Rectangle]::new(0, 0, $s.Width - 1, $s.Height - 1)
            $path = New-Object System.Drawing.Drawing2D.GraphicsPath

            $path.AddArc($rect.X, $rect.Y, $radius * 2, $radius * 2, 180, 90)
            $path.AddArc($rect.Right - $radius * 2, $rect.Y, $radius * 2, $radius * 2, 270, 90)
            $path.AddArc($rect.Right - $radius * 2, $rect.Bottom - $radius * 2, $radius * 2, $radius * 2, 0, 90)
            $path.AddArc($rect.X, $rect.Bottom - $radius * 2, $radius * 2, $radius * 2, 90, 90)
            $path.CloseFigure()

            # Fill background
            $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
            $g.FillPath($brush, $path)

            # Draw border
            $pen = New-Object System.Drawing.Pen($clrBorder, 1)
            $g.DrawPath($pen, $path)

            # Draw blue top edge (gradient bar)
            $topPath = New-Object System.Drawing.Drawing2D.GraphicsPath
            $topPath.AddArc($rect.X, $rect.Y, $radius * 2, $radius * 2, 180, 90)
            $topPath.AddLine($rect.X + $radius, $rect.Y, $rect.Right - $radius, $rect.Y)
            $topPath.AddArc($rect.Right - $radius * 2, $rect.Y, $radius * 2, $radius * 2, 270, 90)
            $topPath.AddLine($rect.Right, $rect.Y + $radius, $rect.Right, $rect.Y + 4)
            $topPath.AddLine($rect.Right, $rect.Y + 4, $rect.X, $rect.Y + 4)
            $topPath.CloseFigure()

            $gradientBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
                [System.Drawing.Point]::new(0, 0),
                [System.Drawing.Point]::new($s.Width, 0),
                $clrAccent,
                [System.Drawing.Color]::FromArgb(0, 82, 163)
            )
            $g.FillPath($gradientBrush, $topPath)

            $brush.Dispose()
            $pen.Dispose()
            $path.Dispose()
            $topPath.Dispose()
            $gradientBrush.Dispose()
        }.GetNewClosure())

        # Title label
        $lblTitle = New-Object System.Windows.Forms.Label
        $lblTitle.Text = $Title
        $lblTitle.Location = [System.Drawing.Point]::new(16, 20)
        $lblTitle.AutoSize = $true
        $lblTitle.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 13)
        $lblTitle.ForeColor = $clrText
        $card.Controls.Add($lblTitle)

        # Subtitle label
        $lblSub = New-Object System.Windows.Forms.Label
        $lblSub.Text = $Subtitle
        $lblSub.Location = [System.Drawing.Point]::new(16, 46)
        $lblSub.Size = [System.Drawing.Size]::new($W - 32, 50)
        $lblSub.Font = $FontSub
        $lblSub.ForeColor = $clrMuted
        $card.Controls.Add($lblSub)

        # Hover effect - move up on mouse enter
        $card.Add_MouseEnter({
            param($s, $e)
            $s.Top = $s.OriginalY - 4
            $s.Invalidate()
        }.GetNewClosure())

        # Hover effect - move back down on mouse leave
        $card.Add_MouseLeave({
            param($s, $e)
            $s.Top = $s.OriginalY
            $s.Invalidate()
        }.GetNewClosure())

        $MenuForm.Controls.Add($card)
        return $card
    }

    # Dashboard layout
    $margin = 32; $gap = 20; $y = 110

    # Quick stats row - 4 stat cards
    $statW = 220; $statH = 110
    $stat1 = MkStatCard $margin $y $statW $statH 'Active Projects' '3' 'Running'
    $stat2 = MkStatCard ($margin + $statW + $gap) $y $statW $statH 'Users Migrated' '1,247' 'This month'
    $stat3 = MkStatCard ($margin + ($statW + $gap) * 2) $y $statW $statH 'Success Rate' '94%' 'Last 30 days'
    $stat4 = MkStatCard ($margin + ($statW + $gap) * 3) $y $statW $statH 'Total Data' '2.4 TB' 'Transferred'
    $y += $statH + $gap + 10

    # Section header
    $lblActions = New-Object System.Windows.Forms.Label
    $lblActions.Text = 'Quick Actions'
    $lblActions.Location = [System.Drawing.Point]::new($margin, $y)
    $lblActions.AutoSize = $true
    $lblActions.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 12)
    $lblActions.ForeColor = $clrText
    $MenuForm.Controls.Add($lblActions)
    $y += 35

    # Action cards - 3 wide
    $cardW = 305; $cardH = 110
    $card1 = MkCard $margin $y $cardW $cardH 'App Registration' 'Register Entra ID app and grant API permissions'
    $card2 = MkCard ($margin + $cardW + $gap) $y $cardW $cardH 'AOS Setup' 'Configure AvePoint Online Services tenant'
    $card3 = MkCard ($margin + ($cardW + $gap) * 2) $y $cardW $cardH 'Connections' 'Manage connections and mappings'
    $y += $cardH + $gap

    # Row 2
    $card4 = MkCard $margin $y $cardW $cardH 'Reports' 'View migration results and status'
    $card5 = MkCard ($margin + $cardW + $gap) $y $cardW $cardH 'Monitor' 'Live project monitoring and tracking'
    $card6 = MkCard ($margin + ($cardW + $gap) * 2) $y $cardW $cardH 'Documentation' 'View guides and best practices'

    $footer = New-Object System.Windows.Forms.Panel
    $footer.Height = 56; $footer.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $footer.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)
    $MenuForm.Controls.Add($footer)

    # Version label in footer
    $lblVersion = New-Object System.Windows.Forms.Label
    $lblVersion.Text = "Version $($script:ToolVersion)"
    $lblVersion.Location = [System.Drawing.Point]::new(32, 18)
    $lblVersion.AutoSize = $true
    $lblVersion.Font = $FontSub
    $lblVersion.ForeColor = $clrMuted
    $footer.Controls.Add($lblVersion)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'; $btnClose.Size = [System.Drawing.Size]::new(100, 36)
    $btnClose.Location = [System.Drawing.Point]::new(880, 10)
    $btnClose.BackColor = $clrCloseRed
    $btnClose.ForeColor = [System.Drawing.Color]::White; $btnClose.Font = $FontBold
    $btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnClose.FlatAppearance.BorderSize = 0
    $btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
    $footer.Controls.Add($btnClose)
    $footer.Add_SizeChanged({ $btnClose.Left = $footer.Width - 106 }.GetNewClosure())

    $card1.Add_Click({ Show-AppRegistrationForm })
    $card2.Add_Click({ Show-AosSetupForm })
    $card3.Add_Click({ Show-MigrationRunnerForm })
    $card4.Add_Click({ Show-ReportingForm })
    $card5.Add_Click({
        if ($script:MonitorFormInstance -and
            -not $script:MonitorFormInstance.IsDisposed -and
            $script:MonitorFormInstance.Visible) {
            $script:MonitorFormInstance.BringToFront()
            $script:MonitorFormInstance.Focus()
        } else {
            Show-ProjectMonitorForm
        }
    })
    $card6.Add_Click({
        Start-Process 'https://github.com/MoriteUK/AvepointFlyUtility/wiki'
    })
    $btnClose.Add_Click({ $MenuForm.Close() }.GetNewClosure())

    [System.Windows.Forms.Application]::Run($MenuForm)
}

# ═════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═════════════════════════════════════════════════════════════════════════════
Show-MainMenu
