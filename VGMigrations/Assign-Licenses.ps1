#Requires -Version 7.0
<#
.SYNOPSIS
    Assigns Microsoft 365 licenses to users from an exported CSV file.

.DESCRIPTION
    GUI tool that reads a UserLicenses CSV file (from Export-UserLicenses.ps1 or Run-Assessment.ps1)
    and assigns those licenses back to users in the new tenant after migration.

    Supports:
    - Dry-run mode to preview changes without applying them
    - Progress tracking with detailed logging
    - Error handling and reporting
    - SKU mapping for different tenant SKU IDs

.NOTES
    Run this AFTER users have been created/migrated to the new tenant.
    Requires Microsoft.Graph module and appropriate permissions.
#>

[CmdletBinding()]
param()

# Load WinForms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ── Logging setup ─────────────────────────────────────────────────────────────
$script:LogDir = Join-Path $env:APPDATA 'FlyMigration\Logs'
if (-not (Test-Path $script:LogDir)) { New-Item -ItemType Directory -Path $script:LogDir -Force | Out-Null }
$script:LogFile = Join-Path $script:LogDir "Assign-Licenses-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

function Write-Log {
    param([string]$Msg, [string]$Level = 'INFO')
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff')  [$($Level.PadRight(5))]  $Msg"
    Add-Content -Path $script:LogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
}

Write-Log "=== Assign-Licenses.ps1 started ==="

# ── Design System (matching lib.ps1) ─────────────────────────────────────────
$script:clrBg          = [System.Drawing.Color]::FromArgb(248, 249, 250)
$script:clrAccent      = [System.Drawing.Color]::FromArgb(38, 114, 236)
$script:clrAccentHover = [System.Drawing.Color]::FromArgb(30, 95, 200)
$script:clrSuccess     = [System.Drawing.Color]::FromArgb(0, 130, 70)
$script:clrWarning     = [System.Drawing.Color]::FromArgb(255, 153, 0)
$script:clrDanger      = [System.Drawing.Color]::FromArgb(220, 53, 69)
$script:clrText        = [System.Drawing.Color]::FromArgb(33, 37, 41)
$script:clrMuted       = [System.Drawing.Color]::FromArgb(108, 117, 125)
$script:clrBorder      = [System.Drawing.Color]::FromArgb(222, 226, 230)

$script:FontTitle = New-Object System.Drawing.Font('Segoe UI Light', 18)
$script:FontBody  = New-Object System.Drawing.Font('Segoe UI', 10)
$script:FontBold  = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
$script:FontSmall = New-Object System.Drawing.Font('Segoe UI', 8.5)

# ── License SKU mapping (skuPartNumber → skuId lookup) ───────────────────────
$script:TenantSkuMap = @{}  # Populated when connecting to Graph

# ── Main Form ─────────────────────────────────────────────────────────────────
function Show-LicenseAssignmentUI {
    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'Assign Licenses - Migration Toolkit'
    $form.ClientSize      = [System.Drawing.Size]::new(800, 650)
    $form.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.BackColor       = $script:clrBg
    $form.Font            = $script:FontBody
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.MaximizeBox     = $false

    # Icon
    $icoPath = Join-Path $PSScriptRoot 'FlyMigration.ico'
    if (Test-Path $icoPath) { $form.Icon = [System.Drawing.Icon]::new($icoPath) }

    # ── Header ────────────────────────────────────────────────────────────────
    $header = New-Object System.Windows.Forms.Panel
    $header.Size      = [System.Drawing.Size]::new(800, 60)
    $header.Dock      = [System.Windows.Forms.DockStyle]::Top
    $header.BackColor = $script:clrAccent
    $form.Controls.Add($header)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text      = '  Assign Licenses from CSV'
    $lblTitle.Font      = $script:FontTitle
    $lblTitle.ForeColor = [System.Drawing.Color]::White
    $lblTitle.Location  = [System.Drawing.Point]::new(16, 0)
    $lblTitle.Size      = [System.Drawing.Size]::new(760, 60)
    $lblTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $header.Controls.Add($lblTitle)

    $y = 80

    # ── CSV File Selection ────────────────────────────────────────────────────
    $lblCsv = New-Object System.Windows.Forms.Label
    $lblCsv.Text     = 'License CSV File (UserLicenses_*.csv):'
    $lblCsv.Font     = $script:FontBold
    $lblCsv.Location = [System.Drawing.Point]::new(20, $y)
    $lblCsv.AutoSize = $true
    $form.Controls.Add($lblCsv)
    $y += 26

    $txtCsvPath = New-Object System.Windows.Forms.TextBox
    $txtCsvPath.Location = [System.Drawing.Point]::new(20, $y)
    $txtCsvPath.Size     = [System.Drawing.Size]::new(650, 24)
    $txtCsvPath.Font     = $script:FontBody
    $txtCsvPath.ReadOnly = $true
    $form.Controls.Add($txtCsvPath)

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text       = 'Browse...'
    $btnBrowse.Location   = [System.Drawing.Point]::new(680, $y - 2)
    $btnBrowse.Size       = [System.Drawing.Size]::new(100, 28)
    $btnBrowse.BackColor  = $script:clrAccent
    $btnBrowse.ForeColor  = [System.Drawing.Color]::White
    $btnBrowse.FlatStyle  = [System.Windows.Forms.FlatStyle]::Flat
    $btnBrowse.Font       = $script:FontBold
    $btnBrowse.Cursor     = [System.Windows.Forms.Cursors]::Hand
    $btnBrowse.Add_Click({
        $openFile = New-Object System.Windows.Forms.OpenFileDialog
        $openFile.Filter = 'CSV Files (*.csv)|*.csv|All Files (*.*)|*.*'
        $openFile.Title  = 'Select UserLicenses CSV File'
        if ($openFile.ShowDialog() -eq 'OK') {
            $txtCsvPath.Text = $openFile.FileName
            Write-Log "CSV file selected: $($openFile.FileName)"
        }
    }.GetNewClosure())
    $form.Controls.Add($btnBrowse)
    $y += 38

    # ── Domain Mapping ────────────────────────────────────────────────────────
    $lblDomain = New-Object System.Windows.Forms.Label
    $lblDomain.Text     = 'Domain Mapping (optional - update UPNs for new tenant):'
    $lblDomain.Font     = $script:FontBold
    $lblDomain.Location = [System.Drawing.Point]::new(20, $y)
    $lblDomain.AutoSize = $true
    $form.Controls.Add($lblDomain)
    $y += 26

    $lblOldDomain = New-Object System.Windows.Forms.Label
    $lblOldDomain.Text     = 'Old Domain:'
    $lblOldDomain.Location = [System.Drawing.Point]::new(40, $y)
    $lblOldDomain.Size     = [System.Drawing.Size]::new(100, 20)
    $form.Controls.Add($lblOldDomain)

    $txtOldDomain = New-Object System.Windows.Forms.TextBox
    $txtOldDomain.Location    = [System.Drawing.Point]::new(140, $y - 2)
    $txtOldDomain.Size        = [System.Drawing.Size]::new(200, 24)
    $txtOldDomain.Font        = $script:FontBody
    $txtOldDomain.PlaceholderText = 'e.g., transloc.com'
    $form.Controls.Add($txtOldDomain)

    $lblArrow = New-Object System.Windows.Forms.Label
    $lblArrow.Text     = '→'
    $lblArrow.Location = [System.Drawing.Point]::new(350, $y)
    $lblArrow.Size     = [System.Drawing.Size]::new(20, 20)
    $lblArrow.Font     = New-Object System.Drawing.Font('Segoe UI', 12)
    $form.Controls.Add($lblArrow)

    $lblNewDomain = New-Object System.Windows.Forms.Label
    $lblNewDomain.Text     = 'New Domain:'
    $lblNewDomain.Location = [System.Drawing.Point]::new(380, $y)
    $lblNewDomain.Size     = [System.Drawing.Size]::new(100, 20)
    $form.Controls.Add($lblNewDomain)

    $txtNewDomain = New-Object System.Windows.Forms.TextBox
    $txtNewDomain.Location    = [System.Drawing.Point]::new(480, $y - 2)
    $txtNewDomain.Size        = [System.Drawing.Size]::new(290, 24)
    $txtNewDomain.Font        = $script:FontBody
    $txtNewDomain.PlaceholderText = 'e.g., transloc.onmicrosoft.com'
    $form.Controls.Add($txtNewDomain)
    $y += 38

    # ── Options ───────────────────────────────────────────────────────────────
    $chkDryRun = New-Object System.Windows.Forms.CheckBox
    $chkDryRun.Text     = 'Dry Run (preview only - do not assign licenses)'
    $chkDryRun.Location = [System.Drawing.Point]::new(20, $y)
    $chkDryRun.Size     = [System.Drawing.Size]::new(500, 24)
    $chkDryRun.Font     = $script:FontBody
    $chkDryRun.Checked  = $true
    $form.Controls.Add($chkDryRun)
    $y += 30

    $chkSkipUnlicensed = New-Object System.Windows.Forms.CheckBox
    $chkSkipUnlicensed.Text     = 'Skip users with no licenses in CSV'
    $chkSkipUnlicensed.Location = [System.Drawing.Point]::new(20, $y)
    $chkSkipUnlicensed.Size     = [System.Drawing.Size]::new(500, 24)
    $chkSkipUnlicensed.Font     = $script:FontBody
    $chkSkipUnlicensed.Checked  = $true
    $form.Controls.Add($chkSkipUnlicensed)
    $y += 40

    # ── Progress ──────────────────────────────────────────────────────────────
    $lblProgress = New-Object System.Windows.Forms.Label
    $lblProgress.Text     = 'Ready'
    $lblProgress.Font     = $script:FontBold
    $lblProgress.Location = [System.Drawing.Point]::new(20, $y)
    $lblProgress.Size     = [System.Drawing.Size]::new(760, 20)
    $form.Controls.Add($lblProgress)
    $y += 26

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = [System.Drawing.Point]::new(20, $y)
    $progressBar.Size     = [System.Drawing.Size]::new(760, 24)
    $progressBar.Style    = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $form.Controls.Add($progressBar)
    $y += 34

    # ── Log Output ────────────────────────────────────────────────────────────
    $lblLog = New-Object System.Windows.Forms.Label
    $lblLog.Text     = 'Log:'
    $lblLog.Font     = $script:FontBold
    $lblLog.Location = [System.Drawing.Point]::new(20, $y)
    $lblLog.AutoSize = $true
    $form.Controls.Add($lblLog)
    $y += 22

    $txtLog = New-Object System.Windows.Forms.TextBox
    $txtLog.Location   = [System.Drawing.Point]::new(20, $y)
    $txtLog.Size       = [System.Drawing.Size]::new(760, 260)
    $txtLog.Multiline  = $true
    $txtLog.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $txtLog.Font       = New-Object System.Drawing.Font('Consolas', 9)
    $txtLog.ReadOnly   = $true
    $txtLog.BackColor  = [System.Drawing.Color]::White
    $form.Controls.Add($txtLog)
    $y += 270

    # ── Action Buttons ────────────────────────────────────────────────────────
    $btnConnect = New-Object System.Windows.Forms.Button
    $btnConnect.Text      = 'Connect to Graph'
    $btnConnect.Location  = [System.Drawing.Point]::new(20, $y)
    $btnConnect.Size      = [System.Drawing.Size]::new(180, 36)
    $btnConnect.BackColor = $script:clrAccent
    $btnConnect.ForeColor = [System.Drawing.Color]::White
    $btnConnect.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnConnect.Font      = $script:FontBold
    $btnConnect.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $btnConnect.Add_Click({
        Connect-ToGraph -LogTextBox $txtLog -ProgressLabel $lblProgress
    }.GetNewClosure())
    $form.Controls.Add($btnConnect)

    $btnStart = New-Object System.Windows.Forms.Button
    $btnStart.Text      = 'Start Assignment'
    $btnStart.Location  = [System.Drawing.Point]::new(210, $y)
    $btnStart.Size      = [System.Drawing.Size]::new(180, 36)
    $btnStart.BackColor = $script:clrSuccess
    $btnStart.ForeColor = [System.Drawing.Color]::White
    $btnStart.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnStart.Font      = $script:FontBold
    $btnStart.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $btnStart.Enabled   = $false
    $btnStart.Add_Click({
        Start-LicenseAssignment `
            -CsvPath $txtCsvPath.Text `
            -OldDomain $txtOldDomain.Text `
            -NewDomain $txtNewDomain.Text `
            -DryRun $chkDryRun.Checked `
            -SkipUnlicensed $chkSkipUnlicensed.Checked `
            -LogTextBox $txtLog `
            -ProgressBar $progressBar `
            -ProgressLabel $lblProgress `
            -StartButton $btnStart
    }.GetNewClosure())
    $form.Controls.Add($btnStart)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text      = 'Close'
    $btnClose.Location  = [System.Drawing.Point]::new(690, $y)
    $btnClose.Size      = [System.Drawing.Size]::new(90, 36)
    $btnClose.BackColor = $script:clrDanger
    $btnClose.ForeColor = [System.Drawing.Color]::White
    $btnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnClose.Font      = $script:FontBold
    $btnClose.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $btnClose.Add_Click({ $form.Close() }.GetNewClosure())
    $form.Controls.Add($btnClose)

    # ── Show Form ─────────────────────────────────────────────────────────────
    $form.Add_Shown({
        Log-ToUI -Message 'Welcome to License Assignment Tool' -LogTextBox $txtLog
        Log-ToUI -Message "Log file: $script:LogFile" -LogTextBox $txtLog
        Log-ToUI -Message '' -LogTextBox $txtLog
        Log-ToUI -Message 'Step 1: Connect to Microsoft Graph' -LogTextBox $txtLog
        Log-ToUI -Message 'Step 2: Select the UserLicenses CSV file' -LogTextBox $txtLog
        Log-ToUI -Message 'Step 3: Click "Start Assignment"' -LogTextBox $txtLog
        Log-ToUI -Message '' -LogTextBox $txtLog

        # Enable Start button when both CSV is selected and Graph is connected
        $script:GraphConnected = $false
        $txtCsvPath.Add_TextChanged({
            $btnStart.Enabled = $script:GraphConnected -and (-not [string]::IsNullOrWhiteSpace($txtCsvPath.Text))
        }.GetNewClosure())
    }.GetNewClosure())

    [void]$form.ShowDialog()
}

# ── Helper Functions ──────────────────────────────────────────────────────────
function Log-ToUI {
    param(
        [string]$Message,
        [System.Windows.Forms.TextBox]$LogTextBox
    )

    $LogTextBox.AppendText("$Message`r`n")
    $LogTextBox.SelectionStart = $LogTextBox.Text.Length
    $LogTextBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Connect-ToGraph {
    param(
        [System.Windows.Forms.TextBox]$LogTextBox,
        [System.Windows.Forms.Label]$ProgressLabel
    )

    $ProgressLabel.Text = 'Connecting to Microsoft Graph...'
    Log-ToUI -Message '--- Connecting to Microsoft Graph ---' -LogTextBox $LogTextBox
    Write-Log 'Attempting Graph connection'

    # Check if Microsoft.Graph module is available
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Log-ToUI -Message '[ERROR] Microsoft.Graph module not found!' -LogTextBox $LogTextBox
        Log-ToUI -Message '        Install with: Install-Module Microsoft.Graph -Scope CurrentUser' -LogTextBox $LogTextBox
        $ProgressLabel.Text = 'Error: Module not found'
        Write-Log 'Microsoft.Graph module not found' 'ERROR'
        [System.Windows.Forms.MessageBox]::Show(
            "Microsoft.Graph module not found.`n`nInstall it with:`nInstall-Module Microsoft.Graph -Scope CurrentUser",
            'Module Not Found',
            'OK',
            'Error'
        ) | Out-Null
        return
    }

    try {
        Import-Module Microsoft.Graph.Authentication -DisableNameChecking -ErrorAction Stop

        $scopes = @('User.ReadWrite.All', 'Organization.Read.All')
        Log-ToUI -Message "Scopes: $($scopes -join ', ')" -LogTextBox $LogTextBox

        Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop

        Log-ToUI -Message '[OK] Connected to Microsoft Graph' -LogTextBox $LogTextBox
        Write-Log 'Graph connection successful'

        # Fetch tenant SKUs for license mapping
        Log-ToUI -Message 'Fetching tenant license SKUs...' -LogTextBox $LogTextBox
        $skuResponse = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/subscribedSkus' -Method GET -ErrorAction Stop
        $skus = $skuResponse.value

        $script:TenantSkuMap = @{}
        foreach ($sku in $skus) {
            if ($sku.skuPartNumber) {
                $script:TenantSkuMap[$sku.skuPartNumber] = $sku.skuId
            }
        }

        Log-ToUI -Message "[OK] Found $($script:TenantSkuMap.Count) license SKUs in tenant" -LogTextBox $LogTextBox
        Write-Log "Tenant SKUs loaded: $($script:TenantSkuMap.Count)"

        foreach ($skuPart in ($script:TenantSkuMap.Keys | Sort-Object)) {
            Log-ToUI -Message "  - $skuPart" -LogTextBox $LogTextBox
        }

        $script:GraphConnected = $true
        $ProgressLabel.Text = 'Connected to Graph - Ready to assign licenses'
        Log-ToUI -Message '' -LogTextBox $LogTextBox

    }
    catch {
        Log-ToUI -Message "[ERROR] Connection failed: $($_.Exception.Message)" -LogTextBox $LogTextBox
        $ProgressLabel.Text = 'Connection failed'
        Write-Log "Graph connection failed: $($_.Exception.Message)" 'ERROR'
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to connect to Microsoft Graph:`n`n$($_.Exception.Message)",
            'Connection Error',
            'OK',
            'Error'
        ) | Out-Null
    }
}

function Start-LicenseAssignment {
    param(
        [string]$CsvPath,
        [string]$OldDomain,
        [string]$NewDomain,
        [bool]$DryRun,
        [bool]$SkipUnlicensed,
        [System.Windows.Forms.TextBox]$LogTextBox,
        [System.Windows.Forms.ProgressBar]$ProgressBar,
        [System.Windows.Forms.Label]$ProgressLabel,
        [System.Windows.Forms.Button]$StartButton
    )

    if ([string]::IsNullOrWhiteSpace($CsvPath) -or -not (Test-Path $CsvPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            'Please select a valid UserLicenses CSV file.',
            'Invalid File',
            'OK',
            'Warning'
        ) | Out-Null
        return
    }

    $StartButton.Enabled = $false
    $ProgressLabel.Text = 'Reading CSV file...'
    Log-ToUI -Message '' -LogTextBox $LogTextBox
    Log-ToUI -Message '=== Starting License Assignment ===' -LogTextBox $LogTextBox
    Log-ToUI -Message "CSV File: $CsvPath" -LogTextBox $LogTextBox
    Log-ToUI -Message "Dry Run: $DryRun" -LogTextBox $LogTextBox
    Log-ToUI -Message "Skip Unlicensed: $SkipUnlicensed" -LogTextBox $LogTextBox

    # Domain mapping
    $useDomainMapping = (-not [string]::IsNullOrWhiteSpace($OldDomain)) -and (-not [string]::IsNullOrWhiteSpace($NewDomain))
    if ($useDomainMapping) {
        Log-ToUI -Message "Domain Mapping: @$OldDomain → @$NewDomain" -LogTextBox $LogTextBox
        Write-Log "Domain mapping enabled: $OldDomain -> $NewDomain"
    }

    Log-ToUI -Message '' -LogTextBox $LogTextBox
    Write-Log "Starting assignment from $CsvPath (DryRun=$DryRun)"

    try {
        # Read CSV
        $users = Import-Csv -Path $CsvPath -ErrorAction Stop
        Log-ToUI -Message "Loaded $($users.Count) users from CSV" -LogTextBox $LogTextBox
        Write-Log "CSV loaded: $($users.Count) users"

        # Filter out unlicensed if requested
        if ($SkipUnlicensed) {
            $users = $users | Where-Object {
                $_.LicenseCount -and [int]$_.LicenseCount -gt 0
            }
            Log-ToUI -Message "After filtering: $($users.Count) users with licenses" -LogTextBox $LogTextBox
            Write-Log "Filtered to $($users.Count) licensed users"
        }

        if ($users.Count -eq 0) {
            Log-ToUI -Message '[WARN] No users to process' -LogTextBox $LogTextBox
            $ProgressLabel.Text = 'No users to process'
            $StartButton.Enabled = $true
            return
        }

        $ProgressBar.Maximum = $users.Count
        $ProgressBar.Value = 0

        $stats = @{
            Total = $users.Count
            Success = 0
            Skipped = 0
            Failed = 0
            UnknownSkus = [System.Collections.Generic.HashSet[string]]::new()
        }

        # Process each user
        $current = 0
        foreach ($user in $users) {
            $current++
            $upn = $user.UserPrincipalName

            # Apply domain mapping if configured
            if ($useDomainMapping) {
                $upn = $upn -replace "@$([regex]::Escape($OldDomain))$", "@$NewDomain"
            }

            $ProgressBar.Value = $current
            $ProgressLabel.Text = "Processing $current of $($users.Count): $upn"
            [System.Windows.Forms.Application]::DoEvents()

            $displayUpn = if ($useDomainMapping -and $upn -ne $user.UserPrincipalName) {
                "$($user.UserPrincipalName) → $upn"
            } else {
                $upn
            }

            Log-ToUI -Message "[$current/$($users.Count)] $displayUpn" -LogTextBox $LogTextBox
            Write-Log "Processing user $current/$($users.Count): $upn"

            # Parse SKUs from CSV
            $skuParts = @($user.SkuPartNumbers -split '[;|]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })

            if ($skuParts.Count -eq 0) {
                Log-ToUI -Message "  [SKIP] No licenses" -LogTextBox $LogTextBox
                Write-Log "  Skipped: no licenses for $upn" 'WARN'
                $stats.Skipped++
                continue
            }

            Log-ToUI -Message "  Licenses: $($skuParts -join ', ')" -LogTextBox $LogTextBox

            # Map SKU part numbers to SKU IDs
            $licensesToAdd = [System.Collections.Generic.List[object]]::new()
            foreach ($skuPart in $skuParts) {
                if ($script:TenantSkuMap.ContainsKey($skuPart)) {
                    $skuId = $script:TenantSkuMap[$skuPart]
                    $licensesToAdd.Add([PSCustomObject]@{ SkuId = $skuId })
                    Write-Log "    Mapped $skuPart -> $skuId"
                }
                else {
                    Log-ToUI -Message "  [WARN] SKU not found in tenant: $skuPart" -LogTextBox $LogTextBox
                    Write-Log "    Unknown SKU: $skuPart" 'WARN'
                    $stats.UnknownSkus.Add($skuPart) | Out-Null
                }
            }

            if ($licensesToAdd.Count -eq 0) {
                Log-ToUI -Message "  [SKIP] No matching SKUs in tenant" -LogTextBox $LogTextBox
                Write-Log "  Skipped: no matching SKUs for $upn" 'WARN'
                $stats.Skipped++
                continue
            }

            # Assign licenses (or simulate in dry-run)
            if ($DryRun) {
                Log-ToUI -Message "  [DRY-RUN] Would assign $($licensesToAdd.Count) license(s)" -LogTextBox $LogTextBox
                Write-Log "  Dry-run: would assign $($licensesToAdd.Count) licenses to $upn"
                $stats.Success++
            }
            else {
                try {
                    # Build license assignment request
                    $body = @{
                        addLicenses    = @($licensesToAdd)
                        removeLicenses = @()
                    }

                    $upnEnc = [Uri]::EscapeDataString($upn)
                    Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/users/$upnEnc/assignLicense" `
                        -Method POST -Body $body -ErrorAction Stop | Out-Null

                    Log-ToUI -Message "  [OK] Assigned $($licensesToAdd.Count) license(s)" -LogTextBox $LogTextBox
                    Write-Log "  Success: assigned $($licensesToAdd.Count) licenses to $upn"
                    $stats.Success++
                }
                catch {
                    $errorMsg = $_.Exception.Message
                    Log-ToUI -Message "  [ERROR] $errorMsg" -LogTextBox $LogTextBox
                    Write-Log "  Failed to assign licenses to $upn: $errorMsg" 'ERROR'
                    $stats.Failed++
                }
            }
        }

        # Summary
        Log-ToUI -Message '' -LogTextBox $LogTextBox
        Log-ToUI -Message '=== Assignment Complete ===' -LogTextBox $LogTextBox
        Log-ToUI -Message "Total Users:  $($stats.Total)" -LogTextBox $LogTextBox
        Log-ToUI -Message "Successful:   $($stats.Success)" -LogTextBox $LogTextBox
        Log-ToUI -Message "Skipped:      $($stats.Skipped)" -LogTextBox $LogTextBox
        Log-ToUI -Message "Failed:       $($stats.Failed)" -LogTextBox $LogTextBox

        if ($stats.UnknownSkus.Count -gt 0) {
            Log-ToUI -Message '' -LogTextBox $LogTextBox
            Log-ToUI -Message 'Unknown SKUs (not available in target tenant):' -LogTextBox $LogTextBox
            foreach ($sku in $stats.UnknownSkus) {
                Log-ToUI -Message "  - $sku" -LogTextBox $LogTextBox
            }
        }

        $ProgressLabel.Text = if ($DryRun) { 'Dry run complete' } else { 'Assignment complete' }
        Write-Log "Assignment complete: Success=$($stats.Success) Skipped=$($stats.Skipped) Failed=$($stats.Failed)"

        [System.Windows.Forms.MessageBox]::Show(
            "License assignment complete!`n`nTotal: $($stats.Total)`nSuccessful: $($stats.Success)`nSkipped: $($stats.Skipped)`nFailed: $($stats.Failed)",
            'Complete',
            'OK',
            'Information'
        ) | Out-Null
    }
    catch {
        Log-ToUI -Message "[ERROR] $($_.Exception.Message)" -LogTextBox $LogTextBox
        $ProgressLabel.Text = 'Error occurred'
        Write-Log "Assignment failed: $($_.Exception.Message)" 'ERROR'
        [System.Windows.Forms.MessageBox]::Show(
            "An error occurred:`n`n$($_.Exception.Message)",
            'Error',
            'OK',
            'Error'
        ) | Out-Null
    }
    finally {
        $StartButton.Enabled = $true
    }
}

# ── Entry Point ───────────────────────────────────────────────────────────────
try {
    Show-LicenseAssignmentUI
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" 'ERROR'
    [System.Windows.Forms.MessageBox]::Show(
        "Fatal error:`n`n$($_.Exception.Message)`n`nLog: $script:LogFile",
        'Fatal Error',
        'OK',
        'Error'
    ) | Out-Null
}

Write-Log "=== Assign-Licenses.ps1 exiting ==="
