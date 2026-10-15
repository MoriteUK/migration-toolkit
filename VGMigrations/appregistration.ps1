# ═════════════════════════════════════════════════════════════════════════════
# APP REGISTRATION
# ═════════════════════════════════════════════════════════════════════════════
function Show-AppRegistrationForm {

    function New-Card {
        param($Parent, [int]$X, [int]$Y, [int]$W, [int]$H, [string]$Title = "")
        $outer = New-Object System.Windows.Forms.Panel
        $outer.Location  = [System.Drawing.Point]::new($X, $Y)
        $outer.Size      = [System.Drawing.Size]::new($W, $H)
        $outer.BackColor = $clrBorder
        $Parent.Controls.Add($outer)
        $inner = New-Object System.Windows.Forms.Panel
        $inner.Location  = [System.Drawing.Point]::new(1, 1)
        $inner.Size      = [System.Drawing.Size]::new($W - 2, $H - 2)
        $inner.BackColor = $clrPanel
        $outer.Controls.Add($inner)
        $bar = New-Object System.Windows.Forms.Panel
        $bar.Location  = [System.Drawing.Point]::new(0, 0)
        $bar.Size      = [System.Drawing.Size]::new(4, $H - 2)
        $bar.BackColor = $clrAccent
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

    # New-Lbl, New-TB, New-Btn, New-Dot are provided by lib.ps1 (dot-sourced before this file).

    function Write-Log {
        param([string]$Msg, [string]$Level = "INFO")
        $ts = Get-Date -Format "HH:mm:ss"
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

    function Invoke-GraphApi {
        param([string]$Method = "GET", [string]$Endpoint, $Body = $null)
        $p = @{
            Method      = $Method
            Uri         = "https://graph.microsoft.com/v1.0$Endpoint"
            Headers     = @{ Authorization = "Bearer $($script:GraphToken)"; "Content-Type" = "application/json" }
            ErrorAction = "Stop"
        }
        if ($Body) { $p.Body = ($Body | ConvertTo-Json -Depth 10) }
        return Invoke-RestMethod @p
    }

    $Form = New-Object System.Windows.Forms.Form
    $Form.Text            = "AvePoint Fly - Target Tenant App Registration"
    $Form.ClientSize      = [System.Drawing.Size]::new(780, 644)
    $Form.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $Form.BackColor       = $clrBg
    $Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $Form.MaximizeBox     = $false
    $Form.Font            = $FontBody
    $_ico = Join-Path $PSScriptRoot 'FlyMigration.ico'; if (Test-Path $_ico) { $Form.Icon = [System.Drawing.Icon]::new($_ico) }

    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Size = [System.Drawing.Size]::new(780, 54); $hdr.BackColor = $clrAccent
    $Form.Controls.Add($hdr)
    $_hdrX = Add-HeaderLogo $hdr 34
    $hdrTitle = New-Object System.Windows.Forms.Label
    $hdrTitle.Text      = "  Target Tenant App Registration"
    $hdrTitle.Font      = New-Object System.Drawing.Font("Segoe UI Semibold", 12)
    $hdrTitle.ForeColor = [System.Drawing.Color]::White
    $hdrTitle.Location  = [System.Drawing.Point]::new($_hdrX, 15); $hdrTitle.AutoSize = $true
    $hdr.Controls.Add($hdrTitle)
    $btnGear = New-Object System.Windows.Forms.Button
    $btnGear.BackColor = $clrAccent
    $btnGear.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnGear.FlatAppearance.BorderSize = 0
    $btnGear.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(0, 78, 152)
    $btnGear.Size = [System.Drawing.Size]::new(38, 38); $btnGear.Location = [System.Drawing.Point]::new(734, 8)
    $btnGear.Cursor = [System.Windows.Forms.Cursors]::Hand; $btnGear.Add_Click({ Show-SettingsDialog })
    if ($script:GearBitmap) {
        $btnGear.Image = $script:GearBitmap; $btnGear.ImageAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    } else {
        $btnGear.Text = [char]0x2699; $btnGear.Font = New-Object System.Drawing.Font("Segoe UI", 16)
        $btnGear.ForeColor = [System.Drawing.Color]::White
    }
    $hdr.Controls.Add($btnGear)

    $footer = New-Object System.Windows.Forms.Panel
    $footer.Height    = 46
    $footer.Dock      = [System.Windows.Forms.DockStyle]::Bottom
    $footer.BackColor = [System.Drawing.Color]::FromArgb(28, 32, 48)
    $Form.Controls.Add($footer)

    $btnReg = New-Object System.Windows.Forms.Button
    $btnReg.Text      = "Authenticate and Register"
    $btnReg.Location  = [System.Drawing.Point]::new(16, 8)
    $btnReg.Size      = [System.Drawing.Size]::new(210, 30)
    $btnReg.BackColor = $clrAccent; $btnReg.ForeColor = [System.Drawing.Color]::White
    $btnReg.Font      = $FontBold; $btnReg.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnReg.FlatAppearance.BorderSize = 0; $btnReg.Cursor = [System.Windows.Forms.Cursors]::Hand
    $footer.Controls.Add($btnReg)

    $dotReg = New-Object System.Windows.Forms.Panel
    $dotReg.Size      = [System.Drawing.Size]::new(12, 12)
    $dotReg.Location  = [System.Drawing.Point]::new(238, 17)
    $dotReg.BackColor = $clrGrey
    $footer.Controls.Add($dotReg)

    $lblResult = New-Object System.Windows.Forms.Label
    $lblResult.Location  = [System.Drawing.Point]::new(260, 14)
    $lblResult.Size      = [System.Drawing.Size]::new(400, 20)
    $lblResult.ForeColor = [System.Drawing.Color]::FromArgb(190, 210, 255)
    $lblResult.Font      = $FontBody; $lblResult.Text = "Waiting..."
    $footer.Controls.Add($lblResult)

    $btnCloseForm = New-Object System.Windows.Forms.Button
    $btnCloseForm.Text      = "Close"
    $btnCloseForm.Size      = [System.Drawing.Size]::new(90, 30)
    $btnCloseForm.Location  = [System.Drawing.Point]::new(662, 8)
    $btnCloseForm.BackColor = [System.Drawing.Color]::FromArgb(200, 55, 55)
    $btnCloseForm.ForeColor = [System.Drawing.Color]::White; $btnCloseForm.Font = $FontBold
    $btnCloseForm.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnCloseForm.FlatAppearance.BorderSize = 0
    $btnCloseForm.Cursor    = [System.Windows.Forms.Cursors]::Hand; $btnCloseForm.Add_Click({ $Form.Close() })
    $footer.Controls.Add($btnCloseForm)
    $footer.Add_SizeChanged({ $btnCloseForm.Left = $footer.Width - 100 })

    $s1 = New-Card -Parent $Form -X 14 -Y 64 -W 752 -H 100 -Title "STEP 1  -  TARGET TENANT"
    New-Lbl $s1 "App Name"      16  33 | Out-Null
    New-Lbl $s1 "Tenant Domain" 310 33 | Out-Null
    $tbAppName = New-TB $s1 16  51 278 -Default "AvePoint Fly Migration"
    $tbTenant  = New-TB $s1 310 51 260 -Default "contoso.onmicrosoft.com"

    $rdoNew = New-Object System.Windows.Forms.RadioButton
    $rdoNew.Text = "Register new app"; $rdoNew.Font = $FontBody
    $rdoNew.Location = [System.Drawing.Point]::new(590, 48); $rdoNew.AutoSize = $true; $rdoNew.Checked = $true
    $s1.Controls.Add($rdoNew)

    $rdoExist = New-Object System.Windows.Forms.RadioButton
    $rdoExist.Text = "Use existing app"; $rdoExist.Font = $FontBody
    $rdoExist.Location = [System.Drawing.Point]::new(590, 70); $rdoExist.AutoSize = $true
    $s1.Controls.Add($rdoExist)

    $s2 = New-Card -Parent $Form -X 14 -Y 174 -W 752 -H 80 -Title "STEP 2  -  EXISTING APP CREDENTIALS (if using existing)"
    New-Lbl $s2 "Client ID"     16 33 | Out-Null
    New-Lbl $s2 "Client Secret" 370 33 | Out-Null
    $tbClientId     = New-TB $s2 16  51 340
    $tbClientSecret = New-TB $s2 370 51 270 -Password $true
    $tbClientId.Enabled = $false; $tbClientSecret.Enabled = $false

    $rdoExist.Add_CheckedChanged({
        $tbClientId.Enabled     = $rdoExist.Checked
        $tbClientSecret.Enabled = $rdoExist.Checked
    })

    $s4 = New-Card -Parent $Form -X 14 -Y 264 -W 752 -H 100 -Title "OUTPUT  -  Copy these values into Fly connections"
    New-Lbl $s4 "Tenant ID"       16 33 | Out-Null
    New-Lbl $s4 "App (Client) ID" 16 58 | Out-Null
    $tbOutTenantId = New-TB $s4 130 30 478 -Default ""
    $tbOutAppId    = New-TB $s4 130 55 330 -Default ""
    $tbOutTenantId.ReadOnly = $true; $tbOutTenantId.BackColor = [System.Drawing.Color]::FromArgb(245,247,252)
    $tbOutAppId.ReadOnly    = $true; $tbOutAppId.BackColor    = [System.Drawing.Color]::FromArgb(245,247,252)
    $btnCopyTenantId = New-Btn $s4 "Copy"        618  26  96 28 $false
    $btnCopyAppId    = New-Btn $s4 "Copy"        468  52  96 28 $false
    $btnCopySecret   = New-Btn $s4 "Copy Secret" 572  52 112 28 $false
    $btnCopyTenantId.Enabled = $false
    $btnCopyAppId.Enabled    = $false
    $btnCopySecret.Enabled   = $false

    $s5 = New-Card -Parent $Form -X 14 -Y 374 -W 752 -H 214 -Title "LOG"
    $script:rtbLog = New-Object System.Windows.Forms.RichTextBox
    $script:rtbLog.Location    = [System.Drawing.Point]::new(16, 26)
    $script:rtbLog.Size        = [System.Drawing.Size]::new(716, 178)
    $script:rtbLog.Font        = $FontMono; $script:rtbLog.BackColor = $clrLogBg
    $script:rtbLog.ForeColor   = [System.Drawing.Color]::FromArgb(190, 210, 255)
    $script:rtbLog.ReadOnly    = $true; $script:rtbLog.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $script:rtbLog.ScrollBars  = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    $s5.Controls.Add($script:rtbLog)

    $script:PlainSecret = $null
    $btnCopySecret.Add_Click({
        if ($script:PlainSecret) {
            [System.Windows.Forms.Clipboard]::SetText($script:PlainSecret)
            Write-Log "Client secret copied to clipboard." "OK"
        }
    })
    $btnCopyTenantId.Add_Click({
        if ($tbOutTenantId.Text) {
            [System.Windows.Forms.Clipboard]::SetText($tbOutTenantId.Text)
            Write-Log "Tenant ID copied to clipboard." "OK"
        }
    })
    $btnCopyAppId.Add_Click({
        if ($tbOutAppId.Text) {
            [System.Windows.Forms.Clipboard]::SetText($tbOutAppId.Text)
            Write-Log "App (Client) ID copied to clipboard." "OK"
        }
    })

    $btnReg.Add_Click({
        $btnReg.Enabled        = $false
        $dotReg.BackColor      = $clrGrey
        $lblResult.Text        = "Working..."
        $lblResult.ForeColor   = $clrMuted
        $btnCopySecret.Enabled = $false

        try {
            $tenantDomain = $tbTenant.Text.Trim()
            $appName      = $tbAppName.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($tenantDomain)) { throw "Tenant domain is required." }
            if ([string]::IsNullOrWhiteSpace($appName))      { throw "App name is required." }

            Write-Log "Resolving tenant ID for $tenantDomain..."
            $oidc = Invoke-RestMethod `
                -Uri "https://login.microsoftonline.com/$tenantDomain/.well-known/openid-configuration" `
                -ErrorAction Stop
            $tenantId = ($oidc.issuer -split "/")[3]
            Write-Log "Tenant ID: $tenantId"

            if ($rdoExist.Checked) {
                $appId     = $tbClientId.Text.Trim()
                $appSecret = $tbClientSecret.Text.Trim()
                if ([string]::IsNullOrWhiteSpace($appId) -or [string]::IsNullOrWhiteSpace($appSecret)) {
                    throw "Client ID and Secret are required for existing app."
                }
                Write-Log "Using existing app: $appId" "OK"
                $script:PlainSecret = $appSecret
                Write-Log "Existing app - verify API permissions and Exchange Administrator role are configured in Entra ID." "WARN"
            }
            else {
                $publicClientId = "1950a258-227b-4e31-a9cf-717495945fc2"
                $escapedScope   = [Uri]::EscapeDataString("https://graph.microsoft.com/.default")

                Write-Log "Starting device code flow..."
                $dcBody = "client_id=$publicClientId" + "&scope=$escapedScope"
                $dcResp = Invoke-RestMethod -Method POST `
                    -Uri "https://login.microsoftonline.com/$tenantDomain/oauth2/v2.0/devicecode" `
                    -Body $dcBody -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop

                Write-Log "URL:  $($dcResp.verification_uri)" "WARN"
                Write-Log "Code: $($dcResp.user_code)" "WARN"

                [System.Windows.Forms.MessageBox]::Show(
                    "Sign in as Global Admin in the TARGET tenant:`n`n" +
                    "URL:   $($dcResp.verification_uri)`n" +
                    "Code:  $($dcResp.user_code)`n`n" +
                    "Click OK after completing sign-in.",
                    "Target Tenant Authentication",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null

                Write-Log "Polling for token..."
                $tokenUrl  = "https://login.microsoftonline.com/$tenantDomain/oauth2/v2.0/token"
                $grantType = [Uri]::EscapeDataString("urn:ietf:params:oauth:grant-type:device_code")
                $script:GraphToken = $null
                $attempts = 0
                while ((-not $script:GraphToken) -and $attempts -lt 40) {
                    Start-Sleep -Seconds 3; $attempts++
                    try {
                        $tokenBody = "client_id=$publicClientId" + "&grant_type=$grantType" + "&device_code=$($dcResp.device_code)"
                        $tr = Invoke-RestMethod -Method POST -Uri $tokenUrl `
                            -Body $tokenBody -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
                        $script:GraphToken = $tr.access_token
                    } catch { }
                }
                if (-not $script:GraphToken) { throw "Authentication timed out." }
                Write-Log "Authenticated OK" "OK"

                Write-Log "Registering app: $appName..."
                $newApp = Invoke-GraphApi -Method POST -Endpoint "/applications" -Body @{
                    displayName    = $appName
                    signInAudience = "AzureADMyOrg"
                }
                $appId = $newApp.appId

                Write-Log "Creating service principal..."
                $newSP = Invoke-GraphApi -Method POST -Endpoint "/servicePrincipals" -Body @{ appId = $appId }

                Write-Log "Creating client secret (1 year)..."
                $secretResp = Invoke-GraphApi -Method POST `
                    -Endpoint "/applications/$($newApp.id)/addPassword" -Body @{
                        passwordCredential = @{
                            displayName = "FlyMigration"
                            endDateTime = (Get-Date).AddYears(1).ToString("yyyy-MM-ddTHH:mm:ssZ")
                        }
                    }
                $script:PlainSecret = $secretResp.secretText
                Write-Log "Secret expires: $($secretResp.endDateTime)" "OK"
                Update-SharedConfig @{ SecretExpiry = $secretResp.endDateTime }

                Write-Log "Resolving resource service principals..."
                $graphResSP = $null
                $exoResSP   = $null

                Write-Log "Looking up Microsoft Graph service principal..." "INFO"
                try {
                    $gResult = Invoke-GraphApi -Endpoint "/servicePrincipals?`$filter=appId eq '00000003-0000-0000-c000-000000000000'&`$select=id,appRoles"
                    $graphResSP = $gResult.value | Select-Object -First 1
                    if ($graphResSP) {
                        Write-Log "Microsoft Graph SP found (id: $($graphResSP.id))" "INFO"
                    } else {
                        Write-Log "Microsoft Graph SP not in results — activating..." "INFO"
                        Invoke-GraphApi -Method POST -Endpoint "/servicePrincipals" -Body @{ appId = '00000003-0000-0000-c000-000000000000' } | Out-Null
                        Start-Sleep -Seconds 3
                        $gResult2 = Invoke-GraphApi -Endpoint "/servicePrincipals?`$filter=appId eq '00000003-0000-0000-c000-000000000000'&`$select=id,appRoles"
                        $graphResSP = $gResult2.value | Select-Object -First 1
                        if ($graphResSP) { Write-Log "Microsoft Graph SP activated OK" "INFO" }
                        else { Write-Log "Microsoft Graph SP still not found after activation" "WARN" }
                    }
                } catch {
                    Write-Log "Microsoft Graph SP error: $($_.Exception.Message)" "WARN"
                }

                Write-Log "Looking up Exchange Online service principal..." "INFO"
                try {
                    $eResult = Invoke-GraphApi -Endpoint "/servicePrincipals?`$filter=appId eq '00000002-0000-0ff1-ce00-000000000000'&`$select=id,appRoles"
                    $exoResSP = $eResult.value | Select-Object -First 1
                    if ($exoResSP) {
                        Write-Log "Exchange Online SP found (id: $($exoResSP.id))" "INFO"
                    } else {
                        Write-Log "Exchange Online SP not in results — activating..." "INFO"
                        Invoke-GraphApi -Method POST -Endpoint "/servicePrincipals" -Body @{ appId = '00000002-0000-0ff1-ce00-000000000000' } | Out-Null
                        Start-Sleep -Seconds 3
                        $eResult2 = Invoke-GraphApi -Endpoint "/servicePrincipals?`$filter=appId eq '00000002-0000-0ff1-ce00-000000000000'&`$select=id,appRoles"
                        $exoResSP = $eResult2.value | Select-Object -First 1
                        if ($exoResSP) { Write-Log "Exchange Online SP activated OK" "INFO" }
                        else { Write-Log "Exchange Online SP still not found after activation" "WARN" }
                    }
                } catch {
                    Write-Log "Exchange Online SP error: $($_.Exception.Message)" "WARN"
                }

                if ($graphResSP -and $exoResSP) {
                    $graphPermNames = @(
                        'Directory.ReadWrite.All', 'User.ReadWrite.All',
                        'Group.ReadWrite.All', 'GroupMember.ReadWrite.All',
                        'Sites.FullControl.All', 'Files.ReadWrite.All',
                        'Mail.ReadWrite', 'MailboxSettings.ReadWrite',
                        'Calendars.ReadWrite', 'Contacts.ReadWrite', 'Notes.ReadWrite.All',
                        'TeamSettings.ReadWrite.All', 'Channel.ReadWrite.All',
                        'ChannelMember.ReadWrite.All', 'Chat.ReadWrite.All'
                    )
                    $exoPermNames = @('Exchange.ManageAsApp')

                    $graphRoles = $graphPermNames | ForEach-Object {
                        $n = $_
                        $r = $graphResSP.appRoles | Where-Object value -eq $n | Select-Object -First 1
                        if (-not $r) { Write-Log "Graph permission '$n' not found - skipping" "WARN" }
                        $r
                    } | Where-Object { $_ }

                    $exoRoles = $exoPermNames | ForEach-Object {
                        $n = $_
                        $r = $exoResSP.appRoles | Where-Object value -eq $n | Select-Object -First 1
                        if (-not $r) { Write-Log "EXO permission '$n' not found - skipping" "WARN" }
                        $r
                    } | Where-Object { $_ }

                    Write-Log "Adding $($graphRoles.Count + $exoRoles.Count) API permissions to app registration..."
                    Invoke-GraphApi -Method PATCH -Endpoint "/applications/$($newApp.id)" -Body @{
                        requiredResourceAccess = @(
                            @{
                                resourceAppId  = '00000003-0000-0000-c000-000000000000'
                                resourceAccess = @($graphRoles | ForEach-Object { @{ id = $_.id; type = 'Role' } })
                            }
                            @{
                                resourceAppId  = '00000002-0000-0ff1-ce00-000000000000'
                                resourceAccess = @($exoRoles | ForEach-Object { @{ id = $_.id; type = 'Role' } })
                            }
                        )
                    } | Out-Null
                    Write-Log "Permissions added to app registration." "OK"

                    Write-Log "Granting admin consent..."
                    $granted = 0; $skipped = 0
                    foreach ($role in $graphRoles) {
                        try {
                            Invoke-GraphApi -Method POST -Endpoint "/servicePrincipals/$($newSP.id)/appRoleAssignments" -Body @{
                                principalId = $newSP.id; resourceId = $graphResSP.id; appRoleId = $role.id
                            } | Out-Null
                            $granted++
                        } catch {
                            if ($_.Exception.Message -like "*already exists*") { $skipped++ }
                            else { Write-Log "Consent warning '$($role.value)': $($_.Exception.Message)" "WARN" }
                        }
                    }
                    foreach ($role in $exoRoles) {
                        try {
                            Invoke-GraphApi -Method POST -Endpoint "/servicePrincipals/$($newSP.id)/appRoleAssignments" -Body @{
                                principalId = $newSP.id; resourceId = $exoResSP.id; appRoleId = $role.id
                            } | Out-Null
                            $granted++
                        } catch {
                            if ($_.Exception.Message -like "*already exists*") { $skipped++ }
                            else { Write-Log "Consent warning '$($role.value)': $($_.Exception.Message)" "WARN" }
                        }
                    }
                    Write-Log "Admin consent: $granted granted, $skipped already existed." "OK"
                } else {
                    Write-Log "Could not resolve Graph/EXO service principals - skipping consent grant." "WARN"
                    Write-Log "Grant admin consent manually in Entra ID > App registrations > API permissions." "WARN"
                }

                Write-Log "Assigning Exchange Administrator role to service principal..."
                $exAdminRole = $null
                try {
                    $exAdminRole = (Invoke-GraphApi -Endpoint "/directoryRoles?`$filter=displayName eq 'Exchange Administrator'").value | Select-Object -First 1
                } catch { }
                if (-not $exAdminRole) {
                    $tmpl = (Invoke-GraphApi -Endpoint "/directoryRoleTemplates?`$filter=displayName eq 'Exchange Administrator'").value | Select-Object -First 1
                    if ($tmpl) { $exAdminRole = Invoke-GraphApi -Method POST -Endpoint "/directoryRoles" -Body @{ roleTemplateId = $tmpl.id } }
                }
                if ($exAdminRole) {
                    try {
                        Invoke-GraphApi -Method POST -Endpoint "/directoryRoles/$($exAdminRole.id)/members/`$ref" -Body @{
                            '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($newSP.id)"
                        } | Out-Null
                        Write-Log "Exchange Administrator role assigned." "OK"
                    } catch {
                        if ($_.Exception.Message -like "*already exists*" -or
                            $_.Exception.Message -like "*One or more added object references already exist*") {
                            Write-Log "Exchange Administrator role already assigned." "OK"
                        } else {
                            Write-Log "Exchange Administrator role assignment: $($_.Exception.Message)" "WARN"
                        }
                    }
                } else {
                    Write-Log "Exchange Administrator role not found in directory." "WARN"
                }
            }

            $tbOutTenantId.Text    = $tenantId
            $tbOutAppId.Text       = $appId
            Update-SharedConfig @{ AppName = $appName; TenantDomain = $tenantDomain; TenantId = $tenantId; AppId = $appId }
            $btnCopyTenantId.Enabled = $true
            $btnCopyAppId.Enabled    = $true
            $btnCopySecret.Enabled   = $true
            $dotReg.BackColor        = $clrGreen
            $lblResult.Text        = "App ID: $appId"
            $lblResult.ForeColor   = $clrGreen
            Write-Log "Complete" "OK"
        }
        catch {
            $dotReg.BackColor    = $clrRed
            $lblResult.Text      = "Failed - see log"
            $lblResult.ForeColor = $clrRed
            $btnReg.Enabled      = $true
            Write-Log "Failed: $($_.Exception.Message)" "ERROR"
        }
    })

    $_sc = Read-SharedConfig
    if ($_sc.TenantDomain -and $tbTenant.Text -eq "contoso.onmicrosoft.com") { $tbTenant.Text  = $_sc.TenantDomain }
    if ($_sc.AppName      -and $tbAppName.Text -eq "AvePoint Fly Migration")  { $tbAppName.Text = $_sc.AppName }
    Write-Log "Enter target tenant details and click Authenticate and Register."
    [void]$Form.ShowDialog()
}
