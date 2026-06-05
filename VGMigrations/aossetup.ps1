# ═════════════════════════════════════════════════════════════════════════════
# AOS TENANT & APP SETUP
# ═════════════════════════════════════════════════════════════════════════════
function Show-AosSetupForm {

    $script:aosSetupConnectorJs = Join-Path $PSScriptRoot 'fly-connector.js'

    [void][System.Reflection.Assembly]::LoadWithPartialName('PresentationFramework')
    [void][System.Reflection.Assembly]::LoadWithPartialName('PresentationCore')

    [xml]$setupXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="AOS Tenant &amp; App Setup"
        Width="820" Height="640" MinWidth="640" MinHeight="500"
        WindowStartupLocation="CenterScreen"
        Background="#F0F2F8">
    <DockPanel>
        <Border DockPanel.Dock="Top" Height="54" Background="#0064B4">
            <DockPanel Margin="10,0,16,0" VerticalAlignment="Center">
                <Image Name="ImgLogo" DockPanel.Dock="Left" Height="34" Width="34" Margin="0,0,8,0"
                       RenderOptions.BitmapScalingMode="HighQuality"/>
                <TextBlock FontFamily="Segoe UI" FontSize="15" Foreground="White" VerticalAlignment="Center">
                    <Run Text="AOS Tenant &amp; App Setup" FontWeight="Light"/>
                </TextBlock>
            </DockPanel>
        </Border>

        <Border DockPanel.Dock="Bottom" Background="#1C2030" Height="50" Padding="16,8">
            <DockPanel LastChildFill="True">
                <Button Name="BtnSignIn" DockPanel.Dock="Left" Content="Sign in to AOS..."
                        Background="#0064B4" Foreground="White" FontWeight="SemiBold"
                        Width="160" Margin="0,0,12,0" BorderThickness="0"/>
                <Button Name="BtnClose" DockPanel.Dock="Right" Content="Close"
                        Background="#C83737" Foreground="White" FontWeight="SemiBold"
                        Width="90" Margin="0" BorderThickness="0"/>
                <TextBlock Name="TxtAuthStatus" VerticalAlignment="Center" Foreground="#BED2FF"
                           Text="Click to sign in — opens Chrome, complete Microsoft SSO, then close it."/>
            </DockPanel>
        </Border>

        <Grid Margin="16,10,16,14">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <!-- Tenant details -->
            <GroupBox Grid.Row="0" Header="Tenant" Margin="0,0,0,6">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="150"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="20"/>
                        <ColumnDefinition Width="150"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Label  Grid.Column="0" Content="Display Name:" VerticalAlignment="Center"/>
                    <TextBox Grid.Column="1" Name="TxtDisplayName" Margin="0,4,0,4"
                             ToolTip="Shown in AOS, e.g. OurVolaris"/>
                    <Label  Grid.Column="3" Content="Search Code:" VerticalAlignment="Center"/>
                    <TextBox Grid.Column="4" Name="TxtSearchCode" Margin="0,4,0,4"
                             ToolTip="Short code used in the AOS Tenant dropdown, e.g. ourvolaris"/>
                </Grid>
            </GroupBox>

            <!-- App profile -->
            <GroupBox Grid.Row="1" Header="App Profile (credentials registered in AOS)" Margin="0,0,0,6">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="150"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="20"/>
                        <ColumnDefinition Width="150"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition/>
                        <RowDefinition/>
                    </Grid.RowDefinitions>
                    <Label  Grid.Row="0" Grid.Column="0" Content="Profile Name:" VerticalAlignment="Center"/>
                    <TextBox Grid.Row="0" Grid.Column="1" Name="TxtAppProfileName" Margin="0,4,0,2"
                             ToolTip="Name for the app profile in AOS App Management, e.g. OurVolaris App"/>
                    <Label  Grid.Row="0" Grid.Column="3" Content="Client ID:" VerticalAlignment="Center"/>
                    <TextBox Grid.Row="0" Grid.Column="4" Name="TxtClientId" Margin="0,4,0,2"
                             ToolTip="App (Client) ID from Entra / the Create App Registration screen"/>
                    <Label  Grid.Row="1" Grid.Column="3" Content="Client Secret:" VerticalAlignment="Center"/>
                    <PasswordBox Grid.Row="1" Grid.Column="4" Name="TxtClientSecret" Margin="0,2,0,4"
                                 ToolTip="Copy from the Create App Registration screen — Copy Secret button"/>
                </Grid>
            </GroupBox>

            <!-- Buttons -->
            <Grid Grid.Row="2" Margin="0,2,0,8">
                <StackPanel HorizontalAlignment="Right" Orientation="Horizontal">
                    <Button Name="BtnClear"    Content="Clear Results"  Margin="0,0,6,0"/>
                    <Button Name="BtnStop"     Content="Stop"           IsEnabled="False" Margin="0,0,6,0"/>
                    <Button Name="BtnRunSetup" Content="Run Setup"
                            Background="#0064B4" Foreground="White" FontWeight="SemiBold" Width="110"/>
                </StackPanel>
            </Grid>

            <!-- Results grid -->
            <DataGrid Grid.Row="3" Name="DgSetup" AutoGenerateColumns="False"
                      IsReadOnly="True" HeadersVisibility="Column"
                      GridLinesVisibility="Horizontal"
                      AlternatingRowBackground="#F5F6FA"
                      Background="White" BorderBrush="#D2D7E4"
                      RowHeight="28" FontSize="12">
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Time"    Binding="{Binding Timestamp}" Width="80"/>
                    <DataGridTextColumn Header="Tenant"  Binding="{Binding Tenant}"    Width="200"/>
                    <DataGridTextColumn Header="Status"  Binding="{Binding Status}"    Width="100">
                        <DataGridTextColumn.CellStyle>
                            <Style TargetType="DataGridCell">
                                <Setter Property="FontWeight" Value="SemiBold"/>
                                <Setter Property="Padding"    Value="6,0"/>
                                <Style.Triggers>
                                    <DataTrigger Binding="{Binding Status}" Value="DONE">   <Setter Property="Foreground" Value="#107C10"/></DataTrigger>
                                    <DataTrigger Binding="{Binding Status}" Value="FAILED"> <Setter Property="Foreground" Value="#D13438"/></DataTrigger>
                                    <DataTrigger Binding="{Binding Status}" Value="SKIPPED"><Setter Property="Foreground" Value="#808080"/></DataTrigger>
                                    <DataTrigger Binding="{Binding Status}" Value="WORKING"><Setter Property="Foreground" Value="#0064B4"/></DataTrigger>
                                </Style.Triggers>
                            </Style>
                        </DataGridTextColumn.CellStyle>
                    </DataGridTextColumn>
                    <DataGridTextColumn Header="Message" Binding="{Binding Message}"   Width="*"/>
                </DataGrid.Columns>
            </DataGrid>

            <!-- Status bar -->
            <Border Grid.Row="4" Background="#EEF0F5" Padding="10,6" Margin="0,8,0,0"
                    BorderBrush="#D2D7E4" BorderThickness="1">
                <TextBlock Name="TxtSetupStatus" Text="Ready" Foreground="#646C78"/>
            </Border>
        </Grid>
    </DockPanel>
</Window>
'@

    $reader = New-Object System.Xml.XmlNodeReader $setupXaml
    $script:aosWindow = [Windows.Markup.XamlReader]::Load($reader)

    $script:aosCtrl = @{}
    $setupXaml.SelectNodes("//*[@Name]") | ForEach-Object { $script:aosCtrl[$_.Name] = $script:aosWindow.FindName($_.Name) }

    # Load icon
    $_iconFile = Join-Path $PSScriptRoot 'FlyMigration.ico'
    if (-not (Test-Path $_iconFile)) { $_iconFile = Join-Path $PSScriptRoot 'ourvolaris.png' }
    if (Test-Path $_iconFile) {
        $dec = [System.Windows.Media.Imaging.BitmapDecoder]::Create(
            [System.Uri]::new($_iconFile),
            [System.Windows.Media.Imaging.BitmapCreateOptions]::None,
            [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
        $script:aosCtrl['ImgLogo'].Source = ($dec.Frames | Sort-Object PixelWidth | Select-Object -Last 1)
        try { $script:aosWindow.Icon = [System.Windows.Media.Imaging.BitmapImage]::new([System.Uri]::new($_iconFile)) } catch {}
    }

    $script:aosResults    = New-Object System.Collections.ObjectModel.ObservableCollection[object]
    $script:aosCurrentProc = $null
    $script:aosPendingById = @{}
    $script:aosCtrl['DgSetup'].ItemsSource = $script:aosResults

    function Set-AosStatus($msg) { $script:aosCtrl['TxtSetupStatus'].Text = $msg }

    function Invoke-AosUIDispatch {
        $script:aosWindow.Dispatcher.Invoke(
            [System.Windows.Threading.DispatcherPriority]::Background,
            [action]{}
        )
    }

    function Find-AosNodeExe {
        try { $c = Get-Command node.exe -ErrorAction Stop; if ($c.Source) { return $c.Source } } catch { }
        foreach ($p in @(
            "$env:ProgramFiles\nodejs\node.exe",
            "${env:ProgramFiles(x86)}\nodejs\node.exe",
            "$env:LOCALAPPDATA\Programs\nodejs\node.exe"
        )) { if ($p -and (Test-Path $p)) { return $p } }
        return $null
    }

    function Add-AosPendingRow($id, $tenant) {
        $row = [pscustomobject]@{
            Id        = $id
            Timestamp = (Get-Date).ToString('HH:mm:ss')
            Tenant    = $tenant
            Status    = 'WORKING'
            Message   = 'queued'
        }
        $script:aosResults.Add($row)
        $script:aosPendingById[$id] = $row
        $script:aosCtrl['DgSetup'].ScrollIntoView($row)
    }

    function Update-AosRow($id, $status, $message) {
        if ($script:aosPendingById.ContainsKey($id)) {
            $row = $script:aosPendingById[$id]
            $row.Timestamp = (Get-Date).ToString('HH:mm:ss')
            $row.Status    = $status
            $row.Message   = $message
            $script:aosCtrl['DgSetup'].Items.Refresh()
        }
    }

    function Invoke-AosConnectorLine($line) {
        try { $obj = $line | ConvertFrom-Json -ErrorAction Stop } catch {
            Set-AosStatus "(non-JSON) $line"
            return
        }
        if ($obj.event) {
            switch ($obj.event) {
                'info'     { Set-AosStatus $obj.message }
                'warn'     { Set-AosStatus "WARN: $($obj.message)" }
                'error'    { Set-AosStatus "ERROR: $($obj.message)" }
                'fatal'    { Set-AosStatus "FATAL: $($obj.message)" }
                'login-ok' {
                    Set-AosStatus "Signed in. Session saved."
                    $script:aosCtrl['TxtAuthStatus'].Text = "Signed in. Session saved."
                }
                'done'     { Set-AosStatus "Setup complete." }
                default    { Set-AosStatus "$($obj.event): $($obj.message)" }
            }
            return
        }
        if ($obj.id -and $obj.status) {
            Update-AosRow $obj.id $obj.status $obj.message
        }
    }

    function Invoke-AosConnector {
        param(
            [Parameter(Mandatory)][string]$Mode,
            [string[]]$StdinLines = @(),
            [string]$DisplayName  = ''
        )

        $node = Find-AosNodeExe
        if (-not $node) {
            [System.Windows.MessageBox]::Show(
                "Node.js was not found.`nInstall Node 18+ from https://nodejs.org and reopen.",
                "Node.js missing", 'OK', 'Error') | Out-Null
            return $false
        }
        if (-not (Test-Path $script:aosSetupConnectorJs)) {
            [System.Windows.MessageBox]::Show(
                "fly-connector.js not found: $script:aosSetupConnectorJs",
                "Connector missing", 'OK', 'Error') | Out-Null
            return $false
        }

        # ArgumentList lets .NET handle quoting — safe for paths containing spaces.
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName               = $node
        $psi.WorkingDirectory       = $PSScriptRoot
        $psi.ArgumentList.Add($script:aosSetupConnectorJs)
        $psi.ArgumentList.Add("--mode=$Mode")
        if ($DisplayName) { $psi.ArgumentList.Add("--display-name=$($DisplayName -replace '[^A-Za-z0-9._-]','_')") }
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.RedirectStandardInput  = $true
        $psi.UseShellExecute        = $false
        $psi.CreateNoWindow         = $true
        $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
        $psi.StandardErrorEncoding  = [System.Text.Encoding]::UTF8

        $proc = New-Object System.Diagnostics.Process
        $proc.StartInfo = $psi
        $proc.EnableRaisingEvents = $true

        Register-ObjectEvent -InputObject $proc -EventName OutputDataReceived -Action {
            param($src, $e)
            if (-not $e.Data) { return }
            $line = $e.Data
            $script:aosWindow.Dispatcher.Invoke([action]{ Invoke-AosConnectorLine $line })
        } | Out-Null

        Register-ObjectEvent -InputObject $proc -EventName ErrorDataReceived -Action {
            param($src, $e)
            if (-not $e.Data) { return }
            $line = $e.Data
            $script:aosWindow.Dispatcher.Invoke([action]{ Set-AosStatus "stderr: $line" })
        } | Out-Null

        $script:aosCurrentProc = $proc
        $null = $proc.Start()
        $proc.BeginOutputReadLine()
        $proc.BeginErrorReadLine()

        foreach ($line in $StdinLines) { $proc.StandardInput.WriteLine($line) }
        $proc.StandardInput.Close()

        while (-not $proc.HasExited) {
            Invoke-AosUIDispatch
            Start-Sleep -Milliseconds 100
        }
        Start-Sleep -Milliseconds 250
        Invoke-AosUIDispatch

        Get-EventSubscriber | Where-Object SourceObject -eq $proc | Unregister-Event
        $script:aosCurrentProc = $null
        return ($proc.ExitCode -eq 0)
    }

    # ── Auto-derive profile name from display name ────────────────────────────
    $script:aosCtrl['TxtDisplayName'].Add_TextChanged({
        $dn = $script:aosCtrl['TxtDisplayName'].Text.Trim()
        $current = $script:aosCtrl['TxtAppProfileName'].Text
        # Only auto-update if it still looks auto-generated (ends with " App" or is empty)
        if ([string]::IsNullOrWhiteSpace($current) -or $current -match ' App$') {
            if ($dn) { $script:aosCtrl['TxtAppProfileName'].Text = "$dn App" }
        }
    })

    # ── Sign In ───────────────────────────────────────────────────────────────
    $script:aosCtrl['BtnSignIn'].Add_Click({
        $script:aosCtrl['BtnSignIn'].IsEnabled    = $false
        $script:aosCtrl['BtnRunSetup'].IsEnabled  = $false
        Set-AosStatus "Launching browser for AOS sign-in..."
        try {
            Invoke-AosConnector -Mode 'login' | Out-Null
        } finally {
            $script:aosCtrl['BtnSignIn'].IsEnabled   = $true
            $script:aosCtrl['BtnRunSetup'].IsEnabled = $true
        }
    })

    # ── Clear results ─────────────────────────────────────────────────────────
    $script:aosCtrl['BtnClear'].Add_Click({
        $script:aosResults.Clear()
        $script:aosPendingById.Clear()
        Set-AosStatus "Cleared."
    })

    # ── Stop ──────────────────────────────────────────────────────────────────
    $script:aosCtrl['BtnStop'].Add_Click({
        if ($script:aosCurrentProc -and -not $script:aosCurrentProc.HasExited) {
            try { $script:aosCurrentProc.Kill() } catch { }
            Set-AosStatus "Stopped."
        }
    })

    # ── Run Setup ─────────────────────────────────────────────────────────────
    $script:aosCtrl['BtnRunSetup'].Add_Click({
        $displayName = $script:aosCtrl['TxtDisplayName'].Text.Trim()
        $searchCode  = $script:aosCtrl['TxtSearchCode'].Text.Trim()
        $appProfile  = $script:aosCtrl['TxtAppProfileName'].Text.Trim()
        $clientId    = $script:aosCtrl['TxtClientId'].Text.Trim()
        $clientSecret = $script:aosCtrl['TxtClientSecret'].Password

        if ([string]::IsNullOrWhiteSpace($displayName) -or
            [string]::IsNullOrWhiteSpace($searchCode)  -or
            [string]::IsNullOrWhiteSpace($appProfile)  -or
            [string]::IsNullOrWhiteSpace($clientId)    -or
            [string]::IsNullOrWhiteSpace($clientSecret)) {
            [System.Windows.MessageBox]::Show(
                "All fields are required:`n• Display Name`n• Search Code`n• Profile Name`n• Client ID`n• Client Secret",
                "Validation", 'OK', 'Warning') | Out-Null
            return
        }

        $id = [guid]::NewGuid().ToString('N').Substring(0, 8)
        $task = [pscustomobject]@{
            id                = $id
            tenantDisplayName = $displayName
            tenantSearch      = $searchCode
            appProfileName    = $appProfile
            clientId          = $clientId
            clientSecret      = $clientSecret
        }

        Add-AosPendingRow $id $displayName
        $stdinLines = @($task | ConvertTo-Json -Compress)

        # Persist display name + search code to shared config
        Update-SharedConfig @{
            TenantName   = $displayName
            TenantSearch = $searchCode
        }

        $script:aosCtrl['BtnRunSetup'].IsEnabled = $false
        $script:aosCtrl['BtnSignIn'].IsEnabled   = $false
        $script:aosCtrl['BtnStop'].IsEnabled     = $true
        Set-AosStatus "Running setup..."
        try {
            Invoke-AosConnector -Mode 'setup' -StdinLines $stdinLines -DisplayName $displayName | Out-Null
        } finally {
            $script:aosCtrl['BtnRunSetup'].IsEnabled = $true
            $script:aosCtrl['BtnSignIn'].IsEnabled   = $true
            $script:aosCtrl['BtnStop'].IsEnabled     = $false
        }
    })

    $script:aosCtrl['BtnClose'].Add_Click({ $script:aosWindow.Close() })

    # ── Pre-fill from shared config ───────────────────────────────────────────
    $_sc = Read-SharedConfig
    if ($_sc.TenantName)   { $script:aosCtrl['TxtDisplayName'].Text  = $_sc.TenantName }
    if ($_sc.TenantSearch) { $script:aosCtrl['TxtSearchCode'].Text   = $_sc.TenantSearch }
    if ($_sc.TenantName)   { $script:aosCtrl['TxtAppProfileName'].Text = "$($_sc.TenantName) App" }
    if ($_sc.TargetAppId)  { $script:aosCtrl['TxtClientId'].Text     = $_sc.TargetAppId }

    # Session status
    $authFile = Join-Path $PSScriptRoot 'auth\storageState.json'
    if (Test-Path $authFile) {
        $script:aosCtrl['TxtAuthStatus'].Text = "Previous session found. Sign in again only if expired."
    }

    [void]$script:aosWindow.ShowDialog()
}
