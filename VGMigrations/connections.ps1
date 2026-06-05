# ═════════════════════════════════════════════════════════════════════════════
# CONNECTION CREATOR (WPF)
# $script:ctrl / $script:wpfResults / $script:wpfWindow are $script:-scoped so
# nested helper functions can reach them without closure capture.
# ═════════════════════════════════════════════════════════════════════════════
function Show-ConnectionsForm {

    $script:connSvcDef = @(
        [pscustomobject]@{ Label='Exchange Online';      CheckboxName='ChkExo'    }
        [pscustomobject]@{ Label='SharePoint Online';    CheckboxName='ChkSpo'    }
        [pscustomobject]@{ Label='OneDrive';             CheckboxName='ChkOd4b'   }
        [pscustomobject]@{ Label='Microsoft Teams';      CheckboxName='ChkTeams'  }
        [pscustomobject]@{ Label='Microsoft Teams Chat'; CheckboxName='ChkChat'   }
        [pscustomobject]@{ Label='Microsoft 365 Groups'; CheckboxName='ChkGroups' }
    )

    $script:connectorJs  = Join-Path $PSScriptRoot 'fly-connector.js'
    $script:connSettings = Join-Path $env:LOCALAPPDATA 'FlyConnectionCreator\settings.json'

    [xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="AvePoint Fly - Connection Creator"
        Height="780" Width="980"
        WindowStartupLocation="CenterScreen"
        FontFamily="Segoe UI" FontSize="13"
        Background="#F0F2F7">
    <Window.Resources>
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="Margin" Value="0,2,0,8"/>
            <Setter Property="BorderBrush" Value="#D2D7E4"/>
            <Setter Property="Background" Value="White"/>
        </Style>
        <Style TargetType="Label">
            <Setter Property="Padding" Value="0,4,8,0"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="#1C1C20"/>
        </Style>
        <Style TargetType="GroupBox">
            <Setter Property="Margin" Value="0,0,0,10"/>
            <Setter Property="Padding" Value="12,8,12,8"/>
            <Setter Property="BorderBrush" Value="#D2D7E4"/>
            <Setter Property="Background" Value="White"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="14,6"/>
            <Setter Property="Margin" Value="0,0,8,0"/>
            <Setter Property="MinWidth" Value="140"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#F0F2F7"/>
            <Setter Property="Foreground" Value="#1C1C20"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="BorderBrush" Value="#D2D7E4"/>
            <Setter Property="BorderThickness" Value="0,0,0,1"/>
        </Style>
    </Window.Resources>

    <DockPanel>
        <Border DockPanel.Dock="Top" Background="#0064B4" Padding="16,8">
            <DockPanel LastChildFill="True">
                <Image Name="ImgLogo" DockPanel.Dock="Left" Height="34" Margin="0,0,12,0"
                       VerticalAlignment="Center" Stretch="Uniform"
                       RenderOptions.BitmapScalingMode="HighQuality"/>
                <TextBlock FontFamily="Segoe UI Semibold" FontSize="15" Foreground="White" VerticalAlignment="Center">
                    <Run Text="Connection Creator" FontWeight="Light"/>
                </TextBlock>
            </DockPanel>
        </Border>

        <Border DockPanel.Dock="Bottom" Background="#1C2030" Height="50" Padding="16,8">
            <DockPanel LastChildFill="True">
                <Button Name="BtnSignIn" DockPanel.Dock="Left" Content="Sign in to AOS..."
                        Background="#0064B4" Foreground="White" FontWeight="SemiBold"
                        MinWidth="160" Margin="0,0,12,0"/>
                <Button Name="BtnClose" DockPanel.Dock="Right" Content="Close"
                        Background="#C83737" Foreground="White" FontWeight="SemiBold"
                        MinWidth="90" Margin="0" HorizontalAlignment="Right"/>
                <TextBlock Name="TxtAuthStatus" VerticalAlignment="Center" Foreground="#BED2FF"
                           Text="Click to sign in - opens Chrome, complete Microsoft SSO, then close it."/>
            </DockPanel>
        </Border>

        <Grid Margin="16,12,16,16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <GroupBox Grid.Row="0" Header="Tenant">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="160"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="20"/>
                    <ColumnDefinition Width="160"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition/>
                    <RowDefinition/>
                </Grid.RowDefinitions>
                <Label   Grid.Row="0" Grid.Column="0" Content="Display Name:"/>
                <TextBox Grid.Row="0" Grid.Column="1" Name="TxtTenantName" ToolTip="Used in the Connection name, e.g. OurVolaris"/>
                <Label   Grid.Row="0" Grid.Column="3" Content="Search Code:"/>
                <TextBox Grid.Row="0" Grid.Column="4" Name="TxtTenantSearch" ToolTip="Short code shown in AOS Tenant dropdown, e.g. ourvolaris"/>
                <Label   Grid.Row="1" Grid.Column="0" Content="Credentials Name:"/>
                <TextBox Grid.Row="1" Grid.Column="1" Name="TxtCredentialsName" ToolTip="Substring matched against App profile and Service account dropdowns, e.g. ITVolaris"/>
            </Grid>
        </GroupBox>

        <GroupBox Grid.Row="1" Header="Workloads">
            <WrapPanel>
                <CheckBox Name="ChkExo"    Content="Exchange Online"       IsChecked="True" Margin="0,4,28,4"/>
                <CheckBox Name="ChkSpo"    Content="SharePoint Online"     IsChecked="True" Margin="0,4,28,4"/>
                <CheckBox Name="ChkOd4b"   Content="OneDrive"              IsChecked="True" Margin="0,4,28,4"/>
                <CheckBox Name="ChkTeams"  Content="Microsoft Teams"       IsChecked="True" Margin="0,4,28,4"/>
                <CheckBox Name="ChkChat"   Content="Microsoft Teams Chat"  IsChecked="True" Margin="0,4,28,4"/>
                <CheckBox Name="ChkGroups" Content="Microsoft 365 Groups"  IsChecked="True" Margin="0,4,28,4"/>
            </WrapPanel>
        </GroupBox>

        <Grid Grid.Row="2" Margin="0,4,0,12">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                <TextBlock Text="Connection type:" VerticalAlignment="Center" Margin="4,0,10,0"
                           FontWeight="SemiBold" Foreground="#1C1C20"/>
                <RadioButton Name="RdoDestination" Content="Destination tenant" IsChecked="True"
                             VerticalAlignment="Center" Margin="0,0,16,0"/>
                <RadioButton Name="RdoSource"      Content="Source tenant"
                             VerticalAlignment="Center"/>
            </StackPanel>
            <StackPanel Grid.Column="1" Orientation="Horizontal">
                <Button Name="BtnOpenLogs" Content="Open Logs Folder"/>
                <Button Name="BtnClear"    Content="Clear Results"/>
                <Button Name="BtnSaveLog"  Content="Save Log..."/>
                <Button Name="BtnCancel"   Content="Stop"             IsEnabled="False"/>
                <Button Name="BtnCreate"   Content="Create Connections"
                        Background="#0064B4" Foreground="White" FontWeight="SemiBold"/>
            </StackPanel>
        </Grid>

        <DataGrid Grid.Row="3" Name="DgResults" AutoGenerateColumns="False"
                  IsReadOnly="True" HeadersVisibility="Column"
                  GridLinesVisibility="Horizontal"
                  AlternatingRowBackground="#F5F6FA"
                  Background="White" BorderBrush="#D2D7E4"
                  RowHeight="28" FontSize="12">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Time"             Binding="{Binding Timestamp}"      Width="80"/>
                <DataGridTextColumn Header="Workload"         Binding="{Binding Workload}"       Width="180"/>
                <DataGridTextColumn Header="Connection Name"  Binding="{Binding ConnectionName}" Width="260"/>
                <DataGridTextColumn Header="Status"           Binding="{Binding Status}"         Width="100">
                    <DataGridTextColumn.CellStyle>
                        <Style TargetType="DataGridCell">
                            <Setter Property="FontWeight" Value="SemiBold"/>
                            <Setter Property="Padding" Value="6,0"/>
                            <Style.Triggers>
                                <DataTrigger Binding="{Binding Status}" Value="CREATED"><Setter Property="Foreground" Value="#107C10"/></DataTrigger>
                                <DataTrigger Binding="{Binding Status}" Value="FAILED"> <Setter Property="Foreground" Value="#D13438"/></DataTrigger>
                                <DataTrigger Binding="{Binding Status}" Value="SKIPPED"><Setter Property="Foreground" Value="#808080"/></DataTrigger>
                                <DataTrigger Binding="{Binding Status}" Value="WORKING"><Setter Property="Foreground" Value="#0064B4"/></DataTrigger>
                            </Style.Triggers>
                        </Style>
                    </DataGridTextColumn.CellStyle>
                </DataGridTextColumn>
                <DataGridTextColumn Header="Message" Binding="{Binding Message}" Width="*"/>
            </DataGrid.Columns>
        </DataGrid>

        <Border Grid.Row="4" Background="#EEF0F5" Padding="10,6" Margin="0,8,0,0" BorderBrush="#D2D7E4" BorderThickness="1">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Name="TxtStatus"  Text="Ready" Foreground="#646C78"/>
                <TextBlock Grid.Column="1" Name="TxtSummary" Text="" FontWeight="SemiBold" Foreground="#1C1C20"/>
            </Grid>
        </Border>
        </Grid>
    </DockPanel>
</Window>
'@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $script:wpfWindow = [Windows.Markup.XamlReader]::Load($reader)

    $script:ctrl = @{}
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object { $script:ctrl[$_.Name] = $script:wpfWindow.FindName($_.Name) }

    # Load icon (prefer FlyMigration.ico, fall back to ourvolaris.png)
    $_iconFile = Join-Path $PSScriptRoot "FlyMigration.ico"
    if (-not (Test-Path $_iconFile)) { $_iconFile = Join-Path $PSScriptRoot "ourvolaris.png" }
    if (Test-Path $_iconFile) {
        $dec = [System.Windows.Media.Imaging.BitmapDecoder]::Create(
            [System.Uri]::new($_iconFile),
            [System.Windows.Media.Imaging.BitmapCreateOptions]::None,
            [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
        $script:ctrl['ImgLogo'].Source = ($dec.Frames | Sort-Object PixelWidth | Select-Object -Last 1)
        try { $script:wpfWindow.Icon = [System.Windows.Media.Imaging.BitmapImage]::new([System.Uri]::new($_iconFile)) } catch {}
    }

    $script:wpfResults = New-Object System.Collections.ObjectModel.ObservableCollection[object]
    $script:ctrl['DgResults'].ItemsSource = $script:wpfResults

    $script:currentProcess = $null
    $script:pendingByid    = @{}

    function Set-Status($message) { $script:ctrl['TxtStatus'].Text = $message }

    function Update-Summary {
        $c = ($script:wpfResults | Where-Object Status -eq 'CREATED').Count
        $s = ($script:wpfResults | Where-Object Status -eq 'SKIPPED').Count
        $f = ($script:wpfResults | Where-Object Status -eq 'FAILED' ).Count
        $parts = @()
        if ($c) { $parts += "$c created" }
        if ($s) { $parts += "$s skipped" }
        if ($f) { $parts += "$f failed" }
        $script:ctrl['TxtSummary'].Text = $parts -join '  |  '
    }

    function Invoke-UIDispatch {
        $script:wpfWindow.Dispatcher.Invoke(
            [System.Windows.Threading.DispatcherPriority]::Background,
            [action]{}
        )
    }

    function Find-NodeExe {
        try {
            $cmd = Get-Command node.exe -ErrorAction Stop
            if ($cmd -and $cmd.Source) { return $cmd.Source }
        } catch { }
        foreach ($p in @(
            "$env:ProgramFiles\nodejs\node.exe",
            "${env:ProgramFiles(x86)}\nodejs\node.exe",
            "$env:LOCALAPPDATA\Programs\nodejs\node.exe"
        )) { if ($p -and (Test-Path $p)) { return $p } }
        return $null
    }

    function Import-ConnectionSettings {
        $local  = $null
        $shared = Read-SharedConfig
        if (Test-Path $script:connSettings) {
            try { $local = Get-Content $script:connSettings -Raw | ConvertFrom-Json } catch { }
        }
        $tn = if ($local.TenantName)      { $local.TenantName }      else { $shared.TenantName }
        $ts = if ($local.TenantSearch)    { $local.TenantSearch }    else { $shared.TenantSearch }
        $cn = if ($local.CredentialsName) { $local.CredentialsName } else { $shared.CredentialsName }
        if ($tn) { $script:ctrl['TxtTenantName'].Text      = $tn }
        if ($ts) { $script:ctrl['TxtTenantSearch'].Text    = $ts }
        if ($cn) { $script:ctrl['TxtCredentialsName'].Text = $cn }
    }

    function Save-Settings {
        try {
            $dir = Split-Path $script:connSettings
            if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            @{
                TenantName      = $script:ctrl['TxtTenantName'].Text
                TenantSearch    = $script:ctrl['TxtTenantSearch'].Text
                CredentialsName = $script:ctrl['TxtCredentialsName'].Text
            } | ConvertTo-Json | Set-Content $script:connSettings -Encoding UTF8
            Update-SharedConfig @{
                TenantName      = $script:ctrl['TxtTenantName'].Text
                TenantSearch    = $script:ctrl['TxtTenantSearch'].Text
                CredentialsName = $script:ctrl['TxtCredentialsName'].Text
            }
        } catch { }
    }

    function Add-PendingRow($id, $workload, $connectionName) {
        $row = [pscustomobject]@{
            Id             = $id
            Timestamp      = (Get-Date).ToString('HH:mm:ss')
            Workload       = $workload
            ConnectionName = $connectionName
            Status         = 'WORKING'
            Message        = 'queued'
        }
        $script:wpfResults.Add($row)
        $script:pendingByid[$id] = $row
        $script:ctrl['DgResults'].ScrollIntoView($row)
        Update-Summary
    }

    function Update-Row($id, $status, $message) {
        if ($script:pendingByid.ContainsKey($id)) {
            $row = $script:pendingByid[$id]
            $row.Timestamp = (Get-Date).ToString('HH:mm:ss')
            $row.Status    = $status
            $row.Message   = $message
            $script:ctrl['DgResults'].Items.Refresh()
            Update-Summary
        }
    }

    function Invoke-Connector {
        param(
            [Parameter(Mandatory)] [string]$Mode,
            [string[]]$StdinLines = @(),
            [string]$DisplayName = ''
        )

        $node = Find-NodeExe
        if (-not $node) {
            [System.Windows.MessageBox]::Show(
                "Node.js was not found on PATH.`n`nInstall Node 18+ from https://nodejs.org and reopen this GUI.",
                "Node.js missing",'OK','Error') | Out-Null
            return $false
        }
        if (-not (Test-Path $script:connectorJs)) {
            [System.Windows.MessageBox]::Show(
                "fly-connector.js not found next to this script.`n`nExpected: $script:connectorJs",
                "Connector missing",'OK','Error') | Out-Null
            return $false
        }

        # ArgumentList lets .NET handle quoting — safe for paths containing spaces.
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName               = $node
        $psi.WorkingDirectory       = $PSScriptRoot
        $psi.ArgumentList.Add($script:connectorJs)
        $psi.ArgumentList.Add("--mode=$Mode")
        if ($DisplayName) {
            $psi.ArgumentList.Add("--display-name=$($DisplayName -replace '[^A-Za-z0-9._-]','_')")
        }
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

        $stdoutHandler = {
            param($src, $e)
            if (-not $e.Data) { return }
            $line = $e.Data
            $script:wpfWindow.Dispatcher.Invoke([action]{
                Invoke-ConnectorLine $line
            })
        }
        $stderrHandler = {
            param($src, $e)
            if (-not $e.Data) { return }
            $line = $e.Data
            $script:wpfWindow.Dispatcher.Invoke([action]{
                Set-Status "stderr: $line"
            })
        }
        Register-ObjectEvent -InputObject $proc -EventName OutputDataReceived -Action $stdoutHandler | Out-Null
        Register-ObjectEvent -InputObject $proc -EventName ErrorDataReceived  -Action $stderrHandler | Out-Null

        $script:currentProcess = $proc
        $null = $proc.Start()
        $proc.BeginOutputReadLine()
        $proc.BeginErrorReadLine()

        foreach ($line in $StdinLines) { $proc.StandardInput.WriteLine($line) }
        $proc.StandardInput.Close()

        while (-not $proc.HasExited) {
            Invoke-UIDispatch
            Start-Sleep -Milliseconds 100
        }
        Start-Sleep -Milliseconds 250
        Invoke-UIDispatch

        Get-EventSubscriber | Where-Object SourceObject -eq $proc | Unregister-Event
        $script:currentProcess = $null
        return ($proc.ExitCode -eq 0)
    }

    function Invoke-ConnectorLine($line) {
        try { $obj = $line | ConvertFrom-Json -ErrorAction Stop } catch {
            Set-Status "(non-JSON) $line"
            return
        }
        if ($obj.event) {
            switch ($obj.event) {
                'info'     { Set-Status $obj.message }
                'warn'     { Set-Status "WARN: $($obj.message)" }
                'error'    { Set-Status "ERROR: $($obj.message)" }
                'fatal'    { Set-Status "FATAL: $($obj.message)" }
                'login-ok' { Set-Status "Signed in. $($obj.message)"; $script:ctrl['TxtAuthStatus'].Text = "Signed in. Session saved." }
                'done'     { Set-Status "Connector finished." }
                default    { Set-Status "$($obj.event): $($obj.message)" }
            }
            return
        }
        if ($obj.id -and $obj.status) {
            Update-Row $obj.id $obj.status $obj.message
            return
        }
    }

    $script:ctrl['BtnSignIn'].Add_Click({
        $script:ctrl['BtnSignIn'].IsEnabled = $false
        $script:ctrl['BtnCreate'].IsEnabled = $false
        Set-Status "Launching browser for Microsoft SSO sign-in..."
        $script:ctrl['TxtAuthStatus'].Text = "Signing in - complete the SSO flow in the Chrome window, then close it."
        try {
            Invoke-Connector -Mode 'login' | Out-Null
        } finally {
            $script:ctrl['BtnSignIn'].IsEnabled = $true
            $script:ctrl['BtnCreate'].IsEnabled = $true
        }
    })

    $script:ctrl['BtnClear'].Add_Click({
        $script:wpfResults.Clear()
        $script:pendingByid.Clear()
        Update-Summary
        Set-Status "Results cleared."
    })

    $script:ctrl['BtnOpenLogs'].Add_Click({
        $logsDir = Join-Path $PSScriptRoot 'logs'
        if (-not (Test-Path $logsDir)) { New-Item -ItemType Directory -Path $logsDir -Force | Out-Null }
        Start-Process explorer.exe $logsDir
        Set-Status "Opened: $logsDir"
    })

    $script:ctrl['BtnSaveLog'].Add_Click({
        if ($script:wpfResults.Count -eq 0) {
            [System.Windows.MessageBox]::Show("Nothing to save.","Save Log",'OK','Information') | Out-Null
            return
        }
        $dlg = New-Object Microsoft.Win32.SaveFileDialog
        $dlg.Filter = "CSV (*.csv)|*.csv"
        $tag = $script:ctrl['TxtTenantName'].Text
        if ([string]::IsNullOrWhiteSpace($tag)) { $tag = 'tenant' }
        $dlg.FileName = "fly-connections-$tag-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
        if ($dlg.ShowDialog()) {
            $script:wpfResults | Select-Object Timestamp, Workload, ConnectionName, Status, Message |
                Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
            Set-Status "Log saved: $($dlg.FileName)"
        }
    })

    $script:ctrl['BtnCancel'].Add_Click({
        if ($script:currentProcess -and -not $script:currentProcess.HasExited) {
            try { $script:currentProcess.Kill() } catch { }
            Set-Status "Cancelled."
        }
    })

    $script:ctrl['BtnCreate'].Add_Click({
        $tenantName = $script:ctrl['TxtTenantName'].Text.Trim()
        $tenantSrch = $script:ctrl['TxtTenantSearch'].Text.Trim()
        $credName   = $script:ctrl['TxtCredentialsName'].Text.Trim()

        if ([string]::IsNullOrWhiteSpace($tenantName) -or [string]::IsNullOrWhiteSpace($tenantSrch) -or [string]::IsNullOrWhiteSpace($credName)) {
            [System.Windows.MessageBox]::Show("Enter Display Name, Search Code, and Credentials Name.","Validation",'OK','Warning') | Out-Null
            return
        }

        $selected = $script:connSvcDef | Where-Object { $script:ctrl[$_.CheckboxName].IsChecked }
        if (-not $selected) {
            [System.Windows.MessageBox]::Show("Select at least one workload.","Validation",'OK','Warning') | Out-Null
            return
        }

        Save-Settings

        $tasks = @()
        foreach ($svc in $selected) {
            $id   = [guid]::NewGuid().ToString('N').Substring(0,8)
            $name = "$tenantName - $($svc.Label)"
            $tasks += [pscustomobject]@{
                id              = $id
                tenantName      = $tenantName
                tenantSearch    = $tenantSrch
                workloadLabel   = $svc.Label
                connectionName  = $name
                credentialsName = $credName
            }
            Add-PendingRow $id $svc.Label $name
        }

        $stdinLines = $tasks | ForEach-Object { $_ | ConvertTo-Json -Compress }

        $script:ctrl['BtnCreate'].IsEnabled = $false
        $script:ctrl['BtnSignIn'].IsEnabled = $false
        $script:ctrl['BtnCancel'].IsEnabled = $true

        $labelToKey = @{
            'Exchange Online'       = 'Exchange'
            'SharePoint Online'     = 'SharePoint'
            'OneDrive'              = 'OneDrive'
            'Microsoft Teams'       = 'Teams'
            'Microsoft Teams Chat'  = 'Teams Chat'
            'Microsoft 365 Groups'  = 'Groups'
        }

        Set-Status "Driving the AOS portal..."
        try {
            $ok = Invoke-Connector -Mode 'create' -StdinLines $stdinLines -DisplayName $tenantName

            $workloadsJsonPath = Join-Path $PSScriptRoot 'workloads.json'
            $wlJson = if (Test-Path $workloadsJsonPath) {
                try { Get-Content $workloadsJsonPath -Raw | ConvertFrom-Json } catch { [pscustomobject]@{} }
            } else { [pscustomobject]@{} }

            $connKeyName = if ($script:ctrl['RdoSource'].IsChecked) { 'Source' } else { 'Destination' }
            $updated = 0
            foreach ($task in $tasks) {
                $row = $script:pendingByid[$task.id]
                if ($row -and $row.Status -eq 'CREATED') {
                    $key = $labelToKey[$task.workloadLabel]
                    if ($key) {
                        if (-not $wlJson.PSObject.Properties[$key]) {
                            $wlJson | Add-Member -NotePropertyName $key -NotePropertyValue ([pscustomobject]@{ Policy = ""; Source = ""; Destination = "" })
                        }
                        $wlJson.$key.$connKeyName = $task.connectionName
                        $updated++
                    }
                }
            }
            if ($updated -gt 0) {
                $wlJson | ConvertTo-Json -Depth 3 | Set-Content $workloadsJsonPath -Encoding UTF8
            }

            $summary = if ($ok) { "Done - $tenantName complete." } else { "Finished with errors." }
            if ($updated -gt 0) { $summary += "  $updated $connKeyName connection(s) written to workloads.json." }
            Set-Status $summary
        } finally {
            $script:ctrl['BtnCreate'].IsEnabled = $true
            $script:ctrl['BtnSignIn'].IsEnabled = $true
            $script:ctrl['BtnCancel'].IsEnabled = $false
        }
    })

    $script:ctrl['BtnClose'].Add_Click({ $script:wpfWindow.Close() })

    Import-ConnectionSettings
    Update-Summary

    $node = Find-NodeExe
    if (-not $node) {
        $script:ctrl['TxtAuthStatus'].Text = "Node.js not found on PATH. Install Node 18+ from https://nodejs.org."
        Set-Status "Node.js missing - install Node 18+ and reopen."
    } elseif (-not (Test-Path $script:connectorJs)) {
        $script:ctrl['TxtAuthStatus'].Text = "fly-connector.js not found next to this script."
        Set-Status "Place fly-connector.js + package.json next to this script and run 'npm install' there."
    } else {
        $authFile = Join-Path $PSScriptRoot 'auth\storageState.json'
        if (Test-Path $authFile) {
            $script:ctrl['TxtAuthStatus'].Text = "Previous session found. Sign in again only if it has expired."
        }
        Set-Status "Ready"
    }

    [void]$script:wpfWindow.ShowDialog()
}
