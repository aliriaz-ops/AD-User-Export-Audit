#Requires -Version 5.1
<#
.SYNOPSIS
    WPF GUI app: AD User Export & Audit.

.DESCRIPTION
    Dark-themed WPF interface to:
      - Select date range (whenChanged)
      - Choose user status: Disabled / Active / Both
      - Optional OU scope filter
      - Optional DisplayName wildcard filter
      - Column selection (properties to export)
      - Live preview grid
      - CSV / Out-GridView output
      - Transcript logging to %TEMP%

.NOTES
    Requires: ActiveDirectory module (RSAT)
    Run as a user with read access to AD.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── Transcript log ────────────────────────────────────────────────────────────
$LogPath = Join-Path $env:TEMP ("ADUserExportAudit_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))
Start-Transcript -Path $LogPath -Append | Out-Null

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms   # for SaveFileDialog

# ── XAML ─────────────────────────────────────────────────────────────────────
[xml]$Xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="AD User Export &amp; Audit"
        Height="660" Width="700"
        MinHeight="560" MinWidth="580"
        WindowStartupLocation="CenterScreen"
        Background="#1A1A2E"
        Foreground="#E0E0E0"
        FontFamily="Segoe UI"
        FontSize="13">
</Window>
'@

    <Window.Resources>
        <!-- Base dark control style -->
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="#C8C8D4"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="Padding" Value="0"/>
        </Style>

        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#16213E"/>
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="BorderBrush" Value="#3A3A5C"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="CaretBrush" Value="#7B9FFF"/>
        </Style>

        <Style TargetType="DatePicker">
            <Setter Property="Background" Value="#16213E"/>
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="BorderBrush" Value="#3A3A5C"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="4"/>
        </Style>

        <Style TargetType="Button">
            <Setter Property="Background" Value="#4A5568"/>
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="14,7"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#5A6A84"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#374151"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Accent button (primary action) -->
        <Style x:Key="AccentButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#3B5BDB"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#4C6EF5"/>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                    <Setter Property="Background" Value="#2F4AC0"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- Danger button -->
        <Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#C0392B"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#E74C3C"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="#C8C8D4"/>
            <Setter Property="Margin" Value="0,2"/>
        </Style>

        <Style TargetType="RadioButton">
            <Setter Property="Foreground" Value="#C8C8D4"/>
            <Setter Property="Margin" Value="0,2"/>
        </Style>

        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="#16213E"/>
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="BorderBrush" Value="#3A3A5C"/>
            <Setter Property="RowBackground" Value="#1E2A45"/>
            <Setter Property="AlternatingRowBackground" Value="#1A2438"/>
            <Setter Property="HorizontalGridLinesBrush" Value="#2A2A4A"/>
            <Setter Property="VerticalGridLinesBrush" Value="#2A2A4A"/>
            <Setter Property="SelectionMode" Value="Extended"/>
            <Setter Property="AutoGenerateColumns" Value="False"/>
            <Setter Property="IsReadOnly" Value="True"/>
            <Setter Property="CanUserAddRows" Value="False"/>
        </Style>

        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#0F3460"/>
            <Setter Property="Foreground" Value="#7B9FFF"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="BorderBrush" Value="#3A3A5C"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>

        <Style TargetType="ProgressBar">
            <Setter Property="Background" Value="#16213E"/>
            <Setter Property="Foreground" Value="#3B5BDB"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Height" Value="4"/>
        </Style>

        <Style TargetType="GroupBox">
            <Setter Property="Foreground" Value="#7B9FFF"/>
            <Setter Property="BorderBrush" Value="#3A3A5C"/>
            <Setter Property="Margin" Value="0,4"/>
            <Setter Property="Padding" Value="8,6"/>
        </Style>
    </Window.Resources>

    <DockPanel>
        <!-- Title bar accent strip -->
        <Border DockPanel.Dock="Top" Height="3" Background="#3B5BDB"/>

        <!-- Status bar -->
        <Border DockPanel.Dock="Bottom" Background="#0F0F1E" Padding="10,4">
            <DockPanel>
                <ProgressBar x:Name="ProgressBar" DockPanel.Dock="Right"
                             Width="140" Visibility="Collapsed"
                             IsIndeterminate="True" Margin="8,0,0,0"/>
                <TextBlock x:Name="StatusText" Foreground="#7B9FFF"
                           VerticalAlignment="Center"
                           Text="Ready — select status, date range and click Search."/>
            </DockPanel>
        </Border>

        <ScrollViewer VerticalScrollBarVisibility="Auto">
            <StackPanel Margin="16,12">

                <!-- ── Status + Date Range ── -->
                <Grid Margin="0,0,0,4">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="2*"/>
                    </Grid.ColumnDefinitions>

                    <GroupBox Header="👥  User Status" Grid.Column="0" Margin="0,0,8,0">
                        <StackPanel>
                            <RadioButton x:Name="StatusDisabled" Content="Disabled only" IsChecked="True"/>
                            <RadioButton x:Name="StatusEnabled"  Content="Active (enabled) only"/>
                            <RadioButton x:Name="StatusBoth"     Content="Both disabled and active"/>
                        </StackPanel>
                    </GroupBox>

                    <GroupBox Header="📅  Date Range  (whenChanged)" Grid.Column="1">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="90"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="90"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Label Grid.Column="0" Content="Start date:"/>
                            <DatePicker x:Name="StartDatePicker" Grid.Column="1" Margin="0,0,12,0"/>
                            <Label Grid.Column="2" Content="End date:"/>
                            <DatePicker x:Name="EndDatePicker"   Grid.Column="3"/>
                        </Grid>
                    </GroupBox>
                </Grid>

                <!-- ── Scope / Filter ── -->
                <GroupBox Header="🔍  Scope &amp; Filter">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="90"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>

                        <Label Grid.Row="0" Grid.Column="0" Content="OU (DN):"/>
                        <TextBox x:Name="OuTextBox" Grid.Row="0" Grid.Column="1"
                                 Margin="0,0,8,4"
                                 ToolTip="Optional. E.g.: OU=Users,DC=corp,DC=local  — leave blank to search all."/>
                        <Button x:Name="BrowseOuButton" Grid.Row="0" Grid.Column="2"
                                Content="Help…" Width="72" Margin="0,0,0,4"/>

                        <Label Grid.Row="1" Grid.Column="0" Content="Name filter:"/>
                        <TextBox x:Name="NameFilterBox" Grid.Row="1" Grid.Column="1"
                                 Grid.ColumnSpan="2"
                                 ToolTip="Optional wildcard filter on DisplayName. E.g.: John*"/>
                    </Grid>
                </GroupBox>

                <!-- ── Columns to export ── -->
                <GroupBox Header="📋  Columns to Export">
                    <WrapPanel x:Name="ColumnPanel" Orientation="Horizontal">
                        <CheckBox x:Name="ColDisplayName"    Content="DisplayName"    IsChecked="True" Margin="0,2,16,2"/>
                        <CheckBox x:Name="ColSam"            Content="SamAccountName" IsChecked="True" Margin="0,2,16,2"/>
                        <CheckBox x:Name="ColUPN"            Content="UserPrincipalName" IsChecked="True" Margin="0,2,16,2"/>
                        <CheckBox x:Name="ColMail"           Content="EmailAddress"   IsChecked="True" Margin="0,2,16,2"/>
                        <CheckBox x:Name="ColEnabled"        Content="Enabled"        IsChecked="True" Margin="0,2,16,2"/>
                        <CheckBox x:Name="ColDisabledDate"   Content="DisabledDate"   IsChecked="True" Margin="0,2,16,2"/>
                        <CheckBox x:Name="ColWhenChanged"    Content="whenChanged"    IsChecked="False" Margin="0,2,16,2"/>
                        <CheckBox x:Name="ColWhenCreated"    Content="whenCreated"    IsChecked="False" Margin="0,2,16,2"/>
                        <CheckBox x:Name="ColLastLogon"      Content="LastLogonDate"  IsChecked="True" Margin="0,2,16,2"/>
                        <CheckBox x:Name="ColDescription"    Content="Description"    IsChecked="True" Margin="0,2,16,2"/>
                        <CheckBox x:Name="ColDept"           Content="Department"     IsChecked="False" Margin="0,2,16,2"/>
                        <CheckBox x:Name="ColTitle"          Content="Title"          IsChecked="False" Margin="0,2,16,2"/>
                        <CheckBox x:Name="ColManager"        Content="Manager"        IsChecked="False" Margin="0,2,16,2"/>
                        <CheckBox x:Name="ColOU"             Content="OU"             IsChecked="True" Margin="0,2,16,2"/>
                    </WrapPanel>
                </GroupBox>

                <!-- ── Action buttons ── -->
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,8">
                    <Button x:Name="SearchButton"  Style="{StaticResource AccentButton}"
                            Content="🔍  Search AD"  Width="140" Margin="0,0,10,0"/>
                    <Button x:Name="ExportCsvButton" Content="💾  Export CSV"
                            Width="130" Margin="0,0,10,0" IsEnabled="False"/>
                    <Button x:Name="GridViewButton"  Content="📊  Open GridView"
                            Width="130" Margin="0,0,10,0" IsEnabled="False"/>
                    <Button x:Name="ClearButton" Style="{StaticResource DangerButton}"
                            Content="✖  Clear"  Width="80"/>
                </StackPanel>

                <!-- ── Results preview ── -->
                <GroupBox x:Name="ResultsGroup" Header="Results Preview  (0 records)" Margin="0,0,0,8">
                    <DataGrid x:Name="ResultsGrid" Height="240"
                              ColumnWidth="*" GridLinesVisibility="Horizontal">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="DisplayName"     Binding="{Binding DisplayName}"/>
                            <DataGridTextColumn Header="SamAccountName"  Binding="{Binding SamAccountName}"/>
                            <DataGridTextColumn Header="Email"           Binding="{Binding EmailAddress}"/>
                            <DataGridTextColumn Header="Enabled"         Binding="{Binding Enabled}"/>
                            <DataGridTextColumn Header="DisabledDate"    Binding="{Binding DisabledDate}"/>
                            <DataGridTextColumn Header="OU"              Binding="{Binding OU}"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </GroupBox>

            </StackPanel>
        </ScrollViewer>
    </DockPanel>
</Window>
'@

# ── Load XAML ─────────────────────────────────────────────────────────────────
$reader = New-Object System.Xml.XmlNodeReader $Xaml
$Window = [Windows.Markup.XamlReader]::Load($reader)

# ── Helper: find named controls ───────────────────────────────────────────────
function Get-Control { param([string]$Name) $Window.FindName($Name) }

$StartDatePicker  = Get-Control 'StartDatePicker'
$EndDatePicker    = Get-Control 'EndDatePicker'
$OuTextBox        = Get-Control 'OuTextBox'
$NameFilterBox    = Get-Control 'NameFilterBox'
$SearchButton     = Get-Control 'SearchButton'
$ExportCsvButton  = Get-Control 'ExportCsvButton'
$GridViewButton   = Get-Control 'GridViewButton'
$ClearButton      = Get-Control 'ClearButton'
$ResultsGrid      = Get-Control 'ResultsGrid'
$ResultsGroup     = Get-Control 'ResultsGroup'
$StatusText       = Get-Control 'StatusText'
$ProgressBar      = Get-Control 'ProgressBar'
$BrowseOuButton   = Get-Control 'BrowseOuButton'

$StatusDisabled   = Get-Control 'StatusDisabled'
$StatusEnabled    = Get-Control 'StatusEnabled'
$StatusBoth       = Get-Control 'StatusBoth'

# Column checkboxes
$ColMap = [ordered]@{
    DisplayName         = Get-Control 'ColDisplayName'
    SamAccountName      = Get-Control 'ColSam'
    UserPrincipalName   = Get-Control 'ColUPN'
    EmailAddress        = Get-Control 'ColMail'
    Enabled             = Get-Control 'ColEnabled'
    DisabledDate        = Get-Control 'ColDisabledDate'
    whenChanged         = Get-Control 'ColWhenChanged'
    whenCreated         = Get-Control 'ColWhenCreated'
    LastLogonDate       = Get-Control 'ColLastLogon'
    Description         = Get-Control 'ColDescription'
    Department          = Get-Control 'ColDept'
    Title               = Get-Control 'ColTitle'
    Manager             = Get-Control 'ColManager'
    OU                  = Get-Control 'ColOU'
}

# ── Smart defaults ────────────────────────────────────────────────────────────
$StartDatePicker.SelectedDate = (Get-Date).AddMonths(-2).Date
$EndDatePicker.SelectedDate   = (Get-Date).Date

# ── Module check ──────────────────────────────────────────────────────────────
$adAvailable = $false
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    $adAvailable = $true
} catch {
    $StatusText.Text = '⚠  ActiveDirectory module not found — results will be simulated.'
    Write-Warning "ActiveDirectory module unavailable: $_"
}

# ── Shared state ──────────────────────────────────────────────────────────────
$script:QueryResults = $null

# ── Helpers ───────────────────────────────────────────────────────────────────
function Set-Status {
    param([string]$Message, [bool]$Busy = $false)
    $StatusText.Text        = $Message
    $ProgressBar.Visibility = if ($Busy) { 'Visible' } else { 'Collapsed' }
}

function Get-SelectedColumns {
    $ColMap.Keys | Where-Object { $ColMap[$_].IsChecked }
}

function Get-OUFromDN ([string]$dn) {
    if ($dn -match '(?:^|,)(OU=[^,]+)') { return $Matches[1] }
    return $dn -replace '^CN=[^,]+,', '' -replace ',DC=.*$', ''
}

function Invoke-ADQuery {
    param(
        [datetime]$Start,
        [datetime]$End,
        [string]$OuDN,
        [string]$NameFilter,
        [string]$StatusMode  # "Disabled", "Enabled", "Both"
    )

    $adProps = @(
        'DisplayName','SamAccountName','UserPrincipalName','EmailAddress',
        'Enabled','whenChanged','whenCreated','LastLogonDate',
        'Description','Department','Title','Manager','DistinguishedName'
    )

    # Base Get-ADUser parameters
    $params = @{
        Filter     = { Enabled -eq $true -or Enabled -eq $false }  # all users
        Properties = $adProps
    }
    if ($OuDN) { $params['SearchBase'] = $OuDN }

    $users = Get-ADUser @params |
        Where-Object {
            $_.whenChanged -ge $Start -and
            $_.whenChanged -lt  $End -and
            (-not $NameFilter -or $_.DisplayName -like $NameFilter) -and
            (
                ($StatusMode -eq 'Disabled' -and $_.Enabled -eq $false) -or
                ($StatusMode -eq 'Enabled'  -and $_.Enabled -eq $true)  -or
                ($StatusMode -eq 'Both')
            )
        } |
        ForEach-Object {
            [PSCustomObject]@{
                DisplayName       = $_.DisplayName
                SamAccountName    = $_.SamAccountName
                UserPrincipalName = $_.UserPrincipalName
                EmailAddress      = $_.EmailAddress
                Enabled           = $_.Enabled
                whenChanged       = $_.whenChanged
                DisabledDate      = if (-not $_.Enabled -and $_.whenChanged) { $_.whenChanged.ToString('yyyy-MM-dd') } else { $null }
                whenCreated       = $_.whenCreated
                LastLogonDate     = $_.LastLogonDate
                Description       = $_.Description
                Department        = $_.Department
                Title             = $_.Title
                Manager           = if ($_.Manager) { ($_.Manager -split ',')[0] -replace '^CN=','' } else { '' }
                OU                = Get-OUFromDN $_.DistinguishedName
            }
        }

    return @($users)
}

function Update-PreviewGrid ([object[]]$Data) {
    $ResultsGrid.ItemsSource = $Data
    $ResultsGroup.Header     = "Results Preview  ($($Data.Count) record$(if($Data.Count -ne 1){'s'}))"
}

# ── Browse OU helper ──────────────────────────────────────────────────────────
$BrowseOuButton.Add_Click({
    [System.Windows.MessageBox]::Show(
        "Enter the full Distinguished Name of the OU in the text box.`n`nExample:`n  OU=Staff,OU=Users,DC=corp,DC=local",
        "OU Help",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    ) | Out-Null
    $OuTextBox.Focus() | Out-Null
})

# ── Validation ────────────────────────────────────────────────────────────────
function Test-Inputs {
    if (-not $StartDatePicker.SelectedDate -or -not $EndDatePicker.SelectedDate) {
        [System.Windows.MessageBox]::Show("Please select both start and end dates.", "Missing dates",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return $false
    }
    $s = [datetime]$StartDatePicker.SelectedDate
    $e = [datetime]$EndDatePicker.SelectedDate
    if ($e -le $s) {
        [System.Windows.MessageBox]::Show("End date must be after start date.", "Invalid range",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return $false
    }
    return $true
}

function Get-StatusMode {
    if ($StatusDisabled.IsChecked) { return 'Disabled' }
    if ($StatusEnabled.IsChecked)  { return 'Enabled'  }
    return 'Both'
}

# ── Search ────────────────────────────────────────────────────────────────────
$SearchButton.Add_Click({
    if (-not (Test-Inputs)) { return }

    $StartDate  = [datetime]$StartDatePicker.SelectedDate
    $EndDate    = [datetime]$EndDatePicker.SelectedDate
    $OuDN       = $OuTextBox.Text.Trim()
    $NameFilter = $NameFilterBox.Text.Trim()
    $StatusMode = Get-StatusMode

    Set-Status "⏳  Querying Active Directory ($StatusMode users)…" -Busy $true
    $SearchButton.IsEnabled    = $false
    $ExportCsvButton.IsEnabled = $false
    $GridViewButton.IsEnabled  = $false
    $script:QueryResults       = $null

    try {
        if ($adAvailable) {
            $results = Invoke-ADQuery -Start $StartDate -End $EndDate `
                                       -OuDN $OuDN -NameFilter $NameFilter `
                                       -StatusMode $StatusMode
        } else {
            Start-Sleep -Milliseconds 600
            $results = 1..10 | ForEach-Object {
                $enabled = [bool](Get-Random -Minimum 0 -Maximum 2)
                $wc      = $StartDate.AddDays((Get-Random -Maximum 30))
                [PSCustomObject]@{
                    DisplayName       = "Demo User $_"
                    SamAccountName    = "duser$_"
                    UserPrincipalName = "duser$_@demo.local"
                    EmailAddress      = "duser$_@demo.local"
                    Enabled           = $enabled
                    whenChanged       = $wc
                    DisabledDate      = if (-not $enabled) { $wc.ToString('yyyy-MM-dd') } else { $null }
                    whenCreated       = $StartDate.AddDays(-(Get-Random -Maximum 365))
                    LastLogonDate     = $StartDate.AddDays(-(Get-Random -Maximum 60))
                    Description       = "Simulated account"
                    Department        = @('IT','HR','Finance','Sales')[(Get-Random -Maximum 4)]
                    Title             = "Role $_"
                    Manager           = "Manager $(Get-Random -Maximum 3)"
                    OU                = "OU=Demo"
                }
            }
        }

        $script:QueryResults = $results
        Update-PreviewGrid -Data $results

        if ($results.Count -eq 0) {
            Set-Status "✅  No $StatusMode users found in that date range."
        } else {
            Set-Status "✅  Found $($results.Count) $StatusMode user$(if($results.Count -ne 1){'s'}). Ready to export."
            $ExportCsvButton.IsEnabled = $true
            $GridViewButton.IsEnabled  = $true
        }
    } catch {
        Set-Status "❌  Error: $($_.Exception.Message)"
        Write-Error $_
        [System.Windows.MessageBox]::Show(
            "Query failed:`n`n$($_.Exception.Message)", "AD Query Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
    } finally {
        $SearchButton.IsEnabled = $true
        $ProgressBar.Visibility = 'Collapsed'
    }
})

# ── Export CSV ────────────────────────────────────────────────────────────────
$ExportCsvButton.Add_Click({
    if (-not $script:QueryResults -or $script:QueryResults.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No results to export. Run a search first.", "Nothing to export") | Out-Null
        return
    }

    $selectedCols = @(Get-SelectedColumns)
    if ($selectedCols.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Please select at least one column to export.", "No columns selected") | Out-Null
        return
    }

    $s = [datetime]$StartDatePicker.SelectedDate
    $e = [datetime]$EndDatePicker.SelectedDate
    $StatusMode = Get-StatusMode

    $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
    $saveDialog.Filter   = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $saveDialog.FileName = "ADUsers_{0}_{1:yyyyMMdd}-{2:yyyyMMdd}.csv" -f $StatusMode, $s, $e
    $saveDialog.Title    = "Save Export As"

    if ($saveDialog.ShowDialog() -ne $true) { return }
    $OutputPath = $saveDialog.FileName

    Set-Status "💾  Writing CSV…" -Busy $true

    try {
        $script:QueryResults |
            Select-Object $selectedCols |
            Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

        Set-Status "✅  Exported $($script:QueryResults.Count) records → $OutputPath"

        $open = [System.Windows.MessageBox]::Show(
            "Exported $($script:QueryResults.Count) record(s) to:`n$OutputPath`n`nOpen the file now?",
            "Export Complete",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Information
        )
        if ($open -eq 'Yes') { Start-Process $OutputPath }

    } catch {
        Set-Status "❌  Export failed: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show("Export failed:`n$($_.Exception.Message)", "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error) | Out-Null
    } finally {
        $ProgressBar.Visibility = 'Collapsed'
    }
})

# ── Open in Out-GridView ──────────────────────────────────────────────────────
$GridViewButton.Add_Click({
    if (-not $script:QueryResults) { return }
    $selectedCols = @(Get-SelectedColumns)
    if ($selectedCols.Count -eq 0) { $selectedCols = $ColMap.Keys }

    try {
        $script:QueryResults |
            Select-Object $selectedCols |
            Out-GridView -Title "AD User Export & Audit — $($script:QueryResults.Count) records"
    } catch {
        [System.Windows.MessageBox]::Show("GridView error: $_", "Error") | Out-Null
    }
})

# ── Clear ─────────────────────────────────────────────────────────────────────
$ClearButton.Add_Click({
    $script:QueryResults       = $null
    $ResultsGrid.ItemsSource   = $null
    $ResultsGroup.Header       = "Results Preview  (0 records)"
    $ExportCsvButton.IsEnabled = $false
    $GridViewButton.IsEnabled  = $false
    $OuTextBox.Text            = ''
    $NameFilterBox.Text        = ''
    Set-Status "Cleared — select status, date range and click Search."
})

# ── Keyboard shortcuts ────────────────────────────────────────────────────────
$Window.Add_KeyDown({
    param($s, $e)
    if ($e.Key -eq 'Return' -and $SearchButton.IsEnabled) {
        $SearchButton.RaiseEvent(
            [System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent)
        )
    }
    if ($e.Key -eq 'Escape') { $Window.Close() }
})

# ── Show ──────────────────────────────────────────────────────────────────────
Write-Host "Log: $LogPath"
$null = $Window.ShowDialog()
Stop-Transcript | Out-Null
