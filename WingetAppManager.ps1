#Requires -Version 5.0
<#
.SYNOPSIS
    Winget Application Manager - Complete Fixed Version
.DESCRIPTION
    - Fixed Background property error
    - Import/Export with data grid
    - Improved update detection (catches all packages)
    - Clear status display with current app name
#>

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Settings
$script:settingsPath = "$env:APPDATA\WingetManager\settings.json"
$script:settings = @{ Theme = 'Dark'; RecentFiles = @(); LastExportPath = '' }

function Load-Settings {
    if (Test-Path $script:settingsPath) {
        try {
            $loaded = Get-Content $script:settingsPath -Raw | ConvertFrom-Json
            $script:settings.Theme = $loaded.Theme
            $script:settings.RecentFiles = @($loaded.RecentFiles | Select-Object -First 5)
            $script:settings.LastExportPath = $loaded.LastExportPath
        } catch { }
    }
}

function Save-Settings {
    $dir = Split-Path $script:settingsPath -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $script:settings | ConvertTo-Json | Out-File $script:settingsPath -Encoding UTF8
}

Load-Settings

function Get-CultureDateStamp {
    $culture = [System.Globalization.CultureInfo]::CurrentCulture
    $datePattern = $culture.DateTimeFormat.ShortDatePattern -replace '[/\\:]', '-'
    $stamp = (Get-Date).ToString("$datePattern`_HHmmss", $culture)
    $invalid = [System.IO.Path]::GetInvalidFileNameChars() -join ''
    return $stamp -replace "[$([RegEx]::Escape($invalid))]", '-'
}

function Get-ExportFileName {
    return "WingetAppMgr_$($env:COMPUTERNAME)_$(Get-CultureDateStamp).json"
}

function Write-Log {
    param([string]$Message, [string]$Level = 'Info')
    
    $color = switch ($Level) {
        'Error'   { $script:currentTheme.ErrorColor }
        'Warning' { $script:currentTheme.WarningColor }
        'Success' { $script:currentTheme.SuccessColor }
        'Install' { $script:currentTheme.InstallColor }
        'Info'    { $script:currentTheme.InfoColor }
        default   { $script:currentTheme.InfoColor }
    }

    $time = (Get-Date).ToString('HH:mm:ss')
    $line = "[$time] $Message"

    $para = New-Object System.Windows.Documents.Paragraph
    $para.Margin = New-Object System.Windows.Thickness(0, 2, 0, 2)
    $run = New-Object System.Windows.Documents.Run($line)
    $run.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($color)
    $para.Inlines.Add($run)

    $logViewer.Document.Blocks.Add($para)
    $logViewer.ScrollToEnd()
}

# Themes
$script:themes = @{
    Dark = @{
        WindowBg = '#1E1E1E'; SurfaceBg = '#252526'; AccentPrimary = '#0078D4'
        TextPrimary = '#F0F0F0'; TextHint = '#9E9E9E'; BorderColor = '#3F3F46'
        LogBackground = '#1E1E1E'; LogBorder = '#3F3F46'
        SuccessColor = '#66BB6A'; WarningColor = '#FFCA28'; ErrorColor = '#EF5350'
        InfoColor = '#E0E0E0'; InstallColor = '#00BCD4'
    }
    Light = @{
        WindowBg = '#FFFFFF'; SurfaceBg = '#F5F5F5'; AccentPrimary = '#0078D4'
        TextPrimary = '#1E1E1E'; TextHint = '#757575'; BorderColor = '#D0D0D0'
        LogBackground = '#FFFFFF'; LogBorder = '#D0D0D0'
        SuccessColor = '#1B5E20'; WarningColor = '#BF360C'; ErrorColor = '#B71C1C'
        InfoColor = '#212121'; InstallColor = '#006064'
    }
}

$script:currentTheme = $script:themes[$script:settings.Theme]

function Update-ButtonStyles {
    $t = $script:currentTheme
    foreach ($btn in @($btnExport, $btnImport, $btnBrowse, $btnClear, $btnTheme, $btnRefreshPackages, 
                       $btnUpdateSelected, $btnUninstallSelected, $btnSelectAll, $btnSelectNone, 
                       $btnSearch, $btnInstallSelected, $btnCancel)) {
        if ($btn) {
            $btn.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.AccentPrimary)
            $btn.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#FFFFFF')
            $btn.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.AccentPrimary)
        }
    }
}

function Apply-Theme {
    param([string]$ThemeName)
    $script:currentTheme = $script:themes[$ThemeName]
    $t = $script:currentTheme

    $window.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.WindowBg)
    $mainGrid.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.WindowBg)
    $headerPanel.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.SurfaceBg)
    $contentPanel.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.WindowBg)
    $statusBar.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.SurfaceBg)

    if ($mainTabs) { $mainTabs.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.WindowBg) }
    if ($importExportContent) { $importExportContent.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.WindowBg) }
    if ($packageManagerContent) { $packageManagerContent.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.WindowBg) }
    if ($activityLogContent) { $activityLogContent.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.WindowBg) }
    if ($aboutContent) { $aboutContent.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.WindowBg) }

    $logViewer.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.LogBackground)
    $logViewer.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.TextPrimary)
    $logViewer.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.LogBorder)

    # Apply to both grids
    foreach ($grid in @($packageManagerGrid, $importExportGrid)) {
        if ($grid) {
            $grid.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.WindowBg)
            $grid.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.TextPrimary)
            $grid.AlternatingRowBackground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.SurfaceBg)
            $grid.RowBackground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.WindowBg)

            $rowStyle = New-Object System.Windows.Style([System.Windows.Controls.DataGridRow])
            $fgSetter = New-Object System.Windows.Setter([System.Windows.Controls.DataGridRow]::ForegroundProperty, 
                [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.TextPrimary))
            $rowStyle.Setters.Add($fgSetter)
            $grid.RowStyle = $rowStyle

            $headerStyle = New-Object System.Windows.Style([System.Windows.Controls.Primitives.DataGridColumnHeader])
            $headerStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Primitives.DataGridColumnHeader]::BackgroundProperty,
                [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.SurfaceBg))))
            $headerStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Primitives.DataGridColumnHeader]::ForegroundProperty,
                [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.TextPrimary))))
            $headerStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Primitives.DataGridColumnHeader]::FontWeightProperty,
                [System.Windows.FontWeights]::SemiBold)))
            $headerStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.Primitives.DataGridColumnHeader]::PaddingProperty,
                (New-Object System.Windows.Thickness(8, 6, 8, 6)))))
            $grid.ColumnHeaderStyle = $headerStyle
        }
    }

    if ($txtPackageCount) { $txtPackageCount.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.TextPrimary) }
    if ($txtProgressDetail) { $txtProgressDetail.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.TextPrimary) }
    if ($txtCurrentApp) { $txtCurrentApp.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.TextPrimary) }
    foreach ($label in @($titleLabel, $pathLabel, $statusLabel)) {
        if ($label) { $label.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.TextPrimary) }
    }
    if ($subtitleLabel) { $subtitleLabel.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.TextHint) }
    if ($hintText) { $hintText.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.TextHint) }
    
    # About section theming
    if ($aboutTitle) { $aboutTitle.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.TextPrimary) }
    if ($aboutVersion) { $aboutVersion.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.TextPrimary) }
    if ($aboutDescription) { $aboutDescription.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.TextPrimary) }
    if ($aboutAuthorLabel) { $aboutAuthorLabel.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.TextPrimary) }
    if ($aboutAuthor) { $aboutAuthor.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.TextPrimary) }
    if ($aboutLinksLabel) { $aboutLinksLabel.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.TextPrimary) }
    if ($aboutLink1) { $aboutLink1.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.TextPrimary) }
    if ($aboutLink2) { $aboutLink2.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.TextPrimary) }
    if ($aboutLink3) { $aboutLink3.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.TextPrimary) }
    if ($aboutBorder) { $aboutBorder.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.BorderColor) }
    
    # Hyperlink colors
    foreach ($link in @($hyperlinkGitHub, $hyperlinkWebsite, $hyperlinkLinkedIn)) {
        if ($link) {
            $link.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.AccentPrimary)
        }
    }
    if ($txtFilePath) {
        $txtFilePath.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.WindowBg)
        $txtFilePath.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.TextPrimary)
        $txtFilePath.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.BorderColor)
    }
    if ($txtSearch) {
        $txtSearch.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.WindowBg)
        $txtSearch.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.TextPrimary)
        $txtSearch.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFrom($t.BorderColor)
    }

    Update-ButtonStyles
    $script:settings.Theme = $ThemeName
    Save-Settings
}

function Set-UIBusy {
    param([bool]$Busy, [string]$StatusText = 'Ready', [string]$CurrentApp = '', [string]$ProgressDetail = '')
    
    $btnExport.IsEnabled = -not $Busy
    $btnImport.IsEnabled = -not $Busy
    $btnBrowse.IsEnabled = -not $Busy
    $btnClear.IsEnabled = -not $Busy
    if ($btnRefreshPackages) { $btnRefreshPackages.IsEnabled = -not $Busy }
    if ($btnUpdateSelected) { $btnUpdateSelected.IsEnabled = -not $Busy }
    if ($btnUninstallSelected) { $btnUninstallSelected.IsEnabled = -not $Busy }
    if ($btnSelectAll) { $btnSelectAll.IsEnabled = -not $Busy }
    if ($btnSelectNone) { $btnSelectNone.IsEnabled = -not $Busy }
    if ($btnSearch) { $btnSearch.IsEnabled = -not $Busy }
    if ($btnInstallSelected) { $btnInstallSelected.IsEnabled = -not $Busy }
    
    if ($btnCancel) {
        $btnCancel.Visibility = if ($Busy) { 'Visible' } else { 'Collapsed' }
    }

    $statusLabel.Text = $StatusText
    
    if ($txtCurrentApp) {
        $txtCurrentApp.Text = $CurrentApp
        $txtCurrentApp.Visibility = if ($Busy -and $CurrentApp) { 'Visible' } else { 'Collapsed' }
    }
    
    if ($txtProgressDetail) {
        $txtProgressDetail.Text = $ProgressDetail
        $txtProgressDetail.Visibility = if ($Busy -and $ProgressDetail) { 'Visible' } else { 'Collapsed' }
    }
    
    $progressBar.Visibility = if ($Busy) { 'Visible' } else { 'Collapsed' }
    $progressBar.IsIndeterminate = $Busy
}

$script:_activeOp = $null
$script:cancelRequested = $false

function Start-RunspaceOperation {
    param([hashtable]$SyncHash, [scriptblock]$WorkScript, [hashtable]$OpArgs, [scriptblock]$OnComplete)

    $script:_activeOp = $SyncHash
    $script:_activeOp._onComplete = $OnComplete
    $script:cancelRequested = $false

    $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = [System.Threading.ApartmentState]::STA
    $rs.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('syncHash', $SyncHash)
    $rs.SessionStateProxy.SetVariable('opArgs', $OpArgs)

    $ps = [System.Management.Automation.PowerShell]::Create()
    $ps.Runspace = $rs
    [void] $ps.AddScript($WorkScript)

    $SyncHash._ps = $ps
    $SyncHash._rs = $rs
    $SyncHash._handle = $ps.BeginInvoke()
    $SyncHash._cancelled = $false

    $timer = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [TimeSpan]::FromMilliseconds(250)
    $SyncHash._timer = $timer

    $timer.Add_Tick({
        $op = $script:_activeOp
        if ($null -eq $op) {
            $this.Stop()
            return
        }

        if ($script:cancelRequested -and -not $op._cancelled) {
            $op._cancelled = $true
            Write-Log 'Cancellation requested...' 'Warning'
            try { $op._ps.Stop() } catch { }
        }

        $isFinished = $op._ps.InvocationStateInfo.State -ne [System.Management.Automation.PSInvocationState]::Running

        if ($isFinished) {
            $op._timer.Stop()
            try {
                $op._ps.EndInvoke($op._handle)
            } catch {
                if (-not $op.Error) { $op.Error = $_.Exception.Message }
                if ($op._timer -and $op._timer.IsEnabled) { $op._timer.Stop() }
            } finally {
                $op._rs.Close()
                $op._rs.Dispose()
                $op._ps.Dispose()
            }
            & $op._onComplete
            $script:_activeOp = $null
            $script:cancelRequested = $false
        }
    })

    $timer.Start()
}

# Data storage
$script:installedPackages = @()
$script:searchResults = @()
$script:currentView = 'installed'
$script:importExportPackages = @()
$script:isRefreshing = $false

# XAML
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Winget Application Manager" Height="800" Width="1200"
        MinHeight="700" MinWidth="1000" WindowStartupLocation="CenterScreen" AllowDrop="True">
    <Grid x:Name="MainGrid">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <Border x:Name="HeaderPanel" Grid.Row="0" Padding="24,20" BorderThickness="0,0,0,1">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" VerticalAlignment="Center">
                    <TextBlock x:Name="TitleLabel" Text="Winget Application Manager" FontSize="24" FontWeight="SemiBold" Margin="0,0,0,4"/>
                    <TextBlock x:Name="SubtitleLabel" Text="Manage, search, and install applications" FontSize="13" Opacity="0.85"/>
                </StackPanel>
                <Button x:Name="BtnTheme" Grid.Column="1" Content="🌙 Dark Mode" Width="140" Height="36" FontSize="13" Cursor="Hand"/>
            </Grid>
        </Border>
        
        <Border x:Name="ContentPanel" Grid.Row="1" Padding="0" Margin="20,20,20,20">
            <TabControl x:Name="MainTabs" BorderThickness="0,1,0,0">
                <!-- Import/Export Tab with DataGrid -->
                <TabItem x:Name="TabImportExport" Header="Import / Export" FontSize="13" Padding="12,8">
                    <Grid Margin="0,16,0,0" x:Name="ImportExportContent">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        
                        <StackPanel Grid.Row="0" Margin="0,0,0,20">
                            <TextBlock x:Name="PathLabel" Text="JSON Manifest File" FontSize="14" FontWeight="Medium" Margin="0,0,0,8"/>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBox x:Name="TxtFilePath" Grid.Column="0" Height="40" Padding="12,0" FontSize="13" VerticalContentAlignment="Center" BorderThickness="1"/>
                                <Button x:Name="BtnBrowse" Grid.Column="1" Content="📁 Browse" Width="100" Height="40" FontSize="13" Margin="8,0,0,0" Cursor="Hand"/>
                                <Button x:Name="BtnClear" Grid.Column="2" Content="✕ Clear" Width="90" Height="40" FontSize="13" Margin="8,0,0,0" Cursor="Hand"/>
                            </Grid>
                            <TextBlock x:Name="HintText" Text="💡 Drag and drop a JSON file here" FontSize="12" Opacity="0.85" Margin="0,8,0,0"/>
                        </StackPanel>
                        
                        <Grid Grid.Row="1" Margin="0,0,0,16">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Button x:Name="BtnExport" Grid.Column="0" Content="⬆ Export Applications" Height="50" FontSize="14" FontWeight="SemiBold" Margin="0,0,10,0" Cursor="Hand"/>
                            <Button x:Name="BtnImport" Grid.Column="1" Content="⬇ Import Applications" Height="50" FontSize="14" FontWeight="SemiBold" Margin="10,0,0,0" Cursor="Hand"/>
                        </Grid>
                        
                        <DataGrid x:Name="ImportExportGrid" Grid.Row="2" AutoGenerateColumns="False" CanUserAddRows="False" 
                                  HeadersVisibility="Column" GridLinesVisibility="Horizontal" SelectionMode="Single" IsReadOnly="True"
                                  VerticalScrollBarVisibility="Auto">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="2*" IsReadOnly="True"/>
                                <DataGridTextColumn Header="ID" Binding="{Binding Id}" Width="2*" IsReadOnly="True"/>
                                <DataGridTextColumn Header="Version" Binding="{Binding Version}" Width="*" IsReadOnly="True"/>
                                <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="120" IsReadOnly="True"/>
                            </DataGrid.Columns>
                        </DataGrid>
                    </Grid>
                </TabItem>
                
                <!-- Package Manager Tab -->
                <TabItem x:Name="TabPackageManager" Header="Package Manager" FontSize="13" Padding="12,8">
                    <Grid Margin="0,16,0,0" x:Name="PackageManagerContent">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        
                        <Grid Grid.Row="0" Margin="0,0,0,12">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8">
                                <Button x:Name="BtnRefreshPackages" Content="🔄 Refresh Installed" Width="140" Height="36" Margin="0,0,8,0"/>
                                <Button x:Name="BtnSelectAll" Content="✓ Select All" Width="100" Height="36" Margin="0,0,8,0"/>
                                <Button x:Name="BtnSelectNone" Content="✗ Select None" Width="110" Height="36" Margin="0,0,16,0"/>
                                <TextBlock x:Name="TxtPackageCount" VerticalAlignment="Center" FontSize="13" FontWeight="Medium"/>
                            </StackPanel>
                            <Grid Grid.Row="1">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBox x:Name="TxtSearch" Grid.Column="0" Height="36" Padding="10,0" FontSize="13" 
                                         VerticalContentAlignment="Center" BorderThickness="1"/>
                                <Button x:Name="BtnSearch" Grid.Column="1" Content="🔍 Search / Install" Width="140" Height="36" Margin="8,0,0,0" Cursor="Hand"/>
                            </Grid>
                        </Grid>
                        
                        <DataGrid x:Name="PackageManagerGrid" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False"
                                  HeadersVisibility="Column" GridLinesVisibility="Horizontal" SelectionMode="Extended" IsReadOnly="False"
                                  CanUserSortColumns="True" VerticalScrollBarVisibility="Auto">
                            <DataGrid.Columns>
                                <DataGridCheckBoxColumn Header="☑" Binding="{Binding Selected, UpdateSourceTrigger=PropertyChanged}" Width="50"/>
                                <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="2.5*" IsReadOnly="True"/>
                                <DataGridTextColumn Header="ID" Binding="{Binding Id}" Width="2*" IsReadOnly="True"/>
                                <DataGridTextColumn Header="Installed" Binding="{Binding Version}" Width="*" IsReadOnly="True"/>
                                <DataGridTextColumn Header="Available" Binding="{Binding Available}" Width="*" IsReadOnly="True"/>
                                <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="90" IsReadOnly="True"/>
                                <DataGridTextColumn Header="Source" Binding="{Binding Source}" Width="80" IsReadOnly="True"/>
                            </DataGrid.Columns>
                        </DataGrid>
                        
                        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
                            <Button x:Name="BtnUpdateSelected" Content="⬆ Update Selected" Width="140" Height="44" Margin="0,0,8,0" 
                                    FontSize="14" FontWeight="Bold"/>
                            <Button x:Name="BtnInstallSelected" Content="⬇ Install Selected" Width="140" Height="44" Margin="0,0,8,0"
                                    FontSize="14" FontWeight="Bold"/>
                            <Button x:Name="BtnUninstallSelected" Content="🗑 Uninstall Selected" Width="160" Height="44"
                                    FontSize="13"/>
                        </StackPanel>
                    </Grid>
                </TabItem>
                
                <!-- Activity Log Tab -->
                <TabItem x:Name="TabActivityLog" Header="Activity Log" FontSize="13" Padding="12,8">
                    <Grid Margin="0,16,0,0" x:Name="ActivityLogContent">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <TextBlock Grid.Row="0" Text="Real-time operation log" FontSize="14" Opacity="0.85" Margin="0,0,0,12"/>
                        <Border Grid.Row="1" BorderThickness="1" CornerRadius="4">
                            <RichTextBox x:Name="LogViewer" IsReadOnly="True" VerticalScrollBarVisibility="Auto"
                                         FontFamily="Consolas" FontSize="12" Padding="12" BorderThickness="0">
                                <RichTextBox.Document><FlowDocument/></RichTextBox.Document>
                            </RichTextBox>
                        </Border>
                    </Grid>
                </TabItem>
                
                <!-- About Tab -->
                <TabItem x:Name="TabAbout" Header="About" FontSize="13" Padding="12,8">
                    <ScrollViewer x:Name="AboutContent" VerticalScrollBarVisibility="Auto">
                        <StackPanel Margin="40,40,40,40" MaxWidth="700">
                            <TextBlock x:Name="AboutTitle" Text="WinGet Application Manager" FontSize="32" FontWeight="Bold" Margin="0,0,0,8"/>
                            <TextBlock x:Name="AboutVersion" Text="Version 1.0.0" FontSize="16" Opacity="0.7" Margin="0,0,0,32"/>
                            
                            <TextBlock x:Name="AboutDescription" Text="A modern GUI for managing Windows applications with WinGet."
                                       FontSize="14" TextWrapping="Wrap" Margin="0,0,0,32" LineHeight="22"/>
                            
                            <Border x:Name="AboutBorder" BorderThickness="0,1,0,0" Padding="0,24,0,0" Margin="0,0,0,24">
                                <StackPanel>
                                    <TextBlock x:Name="AboutAuthorLabel" Text="👤 Author" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,16"/>
                                    <TextBlock x:Name="AboutAuthor" Text="Christopher Munn" FontSize="14" Margin="0,0,0,24"/>
                                    
                                    <TextBlock x:Name="AboutLinksLabel" Text="🔗 Links" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,16"/>
                                    <StackPanel Margin="0,0,0,8">
                                        <TextBlock x:Name="AboutLink1" FontSize="14" Margin="0,0,0,8">
                                            <Hyperlink x:Name="HyperlinkGitHub" NavigateUri="https://github.com/ChrisMunnPS/WinGet-Application-Manager">
                                                <Run Text="🐙 GitHub: WinGet-Application-Manager"/>
                                            </Hyperlink>
                                        </TextBlock>
                                        <TextBlock x:Name="AboutLink2" FontSize="14" Margin="0,0,0,8">
                                            <Hyperlink x:Name="HyperlinkWebsite" NavigateUri="https://ChrisMunnPS.github.io">
                                                <Run Text="🌐 Website: ChrisMunnPS.github.io"/>
                                            </Hyperlink>
                                        </TextBlock>
                                        <TextBlock x:Name="AboutLink3" FontSize="14">
                                            <Hyperlink x:Name="HyperlinkLinkedIn" NavigateUri="https://www.linkedin.com/in/chrismunn/">
                                                <Run Text="💼 LinkedIn: Chris Munn"/>
                                            </Hyperlink>
                                        </TextBlock>
                                    </StackPanel>
                                </StackPanel>
                            </Border>
                        </StackPanel>
                    </ScrollViewer>
                </TabItem>
            </TabControl>
        </Border>
        
        <!-- Status Bar -->
        <Border x:Name="StatusBar" Grid.Row="2" Padding="24,12" BorderThickness="0,1,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="StatusLabel" Grid.Column="0" Text="Ready" FontSize="12" VerticalAlignment="Center" Margin="0,0,16,0"/>
                <TextBlock x:Name="TxtCurrentApp" Grid.Column="1" FontSize="12" FontWeight="SemiBold" 
                           VerticalAlignment="Center" Visibility="Collapsed"/>
                <TextBlock x:Name="TxtProgressDetail" Grid.Column="2" FontSize="12" FontWeight="Bold" 
                           VerticalAlignment="Center" Margin="16,0" Visibility="Collapsed"/>
                <Button x:Name="BtnCancel" Grid.Column="3" Content="✕ Cancel" Width="90" Height="28" 
                        FontSize="11" Margin="0,0,16,0" Visibility="Collapsed"/>
                <ProgressBar x:Name="ProgressBar" Grid.Column="4" Width="200" Height="4" Visibility="Collapsed"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

$xmlStream = [System.IO.MemoryStream]::new()
$xmlWriter = [System.IO.StreamWriter]::new($xmlStream)
$xmlWriter.Write($xaml)
$xmlWriter.Flush()
$xmlStream.Seek(0, 'Begin') | Out-Null
$window = [System.Windows.Markup.XamlReader]::Load($xmlStream)
$xmlStream.Close()

# Get controls
$mainGrid = $window.FindName('MainGrid')
$headerPanel = $window.FindName('HeaderPanel')
$contentPanel = $window.FindName('ContentPanel')
$statusBar = $window.FindName('StatusBar')
$titleLabel = $window.FindName('TitleLabel')
$subtitleLabel = $window.FindName('SubtitleLabel')
$pathLabel = $window.FindName('PathLabel')
$hintText = $window.FindName('HintText')
$txtFilePath = $window.FindName('TxtFilePath')
$btnBrowse = $window.FindName('BtnBrowse')
$btnClear = $window.FindName('BtnClear')
$btnExport = $window.FindName('BtnExport')
$btnImport = $window.FindName('BtnImport')
$btnTheme = $window.FindName('BtnTheme')
$logViewer = $window.FindName('LogViewer')
$statusLabel = $window.FindName('StatusLabel')
$txtCurrentApp = $window.FindName('TxtCurrentApp')
$txtProgressDetail = $window.FindName('TxtProgressDetail')
$progressBar = $window.FindName('ProgressBar')
$btnCancel = $window.FindName('BtnCancel')
$mainTabs = $window.FindName('MainTabs')
$importExportContent = $window.FindName('ImportExportContent')
$packageManagerContent = $window.FindName('PackageManagerContent')
$activityLogContent = $window.FindName('ActivityLogContent')
$importExportGrid = $window.FindName('ImportExportGrid')
$packageManagerGrid = $window.FindName('PackageManagerGrid')
$btnRefreshPackages = $window.FindName('BtnRefreshPackages')
$btnSelectAll = $window.FindName('BtnSelectAll')
$btnSelectNone = $window.FindName('BtnSelectNone')
$btnUpdateSelected = $window.FindName('BtnUpdateSelected')
$btnInstallSelected = $window.FindName('BtnInstallSelected')
$btnUninstallSelected = $window.FindName('BtnUninstallSelected')
$txtPackageCount = $window.FindName('TxtPackageCount')
$txtSearch = $window.FindName('TxtSearch')
$btnSearch = $window.FindName('BtnSearch')
$aboutContent = $window.FindName('AboutContent')
$aboutTitle = $window.FindName('AboutTitle')
$aboutVersion = $window.FindName('AboutVersion')
$aboutDescription = $window.FindName('AboutDescription')
$aboutBorder = $window.FindName('AboutBorder')
$aboutAuthorLabel = $window.FindName('AboutAuthorLabel')
$aboutAuthor = $window.FindName('AboutAuthor')
$aboutLinksLabel = $window.FindName('AboutLinksLabel')
$aboutLink1 = $window.FindName('AboutLink1')
$aboutLink2 = $window.FindName('AboutLink2')
$aboutLink3 = $window.FindName('AboutLink3')
$hyperlinkGitHub = $window.FindName('HyperlinkGitHub')
$hyperlinkWebsite = $window.FindName('HyperlinkWebsite')
$hyperlinkLinkedIn = $window.FindName('HyperlinkLinkedIn')

# Event handler for hyperlinks in About tab
$window.AddHandler(
    [System.Windows.Documents.Hyperlink]::RequestNavigateEvent,
    [System.Windows.RoutedEventHandler]{
        param($sender, $e)
        try {
            Start-Process $e.Uri.AbsoluteUri
            $e.Handled = $true
        } catch {
            Write-Log "Failed to open link: $($e.Uri.AbsoluteUri)" 'Error'
        }
    }
)

Apply-Theme -ThemeName $script:settings.Theme
if ($btnTheme) {
    $btnTheme.Content = if ($script:settings.Theme -eq 'Dark') { '☀ Light Mode' } else { '🌙 Dark Mode' }
}

# Helper function to determine winget operation success
function Test-WingetSuccess {
    param(
        [int]$ExitCode,
        [string]$Output
    )
    
    # Exit code 0 = definite success
    if ($ExitCode -eq 0) { return $true }
    
    # Some success codes from winget
    # -1978335189 (0x8A15000B) = No applicable update found (can be considered success)
    # -1978335212 (0x8A150014) = No newer package versions are available
    if ($ExitCode -eq -1978335189 -or $ExitCode -eq -1978335212) { return $true }
    
    # Check output for success indicators
    if ($Output -match 'Successfully installed|Installation succeeded|upgrade succeeded|No newer package|No applicable update') {
        return $true
    }
    
    # Check for "already installed" messages (also success)
    if ($Output -match 'already installed|version.*is already installed') {
        return $true
    }
    
    # Otherwise, it's a failure
    return $false
}

# IMPROVED Refresh - better parsing to catch ALL updates
function Refresh-PackageList {
    # Prevent concurrent refresh operations
    if ($script:isRefreshing) {
        Write-Log 'Refresh already in progress, please wait...' 'Warning'
        return
    }
    
    $script:isRefreshing = $true
    Set-UIBusy -Busy $true -StatusText 'Loading packages...'
    Write-Log 'Refreshing installed packages...' 'Info'
    
    $syncHash = [hashtable]::Synchronized(@{
        Dispatcher = $window.Dispatcher
        Packages = $null
        Error = $null
    })
    
    $refreshScript = {
        try {
            $prevEncoding = $null
            try {
                $prevEncoding = [Console]::OutputEncoding
                [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
            } catch { }
            
            # IMPROVED - Smarter parsing that finds the ID column first
            $upgradeOutput = & winget list --upgrade-available 2>&1
            $upgradeMap = @{}
            $inDataSection = $false
            
            foreach ($line in ($upgradeOutput -split "`r?`n")) {
                # Skip empty lines
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                
                # Detect separator line
                if ($line -match '^-+\s*$') {
                    $inDataSection = $true
                    continue
                }
                
                # Skip header and footer
                if ($line -match '^Name\s+Id\s+Version|upgrade available') { continue }
                
                # Only process data section
                if (-not $inDataSection) { continue }
                
                # SMART PARSING: Find the package ID (contains dot) using regex
                # Format: Name [spaces] ID [spaces] CurrentVer [spaces] AvailVer [spaces] Source
                if ($line -match '(\S+\.\S+)\s+(\S+)\s+(\S+)\s+(\S+)?\s*$') {
                    $packageId = $Matches[1].Trim()         # e.g., EpicGames.EpicGamesLauncher or Microsoft.Edge
                    $currentVer = $Matches[2].Trim()        # Current version
                    $availableVer = $Matches[3].Trim()      # Available version
                    
                    # Validate: ID must contain dot, available version must exist
                    if ($packageId -match '\.' -and $availableVer -and $availableVer -ne 'Source' -and $currentVer -ne $availableVer) {
                        $upgradeMap[$packageId] = $availableVer
                    }
                }
            }
            
            # Get all installed
            $listOutput = & winget list 2>&1
            
            if ($prevEncoding) {
                try { [Console]::OutputEncoding = $prevEncoding } catch { }
            }
            
            $packages = @()
            $inList = $false
            
            foreach ($line in ($listOutput -split "`r?`n")) {
                $line = $line.TrimEnd()
                
                if ($line -match '^-{3,}') {
                    $inList = $true
                    continue
                }
                
                if (-not $inList) { continue }
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                if ($line -match '^Name\s+Id\s+Version|^----------') { continue }
                
                # Parse installed packages
                if ($line -match '^(.+?)\s{2,}([\S]+)\s{2,}([\S]+)\s{2,}(.*)$') {
                    $name = $Matches[1].Trim()
                    $id = $Matches[2].Trim()
                    $version = $Matches[3].Trim()
                    $source = $Matches[4].Trim()
                    
                    $available = if ($upgradeMap.ContainsKey($id)) { $upgradeMap[$id] } else { '' }
                    $status = if ($available) { 'Update' } else { 'Installed' }
                    
                    $packages += [PSCustomObject]@{
                        Selected = ($available -ne '')
                        Name = $name
                        Id = $id
                        Version = $version
                        Available = $available
                        Status = $status
                        Source = $source
                        IsInstalled = $true
                    }
                }
                elseif ($line -match '^(.+?)\s{2,}([\S]+)\s{2,}([\S]+)\s*$') {
                    $name = $Matches[1].Trim()
                    $id = $Matches[2].Trim()
                    $version = $Matches[3].Trim()
                    
                    $available = if ($upgradeMap.ContainsKey($id)) { $upgradeMap[$id] } else { '' }
                    $status = if ($available) { 'Update' } else { 'Installed' }
                    
                    $packages += [PSCustomObject]@{
                        Selected = ($available -ne '')
                        Name = $name
                        Id = $id
                        Version = $version
                        Available = $available
                        Status = $status
                        Source = ''
                        IsInstalled = $true
                    }
                }
            }
            
            $syncHash.Packages = $packages
            $syncHash.UpgradeCount = $upgradeMap.Count
        } catch {
            $syncHash.Error = $_.Exception.Message
        }
    }
    
    $refreshComplete = {
        # ALWAYS reset UI and flag first to prevent hanging
        $script:isRefreshing = $false
        Set-UIBusy -Busy $false
        
        $op = $script:_activeOp
        
        if (-not $op) {
            Write-Log "Refresh completed" 'Info'
            return
        }
        
        if ($op.Error) {
            Write-Log "Refresh failed: $($op.Error)" 'Error'
            return
        }
        
        if ($op.Packages) {
            try {
                # Save current selections before replacing data
                $previousSelections = @{}
                if ($script:installedPackages) {
                    foreach ($pkg in $script:installedPackages) {
                        if ($pkg.Selected) {
                            $previousSelections[$pkg.Id] = $true
                        }
                    }
                }
                
                $script:installedPackages = $op.Packages | Sort-Object -Property Name
                $script:currentView = 'installed'
                
                # Restore previous selections (except packages with updates - those auto-select)
                foreach ($pkg in $script:installedPackages) {
                    if ($pkg.Status -eq 'Update') {
                        $pkg.Selected = $true  # Auto-select packages with updates
                    } elseif ($previousSelections.ContainsKey($pkg.Id)) {
                        $pkg.Selected = $true  # Restore previous selection
                    }
                }
                
                $packageManagerGrid.ItemsSource = $script:installedPackages
                
                $count = $script:installedPackages.Count
                $updatable = @($script:installedPackages | Where-Object { $_.Status -eq 'Update' }).Count
                $selected = @($script:installedPackages | Where-Object { $_.Selected }).Count
                $txtPackageCount.Text = "$count installed | $updatable updates | $selected selected"
                Write-Log "Loaded $count packages" 'Success'
                
                if ($updatable -gt 0) {
                    Write-Log "Found $updatable updates available" 'Success'
                }
                if ($selected -gt 0) {
                    Write-Log "Auto-selected $selected packages with updates" 'Info'
                }
            } catch {
                Write-Log "Error displaying packages: $($_.Exception.Message)" 'Error'
            }
        } else {
            Write-Log "No packages loaded" 'Warning'
        }
    }
    
    Start-RunspaceOperation -SyncHash $syncHash -WorkScript $refreshScript -OpArgs @{} -OnComplete $refreshComplete
}

# Search
function Search-Packages {
    $searchText = $txtSearch.Text.Trim()
    
    if ([string]::IsNullOrWhiteSpace($searchText)) {
        if ($script:currentView -eq 'installed') {
            $packageManagerGrid.ItemsSource = $script:installedPackages
            $count = $script:installedPackages.Count
            $updatable = ($script:installedPackages | Where-Object { $_.Status -eq 'Update' }).Count
            $selected = ($script:installedPackages | Where-Object { $_.Selected }).Count
            $txtPackageCount.Text = "$count installed | $updatable updates | $selected selected"
        } else {
            $packageManagerGrid.ItemsSource = $script:searchResults
            $txtPackageCount.Text = "$($script:searchResults.Count) available"
        }
        return
    }
    
    $installed = $script:installedPackages | Where-Object {
        $_.Name -like "*$searchText*" -or $_.Id -like "*$searchText*"
    }
    
    if ($installed.Count -gt 0) {
        $script:currentView = 'installed'
        $packageManagerGrid.ItemsSource = $installed
        $txtPackageCount.Text = "Found $($installed.Count) installed"
        Write-Log "Found $($installed.Count) installed" 'Success'
        return
    }
    
    Write-Log "Searching winget for '$searchText'..." 'Info'
    Set-UIBusy -Busy $true -StatusText 'Searching...'
    
    $syncHash = [hashtable]::Synchronized(@{
        Dispatcher = $window.Dispatcher
        Packages = $null
        Error = $null
    })
    
    $searchScript = {
        $query = $opArgs.Query
        try {
            $prevEncoding = $null
            try {
                $prevEncoding = [Console]::OutputEncoding
                [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
            } catch { }
            
            $searchOutput = & winget search $query 2>&1
            
            if ($prevEncoding) {
                try { [Console]::OutputEncoding = $prevEncoding } catch { }
            }
            
            $packages = @()
            $inList = $false
            
            foreach ($line in ($searchOutput -split "`r?`n")) {
                $line = $line.TrimEnd()
                
                if ($line -match '^-{3,}') {
                    $inList = $true
                    continue
                }
                
                if (-not $inList) { continue }
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                if ($line -match '^Name\s+Id\s+Version|^----------') { continue }
                
                if ($line -match '^(.+?)\s{2,}([\S]+)\s{2,}([\S]+)\s{2,}(.*)$') {
                    $packages += [PSCustomObject]@{
                        Selected = $false
                        Name = $Matches[1].Trim()
                        Id = $Matches[2].Trim()
                        Version = ''
                        Available = $Matches[3].Trim()
                        Status = 'Available'
                        Source = $Matches[4].Trim()
                        IsInstalled = $false
                    }
                }
            }
            
            $syncHash.Packages = $packages
        } catch {
            $syncHash.Error = $_.Exception.Message
        }
    }
    
    $searchComplete = {
        Set-UIBusy -Busy $false
        $op = $script:_activeOp
        
        if ($op.Error) {
            Write-Log "Search failed: $($op.Error)" 'Error'
        } elseif ($op.Packages -and $op.Packages.Count -gt 0) {
            $script:searchResults = $op.Packages
            $script:currentView = 'search'
            $packageManagerGrid.ItemsSource = $script:searchResults
            $txtPackageCount.Text = "$($op.Packages.Count) available"
            Write-Log "Found $($op.Packages.Count) available" 'Install'
        } else {
            Write-Log "No results" 'Warning'
        }
    }
    
    Start-RunspaceOperation -SyncHash $syncHash -WorkScript $searchScript -OpArgs @{ Query = $searchText } -OnComplete $searchComplete
}

function Update-SelectionCount {
    if ($packageManagerGrid.ItemsSource) {
        $items = @($packageManagerGrid.ItemsSource)
        if ($items.Count -gt 0) {
            $selected = @($items | Where-Object { $_.Selected }).Count
            
            if ($script:currentView -eq 'installed' -and $script:installedPackages) {
                $updatable = @($items | Where-Object { $_.Status -eq 'Update' }).Count
                $txtPackageCount.Text = "$($items.Count) shown | $updatable updates | $selected selected"
            } elseif ($script:currentView -eq 'search') {
                $txtPackageCount.Text = "$($items.Count) shown | $selected selected"
            }
        }
    }
}

$btnRefreshPackages.Add_Click({ Refresh-PackageList })
$btnSearch.Add_Click({ Search-Packages })
$txtSearch.Add_KeyDown({ param($s, $e); if ($e.Key -eq 'Return') { Search-Packages } })

$btnSelectAll.Add_Click({
    if ($packageManagerGrid.ItemsSource) {
        foreach ($item in $packageManagerGrid.ItemsSource) { $item.Selected = $true }
        $packageManagerGrid.Items.Refresh()
        Update-SelectionCount
    }
})

$btnSelectNone.Add_Click({
    if ($packageManagerGrid.ItemsSource) {
        foreach ($item in $packageManagerGrid.ItemsSource) { $item.Selected = $false }
        $packageManagerGrid.Items.Refresh()
        Update-SelectionCount
    }
})

$packageManagerGrid.Add_CellEditEnding({
    $window.Dispatcher.InvokeAsync([Action]{ Update-SelectionCount }, [System.Windows.Threading.DispatcherPriority]::Background)
})

# UPDATE SELECTED with proper status display
$btnUpdateSelected.Add_Click({
    $selected = $packageManagerGrid.ItemsSource | Where-Object { $_.Selected -and $_.IsInstalled -and $_.Status -eq 'Update' }
    
    if (-not $selected -or $selected.Count -eq 0) {
        [System.Windows.MessageBox]::Show('No packages with updates selected.', 'Nothing to Update', 
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    $result = [System.Windows.MessageBox]::Show(
        "Update $($selected.Count) package(s)?",
        'Update Selected',
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    
    if ($result -ne [System.Windows.MessageBoxResult]::Yes) { return }
    
    Write-Log '═══════════════════════════════════════' 'Info'
    Write-Log "Updating $($selected.Count) packages..." 'Info'
    
    $mainTabs.SelectedIndex = 2  # Switch to Activity Log
    
    $syncHash = [hashtable]::Synchronized(@{
        Dispatcher = $window.Dispatcher
        WriteLog = { param($msg, $level) Write-Log $msg $level }
        StatusLabel = $statusLabel
        TxtCurrentApp = $txtCurrentApp
        TxtProgressDetail = $txtProgressDetail
        Result = $null
        Error = $null
    })
    
    $updateScript = {
        $packages = $opArgs.Packages
        $results = @()
        $current = 0
        $total = $packages.Count
        
        # Strict success check - based on official WinGet error codes
        function Test-UpdateSuccess {
            param([int]$Code, [string]$Output)
            
            # Exit code 0 = definite success
            if ($Code -eq 0) { return $true }
            
            # Official WinGet success codes for "already at latest version"
            # -1978335189 = 0x8A15002B = APPINSTALLER_CLI_ERROR_UPDATE_NOT_APPLICABLE
            # -1978335212 = 0x8A150044 = No newer package available
            if ($Code -eq -1978335189 -or $Code -eq -1978335212) {
                return $true  # These ARE success - package is already latest!
            }
            
            # Check output for explicit success messages
            if ($Output -match 'Successfully installed|Installation completed successfully|upgrade succeeded') {
                return $true
            }
            
            # Check for "already latest" messages
            if ($Output -match 'No applicable upgrade found|No newer package versions are available|already installed') {
                return $true
            }
            
            # Everything else is a failure
            return $false
        }
        
        # Helper to try closing an app
        function Close-Application {
            param([string]$PackageName)
            
            # Extract likely process name from package name
            $processName = $PackageName -replace '[^a-zA-Z0-9]', ''
            
            # Try to find and close processes
            $closed = $false
            $processes = Get-Process | Where-Object { 
                $_.ProcessName -like "*$processName*" -or
                $_.MainWindowTitle -like "*$PackageName*"
            }
            
            if ($processes) {
                foreach ($proc in $processes) {
                    try {
                        $syncHash._closeMsg = "Closing $($proc.ProcessName)..."
                        $syncHash.Dispatcher.Invoke([Action]{
                            $syncHash.WriteLog.Invoke("  $($syncHash._closeMsg)", 'Info')
                        })
                        
                        $proc.CloseMainWindow() | Out-Null
                        Start-Sleep -Milliseconds 500
                        
                        if (-not $proc.HasExited) {
                            $proc.Kill()
                            $proc.WaitForExit(2000)
                        }
                        $closed = $true
                    } catch {
                        # Couldn't close, continue anyway
                    }
                }
            }
            
            return $closed
        }
        
        try {
            $prevEncoding = $null
            try {
                $prevEncoding = [Console]::OutputEncoding
                [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
            } catch { }
            
            foreach ($pkg in $packages) {
                if ($syncHash._cancelled) {
                    $syncHash.Dispatcher.Invoke([Action]{ $syncHash.WriteLog.Invoke("Cancelled", 'Warning') })
                    break
                }
                
                $current++
                $syncHash._currentApp = $pkg.Name
                $syncHash._progress = "$current/$total"
                
                $syncHash.Dispatcher.Invoke([Action]{
                    $syncHash.StatusLabel.Text = "Updating..."
                    $syncHash.TxtCurrentApp.Text = "Current: $($syncHash._currentApp)"
                    $syncHash.TxtCurrentApp.Visibility = 'Visible'
                    $syncHash.TxtProgressDetail.Text = $syncHash._progress
                    $syncHash.TxtProgressDetail.Visibility = 'Visible'
                    $syncHash.WriteLog.Invoke("[$($syncHash._progress)] Updating $($syncHash._currentApp)...", 'Info')
                })
                
                # Try to close the application first
                $wasClosed = Close-Application -PackageName $pkg.Name
                if ($wasClosed) {
                    Start-Sleep -Seconds 2  # Give it time to fully close
                }
                
                # Update with proper flags including --silent to prevent GUI popups
                $out = & winget upgrade --id $pkg.Id --silent --accept-source-agreements --accept-package-agreements 2>&1
                $exitCode = $LASTEXITCODE
                $outputText = ($out | Out-String).Trim()
                
                $isSuccess = Test-UpdateSuccess -Code $exitCode -Output $outputText
                
                if ($isSuccess) {
                    # Check if it was "already latest" vs "actually updated"
                    if ($exitCode -eq -1978335189) {
                        # UPDATE_NOT_APPLICABLE - already at latest
                        $syncHash.Dispatcher.Invoke([Action]{
                            $syncHash.WriteLog.Invoke("  ℹ Already at latest version", 'Info')
                        })
                    } elseif ($outputText -match 'No applicable|No newer|already') {
                        # Message indicates already latest
                        $syncHash.Dispatcher.Invoke([Action]{
                            $syncHash.WriteLog.Invoke("  ℹ Already at latest version", 'Info')
                        })
                    } else {
                        # Actually updated
                        $syncHash.Dispatcher.Invoke([Action]{
                            $syncHash.WriteLog.Invoke("  ✓ Updated successfully", 'Success')
                        })
                    }
                } else {
                    # Log detailed failure
                    $syncHash._failMsg = "Exit code: $exitCode"
                    
                    # Extract actual error from output
                    if ($outputText -match 'error.*') {
                        $syncHash._failMsg += " - $($Matches[0])"
                    } elseif ($outputText -match 'failed.*') {
                        $syncHash._failMsg += " - $($Matches[0])"
                    } elseif ($outputText -match 'requires.*administrator') {
                        $syncHash._failMsg += " - Requires administrator privileges"
                    } elseif ($outputText -match 'in use|running') {
                        $syncHash._failMsg += " - Application is in use"
                    }
                    
                    $syncHash.Dispatcher.Invoke([Action]{
                        $syncHash.WriteLog.Invoke("  ✗ Failed: $($syncHash._failMsg)", 'Error')
                    })
                }
                
                $results += [PSCustomObject]@{
                    Package = $pkg.Name
                    Success = $isSuccess
                }
            }
            
            if ($prevEncoding) {
                try { [Console]::OutputEncoding = $prevEncoding } catch { }
            }
            
            $syncHash.Result = [PSCustomObject]@{ Results = $results; Cancelled = $syncHash._cancelled }
        } catch {
            $syncHash.Error = $_.Exception.Message
        }
    }
    
    $updateComplete = {
        # ALWAYS reset UI first (prevents hanging)
        Set-UIBusy -Busy $false
        
        $op = $script:_activeOp
        
        if (-not $op) {
            Write-Log "Operation completed" 'Info'
            return
        }
        
        if ($op.Error) {
            Write-Log "Error: $($op.Error)" 'Error'
        } elseif ($op._cancelled -or ($op.Result -and $op.Result.Cancelled)) {
            Write-Log "Operation cancelled" 'Warning'
        } else {
            if ($op.Result -and $op.Result.Results) {
                $successful = @($op.Result.Results | Where-Object { $_.Success }).Count
                $failed = $op.Result.Results.Count - $successful
                
                if ($successful -gt 0 -and $failed -eq 0) {
                    Write-Log "✓ All $successful package(s) updated successfully" 'Success'
                } elseif ($successful -gt 0 -and $failed -gt 0) {
                    Write-Log "⚠ Partial success: $successful updated, $failed failed" 'Warning'
                } else {
                    Write-Log "✗ All updates failed" 'Error'
                }
                
                # List failed packages with details
                if ($failed -gt 0) {
                    $failures = $op.Result.Results | Where-Object { -not $_.Success }
                    Write-Log "" 'Info'
                    Write-Log "Failed packages:" 'Error'
                    foreach ($f in $failures) {
                        Write-Log "  • $($f.Package)" 'Error'
                    }
                    Write-Log "" 'Info'
                    Write-Log "💡 Troubleshooting tips:" 'Info'
                    Write-Log "  1. Close the applications completely" 'Info'
                    Write-Log "  2. Try running PowerShell as Administrator" 'Info'
                    Write-Log "  3. Check if apps are running in system tray" 'Info'
                }
            }
            
            # Don't auto-refresh - let user manually refresh to verify
            Write-Log "Click 'Refresh Installed' to verify the updates" 'Info'
        }
        Write-Log '═══════════════════════════════════════' 'Info'
    }
    
    Start-RunspaceOperation -SyncHash $syncHash -WorkScript $updateScript -OpArgs @{ Packages = $selected } -OnComplete $updateComplete
})

# INSTALL SELECTED
$btnInstallSelected.Add_Click({
    $selected = $packageManagerGrid.ItemsSource | Where-Object { $_.Selected -and -not $_.IsInstalled }
    
    if (-not $selected -or $selected.Count -eq 0) {
        [System.Windows.MessageBox]::Show('No packages selected.', 'Nothing to Install', 
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    $result = [System.Windows.MessageBox]::Show(
        "Install $($selected.Count) package(s)?",
        'Install Selected',
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    
    if ($result -ne [System.Windows.MessageBoxResult]::Yes) { return }
    
    Write-Log '═══════════════════════════════════════' 'Info'
    Write-Log "Installing $($selected.Count) packages..." 'Install'
    
    $mainTabs.SelectedIndex = 2
    
    $syncHash = [hashtable]::Synchronized(@{
        Dispatcher = $window.Dispatcher
        WriteLog = { param($msg, $level) Write-Log $msg $level }
        StatusLabel = $statusLabel
        TxtCurrentApp = $txtCurrentApp
        TxtProgressDetail = $txtProgressDetail
        Result = $null
        Error = $null
    })
    
    $installScript = {
        $packages = $opArgs.Packages
        $results = @()
        $current = 0
        $total = $packages.Count
        
        # Helper to check winget success
        function Test-Success {
            param([int]$Code, [string]$Output)
            if ($Code -eq 0) { return $true }
            if ($Code -eq -1978335189 -or $Code -eq -1978335212) { return $true }
            if ($Output -match 'Successfully|succeeded|already installed') { return $true }
            return $false
        }
        
        try {
            $prevEncoding = $null
            try {
                $prevEncoding = [Console]::OutputEncoding
                [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
            } catch { }
            
            foreach ($pkg in $packages) {
                if ($syncHash._cancelled) {
                    $syncHash.Dispatcher.Invoke([Action]{ $syncHash.WriteLog.Invoke("Cancelled", 'Warning') })
                    break
                }
                
                $current++
                $syncHash._currentApp = $pkg.Name
                $syncHash._progress = "$current/$total"
                
                $syncHash.Dispatcher.Invoke([Action]{
                    $syncHash.StatusLabel.Text = "Installing..."
                    $syncHash.TxtCurrentApp.Text = "Current: $($syncHash._currentApp)"
                    $syncHash.TxtCurrentApp.Visibility = 'Visible'
                    $syncHash.TxtProgressDetail.Text = $syncHash._progress
                    $syncHash.TxtProgressDetail.Visibility = 'Visible'
                    $syncHash.WriteLog.Invoke("[$($syncHash._progress)] Installing $($syncHash._currentApp)...", 'Install')
                })
                
                # Install with proper flags including --silent to prevent GUI popups
                $out = & winget install --id $pkg.Id --silent --accept-source-agreements --accept-package-agreements 2>&1
                $exitCode = $LASTEXITCODE
                $outputText = $out -join "`n"
                
                $isSuccess = Test-Success -Code $exitCode -Output $outputText
                
                if (-not $isSuccess) {
                    $syncHash._failMsg = "Exit code: $exitCode"
                    if ($outputText -match '(error|failed|not found)') {
                        $syncHash._failMsg += " - $($Matches[0])"
                    }
                    $syncHash.Dispatcher.Invoke([Action]{
                        $syncHash.WriteLog.Invoke("  Failed: $($syncHash._failMsg)", 'Warning')
                    })
                }
                
                $results += [PSCustomObject]@{
                    Package = $pkg.Name
                    Success = $isSuccess
                }
            }
            
            if ($prevEncoding) {
                try { [Console]::OutputEncoding = $prevEncoding } catch { }
            }
            
            $syncHash.Result = [PSCustomObject]@{ Results = $results; Cancelled = $syncHash._cancelled }
        } catch {
            $syncHash.Error = $_.Exception.Message
        }
    }
    
    $installComplete = {
        Set-UIBusy -Busy $false
        $op = $script:_activeOp
        
        if (-not $op) { return }
        
        if ($op.Error) {
            Write-Log "Error: $($op.Error)" 'Error'
        } elseif ($op._cancelled -or ($op.Result -and $op.Result.Cancelled)) {
            Write-Log "Cancelled" 'Warning'
        } else {
            if ($op.Result -and $op.Result.Results) {
                $successful = @($op.Result.Results | Where-Object { $_.Success }).Count
                $failed = $op.Result.Results.Count - $successful
                
                if ($failed -gt 0) {
                    Write-Log "✓ Install complete: $successful succeeded, $failed failed" 'Warning'
                    $failures = $op.Result.Results | Where-Object { -not $_.Success }
                    foreach ($f in $failures) {
                        Write-Log "  ✗ Failed: $($f.Package)" 'Error'
                    }
                } else {
                    Write-Log "✓ Install complete: All $successful packages installed" 'Success'
                }
            }
            
            # Don't auto-refresh
            if ($successful -gt 0) {
                Write-Log "Click 'Refresh Installed' to see installed packages" 'Info'
            }
        }
        Write-Log '═══════════════════════════════════════' 'Info'
    }
    
    Start-RunspaceOperation -SyncHash $syncHash -WorkScript $installScript -OpArgs @{ Packages = $selected } -OnComplete $installComplete
})

# UNINSTALL SELECTED
$btnUninstallSelected.Add_Click({
    $selected = $packageManagerGrid.ItemsSource | Where-Object { $_.Selected -and $_.IsInstalled }
    
    if (-not $selected -or $selected.Count -eq 0) {
        [System.Windows.MessageBox]::Show('No packages selected.', 'Nothing to Uninstall', 
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    $result = [System.Windows.MessageBox]::Show(
        "Uninstall $($selected.Count) package(s)?`n`nCannot be undone.",
        'Confirm Uninstall',
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )
    
    if ($result -ne [System.Windows.MessageBoxResult]::Yes) { return }
    
    Write-Log '═══════════════════════════════════════' 'Info'
    Write-Log "Uninstalling $($selected.Count) packages..." 'Warning'
    
    $mainTabs.SelectedIndex = 2
    
    $syncHash = [hashtable]::Synchronized(@{
        Dispatcher = $window.Dispatcher
        WriteLog = { param($msg, $level) Write-Log $msg $level }
        StatusLabel = $statusLabel
        TxtCurrentApp = $txtCurrentApp
        TxtProgressDetail = $txtProgressDetail
        Result = $null
        Error = $null
    })
    
    $uninstallScript = {
        $packages = $opArgs.Packages
        $results = @()
        $current = 0
        $total = $packages.Count
        
        # Helper to check winget success
        function Test-Success {
            param([int]$Code, [string]$Output)
            if ($Code -eq 0) { return $true }
            if ($Output -match 'Successfully|succeeded|uninstalled') { return $true }
            return $false
        }
        
        try {
            $prevEncoding = $null
            try {
                $prevEncoding = [Console]::OutputEncoding
                [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
            } catch { }
            
            foreach ($pkg in $packages) {
                if ($syncHash._cancelled) {
                    $syncHash.Dispatcher.Invoke([Action]{ $syncHash.WriteLog.Invoke("Cancelled", 'Warning') })
                    break
                }
                
                $current++
                $syncHash._currentApp = $pkg.Name
                $syncHash._progress = "$current/$total"
                
                $syncHash.Dispatcher.Invoke([Action]{
                    $syncHash.StatusLabel.Text = "Uninstalling..."
                    $syncHash.TxtCurrentApp.Text = "Current: $($syncHash._currentApp)"
                    $syncHash.TxtCurrentApp.Visibility = 'Visible'
                    $syncHash.TxtProgressDetail.Text = $syncHash._progress
                    $syncHash.TxtProgressDetail.Visibility = 'Visible'
                    $syncHash.WriteLog.Invoke("[$($syncHash._progress)] Uninstalling $($syncHash._currentApp)...", 'Warning')
                })
                
                # Uninstall with proper flags including --silent
                $out = & winget uninstall --id $pkg.Id --silent --accept-source-agreements 2>&1
                $exitCode = $LASTEXITCODE
                $outputText = $out -join "`n"
                
                $isSuccess = Test-Success -Code $exitCode -Output $outputText
                
                if (-not $isSuccess) {
                    $syncHash._failMsg = "Exit code: $exitCode"
                    if ($outputText -match '(error|failed|not found)') {
                        $syncHash._failMsg += " - $($Matches[0])"
                    }
                    $syncHash.Dispatcher.Invoke([Action]{
                        $syncHash.WriteLog.Invoke("  Failed: $($syncHash._failMsg)", 'Warning')
                    })
                }
                
                $results += [PSCustomObject]@{
                    Package = $pkg.Name
                    Success = $isSuccess
                }
            }
            
            if ($prevEncoding) {
                try { [Console]::OutputEncoding = $prevEncoding } catch { }
            }
            
            $syncHash.Result = [PSCustomObject]@{ Results = $results; Cancelled = $syncHash._cancelled }
        } catch {
            $syncHash.Error = $_.Exception.Message
        }
    }
    
    $uninstallComplete = {
        Set-UIBusy -Busy $false
        $op = $script:_activeOp
        
        if (-not $op) { return }
        
        if ($op.Error) {
            Write-Log "Error: $($op.Error)" 'Error'
        } elseif ($op._cancelled -or ($op.Result -and $op.Result.Cancelled)) {
            Write-Log "Cancelled" 'Warning'
        } else {
            if ($op.Result -and $op.Result.Results) {
                $successful = @($op.Result.Results | Where-Object { $_.Success }).Count
                $failed = $op.Result.Results.Count - $successful
                
                if ($failed -gt 0) {
                    Write-Log "✓ Uninstall complete: $successful succeeded, $failed failed" 'Warning'
                    $failures = $op.Result.Results | Where-Object { -not $_.Success }
                    foreach ($f in $failures) {
                        Write-Log "  ✗ Failed: $($f.Package)" 'Error'
                    }
                } else {
                    Write-Log "✓ Uninstall complete: All $successful packages uninstalled" 'Success'
                }
            }
            
            # Don't auto-refresh
            if ($successful -gt 0) {
                Write-Log "Click 'Refresh Installed' to update the list" 'Info'
            }
        }
        Write-Log '═══════════════════════════════════════' 'Info'
    }
    
    Start-RunspaceOperation -SyncHash $syncHash -WorkScript $uninstallScript -OpArgs @{ Packages = $selected } -OnComplete $uninstallComplete
})

# CANCEL
$btnCancel.Add_Click({
    $script:cancelRequested = $true
    $btnCancel.IsEnabled = $false
    $btnCancel.Content = "Cancelling..."
    Write-Log 'Stopping after current package...' 'Warning'
})

# Auto-load
$script:packagesLoaded = $false
$mainTabs.Add_SelectionChanged({
    if ($mainTabs.SelectedIndex -eq 1 -and -not $script:packagesLoaded) {
        $script:packagesLoaded = $true
        Refresh-PackageList
    }
})

# Theme
$btnTheme.Add_Click({
    $newTheme = if ($script:settings.Theme -eq 'Dark') { 'Light' } else { 'Dark' }
    Apply-Theme -ThemeName $newTheme
    $btnTheme.Content = if ($script:settings.Theme -eq 'Dark') { '☀ Light Mode' } else { '🌙 Dark Mode' }
})

$btnClear.Add_Click({ $txtFilePath.Text = ''; Write-Log 'File path cleared' 'Info' })

$btnBrowse.Add_Click({
    $currentText = $txtFilePath.Text.Trim()

    if ($currentText -and (Test-Path $currentText -PathType Leaf)) {
        # File exists - show open dialog
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Title = 'Select JSON manifest to import'
    } else {
        # No file or doesn't exist - show save dialog
        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Title = 'Choose export location'
        $dialog.FileName = Get-ExportFileName
    }

    $dialog.Filter = 'JSON files (*.json)|*.json|All files (*.*)|*.*'
    $dialog.FilterIndex = 1
    
    if ($script:settings.LastExportPath) {
        $dialog.InitialDirectory = Split-Path $script:settings.LastExportPath -Parent
    } else {
        $dialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
    }

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtFilePath.Text = $dialog.FileName
        Write-Log "Selected file: $($dialog.FileName)" 'Info'
    }
})

# Drag and drop support
$window.Add_DragEnter({
    param($sender, $e)
    if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
        $e.Effects = [System.Windows.DragDropEffects]::Copy
    } else {
        $e.Effects = [System.Windows.DragDropEffects]::None
    }
    $e.Handled = $true
})

$window.Add_Drop({
    param($sender, $e)
    if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
        $files = $e.Data.GetData([System.Windows.DataFormats]::FileDrop)
        $firstFile = $files[0]
        if ($firstFile.EndsWith('.json', [System.StringComparison]::OrdinalIgnoreCase)) {
            $txtFilePath.Text = $firstFile
            Write-Log "Dropped file: $firstFile" 'Info'
        } else {
            Write-Log 'Only JSON files are supported' 'Warning'
        }
    }
    $e.Handled = $true
})
$btnExport.Add_Click({
    $filePath = $txtFilePath.Text.Trim()

    if (-not $filePath) {
        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Title = 'Choose export location'
        $dialog.Filter = 'JSON files (*.json)|*.json|All files (*.*)|*.*'
        $dialog.FilterIndex = 1
        $dialog.FileName = Get-ExportFileName
        
        if ($script:settings.LastExportPath) {
            $dialog.InitialDirectory = Split-Path $script:settings.LastExportPath -Parent
        } else {
            $dialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
        }
        
        if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            return
        }
        $filePath = $dialog.FileName
        $txtFilePath.Text = $filePath
    }

    if (-not $filePath.EndsWith('.json', [System.StringComparison]::OrdinalIgnoreCase)) {
        $filePath += '.json'
        $txtFilePath.Text = $filePath
    }

    $script:settings.LastExportPath = $filePath
    Save-Settings

    Write-Log '═══════════════════════════════════════' 'Info'
    Write-Log 'Starting export operation...' 'Info'
    Set-UIBusy -Busy $true -StatusText 'Exporting...'

    $syncHash = [hashtable]::Synchronized(@{
        Dispatcher = $window.Dispatcher
        WriteLog = { param($msg, $level) Write-Log $msg $level }
        ImportExportGrid = $importExportGrid
        Result = $null
        Error = $null
    })

    $exportScript = {
        $filePath = $opArgs.FilePath
        
        try {
            $prevEncoding = $null
            try {
                $prevEncoding = [Console]::OutputEncoding
                [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
            } catch { }
            
            # Run winget export
            $out = & winget export $filePath --accept-source-agreements 2>&1
            $exportExitCode = $LASTEXITCODE
            
            Start-Sleep -Milliseconds 500
            
            if (-not (Test-Path $filePath)) {
                throw "Export failed. Exit code: $exportExitCode"
            }
            
            $jsonContent = Get-Content -Path $filePath -Raw -Encoding UTF8
            if ([string]::IsNullOrWhiteSpace($jsonContent)) {
                throw 'Export file is empty.'
            }
            
            $json = $jsonContent | ConvertFrom-Json
            
            # Extract packages
            if ($json.Packages) {
                $packages = $json.Packages
            } elseif ($json.Sources -and $json.Sources[0].Packages) {
                $packages = $json.Sources[0].Packages
            } else {
                throw 'No packages found in export.'
            }
            
            if ($packages.Count -eq 0) {
                throw 'No packages exported.'
            }
            
            # Update grid with package status
            $gridData = @()
            foreach ($pkg in $packages) {
                $gridData += [PSCustomObject]@{
                    Name = if ($pkg.Name) { $pkg.Name } else { $pkg.PackageIdentifier }
                    Id = $pkg.PackageIdentifier
                    Version = $pkg.Version
                    Status = 'Exported ✓'
                }
                
                $syncHash._msg = "Exported: $(if ($pkg.Name) { $pkg.Name } else { $pkg.PackageIdentifier })"
                $syncHash.Dispatcher.Invoke([Action]{
                    $syncHash.WriteLog.Invoke($syncHash._msg, 'Success')
                })
            }
            
            # Update grid
            $syncHash.Dispatcher.Invoke([Action]{
                $syncHash.ImportExportGrid.ItemsSource = $gridData
            })
            
            if ($prevEncoding) {
                try { [Console]::OutputEncoding = $prevEncoding } catch { }
            }
            
            $syncHash.Result = [PSCustomObject]@{
                Success = $true
                FilePath = $filePath
                Count = $packages.Count
            }
        } catch {
            $syncHash.Error = $_.Exception.Message
        }
    }

    $exportComplete = {
        Set-UIBusy -Busy $false
        $op = $script:_activeOp
        
        if ($op.Error) {
            Write-Log "Export failed: $($op.Error)" 'Error'
        } elseif ($op.Result -and $op.Result.Success) {
            Write-Log "✓ Export successful: $($op.Result.Count) packages" 'Success'
            Write-Log "File: $($op.Result.FilePath)" 'Info'
        }
        Write-Log '═══════════════════════════════════════' 'Info'
    }

    Start-RunspaceOperation -SyncHash $syncHash -WorkScript $exportScript -OpArgs @{ FilePath = $filePath } -OnComplete $exportComplete
})
$btnImport.Add_Click({
    $filePath = $txtFilePath.Text.Trim()

    if (-not $filePath) {
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Title = 'Select JSON manifest to import'
        $dialog.Filter = 'JSON files (*.json)|*.json|All files (*.*)|*.*'
        $dialog.FilterIndex = 1
        
        if ($script:settings.LastExportPath) {
            $dialog.InitialDirectory = Split-Path $script:settings.LastExportPath -Parent
        } else {
            $dialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
        }
        
        if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            return
        }
        $filePath = $dialog.FileName
        $txtFilePath.Text = $filePath
    }

    if (-not (Test-Path $filePath)) {
        [System.Windows.MessageBox]::Show(
            "File not found: $filePath",
            'File Not Found',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return
    }

    Write-Log '═══════════════════════════════════════' 'Info'
    Write-Log 'Starting import operation...' 'Info'
    
    # Load and preview packages
    try {
        $jsonContent = Get-Content -Path $filePath -Raw -Encoding UTF8
        $json = $jsonContent | ConvertFrom-Json
        
        if ($json.Packages) {
            $packages = $json.Packages
        } elseif ($json.Sources -and $json.Sources[0].Packages) {
            $packages = $json.Sources[0].Packages
        } else {
            throw 'No packages found in file.'
        }
        
        $result = [System.Windows.MessageBox]::Show(
            "Import $($packages.Count) package(s) from this file?",
            'Confirm Import',
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        
        if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
            return
        }
        
    } catch {
        [System.Windows.MessageBox]::Show(
            "Failed to read file: $($_.Exception.Message)",
            'Import Error',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return
    }

    Set-UIBusy -Busy $true -StatusText 'Importing...'
    
    # Switch to Import/Export tab to show progress
    $mainTabs.SelectedIndex = 0

    $syncHash = [hashtable]::Synchronized(@{
        Dispatcher = $window.Dispatcher
        WriteLog = { param($msg, $level) Write-Log $msg $level }
        StatusLabel = $statusLabel
        TxtCurrentApp = $txtCurrentApp
        TxtProgressDetail = $txtProgressDetail
        ImportExportGrid = $importExportGrid
        Result = $null
        Error = $null
    })

    $importScript = {
        $filePath = $opArgs.FilePath
        $results = @()
        
        try {
            $prevEncoding = $null
            try {
                $prevEncoding = [Console]::OutputEncoding
                [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
            } catch { }
            
            # Load packages
            $jsonContent = Get-Content -Path $filePath -Raw -Encoding UTF8
            $json = $jsonContent | ConvertFrom-Json
            
            if ($json.Packages) {
                $packages = $json.Packages
            } elseif ($json.Sources -and $json.Sources[0].Packages) {
                $packages = $json.Sources[0].Packages
            } else {
                throw 'No packages found.'
            }
            
            $total = $packages.Count
            $current = 0
            $gridData = @()
            
            foreach ($pkg in $packages) {
                if ($syncHash._cancelled) {
                    $syncHash.Dispatcher.Invoke([Action]{ $syncHash.WriteLog.Invoke("Cancelled", 'Warning') })
                    break
                }
                
                $current++
                $pkgName = if ($pkg.Name) { $pkg.Name } else { $pkg.PackageIdentifier }
                
                # Update grid - Installing
                $gridData += [PSCustomObject]@{
                    Name = $pkgName
                    Id = $pkg.PackageIdentifier
                    Version = $pkg.Version
                    Status = 'Installing...'
                }
                $syncHash.Dispatcher.Invoke([Action]{
                    $syncHash.ImportExportGrid.ItemsSource = $gridData
                })
                
                # Update status
                $syncHash._currentApp = $pkgName
                $syncHash._progress = "$current/$total"
                $syncHash.Dispatcher.Invoke([Action]{
                    $syncHash.StatusLabel.Text = "Installing..."
                    $syncHash.TxtCurrentApp.Text = "Current: $($syncHash._currentApp)"
                    $syncHash.TxtCurrentApp.Visibility = 'Visible'
                    $syncHash.TxtProgressDetail.Text = $syncHash._progress
                    $syncHash.TxtProgressDetail.Visibility = 'Visible'
                    $syncHash.WriteLog.Invoke("[$($syncHash._progress)] Installing $($syncHash._currentApp)...", 'Install')
                })
                
                # Install package
                # Install with proper flags including --silent
                $out = & winget install --id $pkg.PackageIdentifier --silent --accept-source-agreements --accept-package-agreements 2>&1
                $exitCode = $LASTEXITCODE
                
                $success = ($exitCode -eq 0)
                
                # Update grid - result
                $gridData[-1].Status = if ($success) { 'Installed ✓' } else { 'Failed ✗' }
                $syncHash.Dispatcher.Invoke([Action]{
                    $syncHash.ImportExportGrid.ItemsSource = $gridData
                })
                
                $results += [PSCustomObject]@{
                    Package = $pkgName
                    Success = $success
                }
            }
            
            if ($prevEncoding) {
                try { [Console]::OutputEncoding = $prevEncoding } catch { }
            }
            
            $syncHash.Result = [PSCustomObject]@{ Results = $results; Cancelled = $syncHash._cancelled }
        } catch {
            $syncHash.Error = $_.Exception.Message
        }
    }

    $importComplete = {
        Set-UIBusy -Busy $false
        $op = $script:_activeOp
        
        if ($op.Error) {
            Write-Log "Import failed: $($op.Error)" 'Error'
        } elseif ($op.Result.Cancelled) {
            Write-Log "Import cancelled" 'Warning'
        } else {
            $successful = ($op.Result.Results | Where-Object { $_.Success }).Count
            $failed = $op.Result.Results.Count - $successful
            Write-Log "✓ Import complete: $successful succeeded, $failed failed" 'Success'
        }
        Write-Log '═══════════════════════════════════════' 'Info'
    }

    Start-RunspaceOperation -SyncHash $syncHash -WorkScript $importScript -OpArgs @{ FilePath = $filePath } -OnComplete $importComplete
})

Write-Log '═══════════════════════════════════════' 'Info'
Write-Log 'Winget Application Manager' 'Success'
Write-Log 'All improvements applied' 'Info'
Write-Log '═══════════════════════════════════════' 'Info'

[void] $window.ShowDialog()