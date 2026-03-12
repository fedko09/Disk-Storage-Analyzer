<#
Directory Explorer - Complete Build v9
- Responsive WPF GUI for Windows 10/11
- Async directory scans via hidden external PowerShell worker process
- Busy overlay with live scan counts and cancel button
- Recursive scans, network path support, mapped network drives, quick paths, profiles
- Hide hidden/system by default, size filters, text filters
- Export CSV, copy paths, delete selected, open and navigate folders
#>

if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    powershell.exe -STA -ExecutionPolicy Bypass -File $PSCommandPath @args
    exit
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -AssemblyName System.Windows.Forms

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Directory Explorer - PowerShell"
        Height="680"
        Width="1180"
        MinHeight="560"
        MinWidth="980"
        WindowStartupLocation="CenterScreen"
        Background="White"
        Foreground="Black">

  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <DockPanel Grid.Row="0" Margin="0,0,0,8">
      <TextBlock Text="Path:" VerticalAlignment="Center" Margin="0,0,6,0"/>
      <TextBox x:Name="txtPath"
               Width="600"
               Margin="0,0,6,0"
               VerticalAlignment="Center"
               ToolTip="Type or paste a folder path, then click Load or press Enter."/>
      <Button x:Name="btnBrowse"
              Content="Browse..."
              Margin="0,0,6,0"
              Padding="10,2"
              ToolTip="Browse for a folder, including mapped network locations when available."/>
      <Button x:Name="btnUp"
              Content="Up"
              Margin="0,0,6,0"
              Padding="10,2"
              ToolTip="Go to the parent folder of the current directory."/>
      <Button x:Name="btnLoad"
              Content="Load"
              Margin="0,0,6,0"
              Padding="12,2"
              ToolTip="Scan the selected directory."/>
      <Button x:Name="btnOpenExplorer"
              Content="Explorer"
              Padding="10,2"
              ToolTip="Open the current folder in File Explorer."/>
    </DockPanel>

    <DockPanel Grid.Row="1" Margin="0,0,0,8">
      <CheckBox x:Name="chkRecurse"
                Content="Recurse subfolders"
                Margin="0,0,10,0"
                VerticalAlignment="Center"
                ToolTip="Include all nested subfolders. This can take time on large or remote trees."/>
      <CheckBox x:Name="chkShowSystem"
                Content="Show hidden/system"
                Margin="0,0,16,0"
                VerticalAlignment="Center"
                ToolTip="Show items marked Hidden or System. Off by default."/>
      <TextBlock Text="Filter:" VerticalAlignment="Center" Margin="0,0,6,0"/>
      <TextBox x:Name="txtFilter"
               Width="200"
               Margin="0,0,6,0"
               VerticalAlignment="Center"
               ToolTip="Filter by name, extension, or full path."/>
      <Button x:Name="btnClearFilter"
              Content="Clear Filter"
              Padding="8,2"
              Margin="0,0,10,0"
              ToolTip="Clear text and size filters."/>
      <TextBlock Text="Min size (MB):" VerticalAlignment="Center" Margin="0,0,6,0"/>
      <TextBox x:Name="txtMinSizeMB"
               Width="70"
               Margin="0,0,16,0"
               VerticalAlignment="Center"
               ToolTip="Only show files at or above this size. Leave blank for no size filter."/>
      <TextBlock Text="Quick:" VerticalAlignment="Center" Margin="0,0,6,0"/>
      <ComboBox x:Name="cmbQuick"
                Width="170"
                Margin="0,0,16,0"
                VerticalAlignment="Center"
                ToolTip="Jump to common locations or mapped drives."/>
      <TextBlock Text="Profile:" VerticalAlignment="Center" Margin="0,0,6,0"/>
      <ComboBox x:Name="cmbProfile"
                Width="220"
                VerticalAlignment="Center"
                ToolTip="Presets for common cleanup and large-file scans."/>
    </DockPanel>

    <DataGrid x:Name="dgItems"
              Grid.Row="2"
              Margin="0,0,0,8"
              AutoGenerateColumns="False"
              IsReadOnly="True"
              SelectionMode="Extended"
              SelectionUnit="FullRow"
              HeadersVisibility="Column"
              CanUserAddRows="False"
              CanUserDeleteRows="False"
              CanUserReorderColumns="True"
              GridLinesVisibility="Horizontal"
              AlternationCount="2"
              Background="White"
              Foreground="Black"
              RowBackground="White"
              AlternatingRowBackground="#FFF6F6F6">
      <DataGrid.Columns>
        <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="2*"/>
        <DataGridTextColumn Header="Extension" Binding="{Binding Extension}" Width="*"/>
        <DataGridTextColumn Header="Type" Binding="{Binding ItemType}" Width="*"/>
        <DataGridTextColumn Header="Size (KB)" Binding="{Binding SizeKB}" Width="*"/>
        <DataGridTextColumn Header="Created" Binding="{Binding Created}" Width="2*"/>
        <DataGridTextColumn Header="Modified" Binding="{Binding Modified}" Width="2*"/>
        <DataGridTextColumn Header="Attributes" Binding="{Binding Attributes}" Width="1.2*"/>
        <DataGridTextColumn Header="Full Path" Binding="{Binding FullPath}" Width="3*"/>
      </DataGrid.Columns>
    </DataGrid>

    <DockPanel Grid.Row="3">
      <StackPanel Orientation="Horizontal" DockPanel.Dock="Left">
        <Button x:Name="btnDeleteSelected"
                Content="Delete Selected..."
                Padding="12,2"
                Margin="0,0,10,0"
                ToolTip="Delete all selected items after confirmation."/>
        <Button x:Name="btnExportCsv"
                Content="Export CSV..."
                Padding="12,2"
                Margin="0,0,10,0"
                ToolTip="Export the current grid view to CSV."/>
        <Button x:Name="btnRefresh"
                Content="Refresh"
                Padding="12,2"
                ToolTip="Reload the current folder."/>
      </StackPanel>
      <TextBlock x:Name="txtStatus"
                 DockPanel.Dock="Right"
                 HorizontalAlignment="Right"
                 VerticalAlignment="Center"
                 Text="Ready."
                 Foreground="Black"
                 TextTrimming="CharacterEllipsis"/>
    </DockPanel>

    <Grid x:Name="BusyOverlay"
          Grid.RowSpan="4"
          Background="#7F000000"
          Visibility="Collapsed">
      <Border HorizontalAlignment="Center"
              VerticalAlignment="Center"
              Background="White"
              CornerRadius="8"
              Padding="22"
              MinWidth="380"
              MaxWidth="760"
              Opacity="0.98">
        <StackPanel>
          <TextBlock x:Name="txtBusyMessage"
                     Text="Scanning..."
                     Margin="0,0,0,10"
                     FontSize="15"
                     FontWeight="Bold"
                     Foreground="Black"
                     TextWrapping="Wrap"
                     TextAlignment="Center"/>
          <ProgressBar x:Name="pbBusy"
                       IsIndeterminate="True"
                       Width="320"
                       Height="18"
                       HorizontalAlignment="Center"/>
          <Button x:Name="btnCancelScan"
                  Content="Cancel Scan"
                  Width="120"
                  Margin="0,14,0,0"
                  HorizontalAlignment="Center"
                  ToolTip="Cancel the current scan."/>
        </StackPanel>
      </Border>
    </Grid>
  </Grid>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)
if (-not $window) { throw 'Failed to load XAML window.' }

$txtPath           = $window.FindName('txtPath')
$btnBrowse         = $window.FindName('btnBrowse')
$btnUp             = $window.FindName('btnUp')
$btnLoad           = $window.FindName('btnLoad')
$btnOpenExplorer   = $window.FindName('btnOpenExplorer')
$chkRecurse        = $window.FindName('chkRecurse')
$chkShowSystem     = $window.FindName('chkShowSystem')
$txtFilter         = $window.FindName('txtFilter')
$btnClearFilter    = $window.FindName('btnClearFilter')
$txtMinSizeMB      = $window.FindName('txtMinSizeMB')
$cmbQuick          = $window.FindName('cmbQuick')
$cmbProfile        = $window.FindName('cmbProfile')
$dgItems           = $window.FindName('dgItems')
$btnDeleteSelected = $window.FindName('btnDeleteSelected')
$btnExportCsv      = $window.FindName('btnExportCsv')
$btnRefresh        = $window.FindName('btnRefresh')
$txtStatus         = $window.FindName('txtStatus')
$busyOverlay       = $window.FindName('BusyOverlay')
$txtBusyMessage    = $window.FindName('txtBusyMessage')
$btnCancelScan     = $window.FindName('btnCancelScan')
$pbBusy            = $window.FindName('pbBusy')

$script:AllItems = @()
$script:LastScanMeta = [pscustomobject]@{
    Root = ''
    Files = 0
    Dirs = 0
    Skipped = 0
    ElapsedSeconds = 0
}
$script:ScanState = [ordered]@{
    Process       = $null
    Timer         = $null
    TempRoot      = $null
    WorkerScript  = $null
    ItemsFile     = $null
    MetaFile      = $null
    ProgressFile  = $null
    ErrorFile     = $null
    StartTime     = $null
    RequestedPath = $null
}

$cm = New-Object System.Windows.Controls.ContextMenu
$miOpen = New-Object System.Windows.Controls.MenuItem
$miOpen.Header  = 'Open'
$miOpen.ToolTip = 'Open the selected file or folder.'
$miOpenFolder = New-Object System.Windows.Controls.MenuItem
$miOpenFolder.Header  = 'Open Containing Folder'
$miOpenFolder.ToolTip = 'Open File Explorer and select the item.'
$miCopyPaths = New-Object System.Windows.Controls.MenuItem
$miCopyPaths.Header  = 'Copy Path(s)'
$miCopyPaths.ToolTip = 'Copy selected full paths to the clipboard.'
$miDelete = New-Object System.Windows.Controls.MenuItem
$miDelete.Header  = 'Delete Selected...'
$miDelete.ToolTip = 'Delete selected items after confirmation.'
[void]$cm.Items.Add($miOpen)
[void]$cm.Items.Add($miOpenFolder)
[void]$cm.Items.Add($miCopyPaths)
[void]$cm.Items.Add((New-Object System.Windows.Controls.Separator))
[void]$cm.Items.Add($miDelete)
$dgItems.ContextMenu = $cm

function Set-Status {
    param([string]$Message)
    if ($txtStatus) { $txtStatus.Text = $Message }
}

function Set-UiBusyState {
    param([bool]$IsBusy)

    foreach ($control in @(
        $txtPath, $btnBrowse, $btnUp, $btnLoad, $btnOpenExplorer,
        $chkRecurse, $chkShowSystem, $txtFilter, $btnClearFilter,
        $txtMinSizeMB, $cmbQuick, $cmbProfile, $dgItems,
        $btnDeleteSelected, $btnExportCsv, $btnRefresh
    )) {
        if ($control) { $control.IsEnabled = -not $IsBusy }
    }

    if ($btnCancelScan) { $btnCancelScan.IsEnabled = $IsBusy }
}

function Show-Busy {
    param([string]$Message = 'Working...')
    if ($txtBusyMessage) { $txtBusyMessage.Text = $Message }
    if ($busyOverlay) { $busyOverlay.Visibility = 'Visible' }
    Set-UiBusyState -IsBusy $true
    $window.Dispatcher.Invoke([System.Action]{}, [System.Windows.Threading.DispatcherPriority]::Render)
}

function Hide-Busy {
    if ($busyOverlay) { $busyOverlay.Visibility = 'Collapsed' }
    Set-UiBusyState -IsBusy $false
}

function Quote-CliArgument {
    param([Parameter(Mandatory = $true)][string]$Value)
    '"{0}"' -f ($Value -replace '"', '\"')
}

function Remove-ScanArtifacts {
    if ($script:ScanState.Timer) {
        try { $script:ScanState.Timer.Stop() } catch {}
    }
    if ($script:ScanState.Process) {
        try {
            if (-not $script:ScanState.Process.HasExited) {
                $script:ScanState.Process.Kill()
            }
        } catch {}
    }
    if ($script:ScanState.TempRoot -and (Test-Path -LiteralPath $script:ScanState.TempRoot)) {
        try { Remove-Item -LiteralPath $script:ScanState.TempRoot -Recurse -Force -ErrorAction SilentlyContinue } catch {}
    }
    $script:ScanState.Process = $null
    $script:ScanState.Timer = $null
    $script:ScanState.TempRoot = $null
    $script:ScanState.WorkerScript = $null
    $script:ScanState.ItemsFile = $null
    $script:ScanState.MetaFile = $null
    $script:ScanState.ProgressFile = $null
    $script:ScanState.ErrorFile = $null
    $script:ScanState.StartTime = $null
    $script:ScanState.RequestedPath = $null
}

function Update-BusyProgressFromFile {
    if (-not $script:ScanState.ProgressFile) { return }
    if (-not (Test-Path -LiteralPath $script:ScanState.ProgressFile)) { return }

    try {
        $raw = Get-Content -LiteralPath $script:ScanState.ProgressFile -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return }
        $p = $raw | ConvertFrom-Json -ErrorAction Stop
        $currentPath = [string]$p.CurrentPath
        if ($currentPath.Length -gt 110) {
            $currentPath = '...' + $currentPath.Substring($currentPath.Length - 107)
        }
        $message = 'Scanning... Items: {0} | Files: {1} | Dirs: {2} | Skipped: {3}' -f $p.Processed, $p.Files, $p.Dirs, $p.Skipped
        if (-not [string]::IsNullOrWhiteSpace($currentPath)) {
            $message += "`n$currentPath"
        }
        if ($txtBusyMessage) { $txtBusyMessage.Text = $message }
        Set-Status ($message -replace "`n", ' | ')
    }
    catch {}
}

function Get-WorkerScriptText {
@'
param(
    [Parameter(Mandatory=$true)][string]$Path,
    [Parameter(Mandatory=$true)][int]$Recurse,
    [Parameter(Mandatory=$true)][string]$ItemsFile,
    [Parameter(Mandatory=$true)][string]$MetaFile,
    [Parameter(Mandatory=$true)][string]$ProgressFile,
    [Parameter(Mandatory=$true)][string]$ErrorFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-ProgressState {
    param(
        [string]$Message,
        [int]$Processed,
        [int]$Files,
        [int]$Dirs,
        [int]$Skipped,
        [string]$CurrentPath,
        [double]$ElapsedSeconds
    )
    try {
        [ordered]@{
            Message = $Message
            Processed = $Processed
            Files = $Files
            Dirs = $Dirs
            Skipped = $Skipped
            CurrentPath = $CurrentPath
            ElapsedSeconds = [math]::Round($ElapsedSeconds, 2)
        } | ConvertTo-Json -Compress | Set-Content -LiteralPath $ProgressFile -Encoding UTF8
    }
    catch {}
}

try {
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        throw "Path does not exist or is not a folder: $Path"
    }

    $resolved = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
    $items = New-Object System.Collections.ArrayList
    $processed = 0
    $files = 0
    $dirs = 0
    $skipped = 0
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    function Add-ResultItem {
        param([object]$InputObject)

        $sizeKB = ''
        if (-not $InputObject.PSIsContainer) {
            try {
                $sizeKB = [string]([math]::Round($InputObject.Length / 1KB, 1))
            }
            catch {
                $sizeKB = ''
            }
            $script:files++
        }
        else {
            $script:dirs++
        }

        $obj = New-Object psobject -Property @{
            Name       = [string]$InputObject.Name
            Extension  = [string]$InputObject.Extension
            ItemType   = if ($InputObject.PSIsContainer) { 'Directory' } elseif ($InputObject -is [System.IO.FileInfo]) { 'File' } else { 'Other' }
            SizeKB     = $sizeKB
            Created    = [string]$InputObject.CreationTime
            Modified   = [string]$InputObject.LastWriteTime
            Attributes = [string]$InputObject.Attributes
            FullPath   = [string]$InputObject.FullName
        }
        [void]$script:items.Add($obj)
        $script:processed++
    }

    Write-ProgressState -Message 'Starting scan' -Processed 0 -Files 0 -Dirs 0 -Skipped 0 -CurrentPath $resolved -ElapsedSeconds 0

    if ($Recurse -eq 1) {
        $queue = New-Object 'System.Collections.Generic.Queue[string]'
        $queue.Enqueue($resolved)

        while ($queue.Count -gt 0) {
            $current = $queue.Dequeue()
            Write-ProgressState -Message 'Scanning' -Processed $processed -Files $files -Dirs $dirs -Skipped $skipped -CurrentPath $current -ElapsedSeconds $stopwatch.Elapsed.TotalSeconds

            try {
                $children = @(Get-ChildItem -LiteralPath $current -Force -ErrorAction Stop)
            }
            catch {
                $script:skipped++
                continue
            }

            foreach ($child in $children) {
                try {
                    Add-ResultItem -InputObject $child
                    if ($child.PSIsContainer) {
                        try { $queue.Enqueue([string]$child.FullName) } catch { $script:skipped++ }
                    }
                }
                catch {
                    $script:skipped++
                }

                if (($processed % 200) -eq 0) {
                    Write-ProgressState -Message 'Scanning' -Processed $processed -Files $files -Dirs $dirs -Skipped $skipped -CurrentPath $current -ElapsedSeconds $stopwatch.Elapsed.TotalSeconds
                }
            }
        }
    }
    else {
        $children = @(Get-ChildItem -LiteralPath $resolved -Force -ErrorAction Stop)
        foreach ($child in $children) {
            try {
                Add-ResultItem -InputObject $child
            }
            catch {
                $script:skipped++
            }
            if (($processed % 200) -eq 0) {
                Write-ProgressState -Message 'Scanning' -Processed $processed -Files $files -Dirs $dirs -Skipped $skipped -CurrentPath $resolved -ElapsedSeconds $stopwatch.Elapsed.TotalSeconds
            }
        }
    }

    if ($items.Count -gt 0) {
        $items | Export-Csv -LiteralPath $ItemsFile -NoTypeInformation -Encoding UTF8
    }
    else {
        '' | Set-Content -LiteralPath $ItemsFile -Encoding UTF8
    }

    [ordered]@{
        Success = $true
        Root = $resolved
        Files = $files
        Dirs = $dirs
        Skipped = $skipped
        ElapsedSeconds = [math]::Round($stopwatch.Elapsed.TotalSeconds, 2)
    } | ConvertTo-Json -Compress | Set-Content -LiteralPath $MetaFile -Encoding UTF8

    Write-ProgressState -Message 'Completed' -Processed $processed -Files $files -Dirs $dirs -Skipped $skipped -CurrentPath $resolved -ElapsedSeconds $stopwatch.Elapsed.TotalSeconds
    exit 0
}
catch {
    $msg = ($_ | Out-String)
    try { $msg | Set-Content -LiteralPath $ErrorFile -Encoding UTF8 } catch {}
    try {
        [ordered]@{
            Success = $false
            Error = $msg
        } | ConvertTo-Json -Compress | Set-Content -LiteralPath $MetaFile -Encoding UTF8
    }
    catch {}
    exit 1
}
'@
}

function Apply-Filter {
    param([string]$FilterText)

    $source = @($script:AllItems)
    if ($source.Count -eq 0) {
        $dgItems.ItemsSource = $null
        Set-Status 'No items loaded.'
        return
    }

    $filtered = $source
    $showSystem = ($chkShowSystem.IsChecked -eq $true)
    if (-not $showSystem) {
        $filtered = @($filtered | Where-Object { $_.Attributes -notmatch 'Hidden' -and $_.Attributes -notmatch 'System' })
    }

    $minSizeMB = 0.0
    $minSizeKB = 0.0
    if (-not [string]::IsNullOrWhiteSpace($txtMinSizeMB.Text)) {
        [void][double]::TryParse($txtMinSizeMB.Text, [ref]$minSizeMB)
        if ($minSizeMB -lt 0) { $minSizeMB = 0 }
        $minSizeKB = $minSizeMB * 1024
    }
    if ($minSizeKB -gt 0) {
        $filtered = @($filtered | Where-Object {
            if ($_.ItemType -ne 'File') { return $true }
            $size = 0.0
            [void][double]::TryParse([string]$_.SizeKB, [ref]$size)
            $size -ge $minSizeKB
        })
    }

    if (-not [string]::IsNullOrWhiteSpace($FilterText)) {
        $pattern = [System.Text.RegularExpressions.Regex]::Escape($FilterText)
        $filtered = @($filtered | Where-Object {
            $_.Name -match $pattern -or
            $_.Extension -match $pattern -or
            $_.FullPath -match $pattern
        })
    }

    $dgItems.ItemsSource = @($filtered)

    $shownCount = @($filtered).Count
    $fileShown = @($filtered | Where-Object { $_.ItemType -eq 'File' }).Count
    $dirShown = @($filtered | Where-Object { $_.ItemType -eq 'Directory' }).Count
    $shownSizeKB = (@($filtered | ForEach-Object {
        $n = 0.0
        if ([double]::TryParse([string]$_.SizeKB, [ref]$n)) { $n }
    } | Measure-Object -Sum).Sum)
    if (-not $shownSizeKB) { $shownSizeKB = 0 }
    $shownSizeMB = [math]::Round($shownSizeKB / 1024, 2)

    $status = '{0} total item(s) from {1} | Showing {2} (Files: {3}, Dirs: {4}) | Shown size: {5} MB | Skipped: {6} | Last scan: {7}s' -f `
        @($source).Count, `
        $txtPath.Text, `
        $shownCount, `
        $fileShown, `
        $dirShown, `
        $shownSizeMB, `
        $script:LastScanMeta.Skipped, `
        $script:LastScanMeta.ElapsedSeconds
    Set-Status $status
}

function Complete-Scan {
    param([bool]$Cancelled = $false)

    try {
        Update-BusyProgressFromFile

        if ($Cancelled) {
            Set-Status 'Scan cancelled.'
            return
        }

        if (-not $script:ScanState.MetaFile -or -not (Test-Path -LiteralPath $script:ScanState.MetaFile)) {
            $msg = 'Scan finished but no metadata file was produced.'
            if ($script:ScanState.ErrorFile -and (Test-Path -LiteralPath $script:ScanState.ErrorFile)) {
                try { $msg = Get-Content -LiteralPath $script:ScanState.ErrorFile -Raw -ErrorAction Stop } catch {}
            }
            [System.Windows.MessageBox]::Show($msg, 'Error', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
            Set-Status 'Scan failed.'
            return
        }

        $meta = Get-Content -LiteralPath $script:ScanState.MetaFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if (-not $meta.Success) {
            $msg = [string]$meta.Error
            if ([string]::IsNullOrWhiteSpace($msg) -and $script:ScanState.ErrorFile -and (Test-Path -LiteralPath $script:ScanState.ErrorFile)) {
                try { $msg = Get-Content -LiteralPath $script:ScanState.ErrorFile -Raw -ErrorAction Stop } catch {}
            }
            if ([string]::IsNullOrWhiteSpace($msg)) {
                $msg = 'The scan process exited with an unknown error.'
            }
            [System.Windows.MessageBox]::Show($msg, 'Error', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
            Set-Status 'Scan failed.'
            return
        }

        $items = @()
        if ($script:ScanState.ItemsFile -and (Test-Path -LiteralPath $script:ScanState.ItemsFile)) {
            try {
                $peek = Get-Content -LiteralPath $script:ScanState.ItemsFile -TotalCount 1 -ErrorAction Stop
                if (-not [string]::IsNullOrWhiteSpace($peek)) {
                    $items = @(Import-Csv -LiteralPath $script:ScanState.ItemsFile -ErrorAction Stop)
                }
            }
            catch {
                $items = @()
            }
        }

        $script:AllItems = @($items)
        $script:LastScanMeta = [pscustomobject]@{
            Root = [string]$meta.Root
            Files = [int]$meta.Files
            Dirs = [int]$meta.Dirs
            Skipped = [int]$meta.Skipped
            ElapsedSeconds = [double]$meta.ElapsedSeconds
        }
        $txtPath.Text = [string]$meta.Root
        Apply-Filter -FilterText $txtFilter.Text
    }
    catch {
        $msg = ($_ | Out-String)
        [System.Windows.MessageBox]::Show($msg, 'Error', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        Set-Status 'Scan failed.'
    }
    finally {
        Hide-Busy
        Remove-ScanArtifacts
    }
}

function Start-DirectoryScan {
    param([Parameter(Mandatory = $true)][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        [System.Windows.MessageBox]::Show('Please enter a directory path.', 'Missing Path', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        return
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        [System.Windows.MessageBox]::Show("The path does not exist or is not a folder.`n`n$Path", 'Invalid Path', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    Remove-ScanArtifacts

    $resolved = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
    $recurseValue = if ($chkRecurse.IsChecked -eq $true) { 1 } else { 0 }

    $tempRoot = Join-Path $env:TEMP ('DirectoryExplorerWorker_' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

    $workerScript = Join-Path $tempRoot 'Worker.ps1'
    $itemsFile = Join-Path $tempRoot 'Items.csv'
    $metaFile = Join-Path $tempRoot 'Meta.json'
    $progressFile = Join-Path $tempRoot 'Progress.json'
    $errorFile = Join-Path $tempRoot 'Error.txt'

    Set-Content -LiteralPath $workerScript -Value (Get-WorkerScriptText) -Encoding UTF8

    $argLine = @(
        '-NoProfile'
        '-ExecutionPolicy Bypass'
        '-File ' + (Quote-CliArgument $workerScript)
        '-Path ' + (Quote-CliArgument $resolved)
        '-Recurse ' + $recurseValue
        '-ItemsFile ' + (Quote-CliArgument $itemsFile)
        '-MetaFile ' + (Quote-CliArgument $metaFile)
        '-ProgressFile ' + (Quote-CliArgument $progressFile)
        '-ErrorFile ' + (Quote-CliArgument $errorFile)
    ) -join ' '

    Show-Busy ("Scanning...`n$resolved")
    Set-Status ("Starting scan: $resolved")

    $process = Start-Process -FilePath 'powershell.exe' -ArgumentList $argLine -WindowStyle Hidden -PassThru

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(450)
    $timer.Add_Tick({
        Update-BusyProgressFromFile
        if ($script:ScanState.Process -and $script:ScanState.Process.HasExited) {
            try { $script:ScanState.Timer.Stop() } catch {}
            Complete-Scan -Cancelled $false
        }
    })

    $script:ScanState.Process = $process
    $script:ScanState.Timer = $timer
    $script:ScanState.TempRoot = $tempRoot
    $script:ScanState.WorkerScript = $workerScript
    $script:ScanState.ItemsFile = $itemsFile
    $script:ScanState.MetaFile = $metaFile
    $script:ScanState.ProgressFile = $progressFile
    $script:ScanState.ErrorFile = $errorFile
    $script:ScanState.StartTime = Get-Date
    $script:ScanState.RequestedPath = $resolved

    $timer.Start()
}

function Go-UpDirectory {
    if ([string]::IsNullOrWhiteSpace($txtPath.Text)) { return }
    try {
        $resolved = (Resolve-Path -LiteralPath $txtPath.Text -ErrorAction Stop).ProviderPath
        $parent = [System.IO.Directory]::GetParent($resolved)
        if ($parent -and $parent.FullName) {
            Start-DirectoryScan -Path $parent.FullName
        }
        else {
            Set-Status "Already at top-level: $resolved"
        }
    }
    catch {
        [System.Windows.MessageBox]::Show((($_ | Out-String).Trim()), 'Cannot Go Up', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
    }
}

function Open-ItemPath {
    param([string]$FullPath)
    try {
        if (Test-Path -LiteralPath $FullPath) {
            Start-Process -FilePath $FullPath
        }
        else {
            [System.Windows.MessageBox]::Show("The item no longer exists:`n`n$FullPath", 'Not Found', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        }
    }
    catch {
        [System.Windows.MessageBox]::Show((($_ | Out-String).Trim()), 'Error', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
    }
}

function Open-ContainingFolder {
    param([string]$FullPath)
    try {
        if (Test-Path -LiteralPath $FullPath) {
            Start-Process explorer.exe ("/select,`"$FullPath`"")
        }
        else {
            [System.Windows.MessageBox]::Show("The item no longer exists:`n`n$FullPath", 'Not Found', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        }
    }
    catch {
        [System.Windows.MessageBox]::Show((($_ | Out-String).Trim()), 'Error', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
    }
}

function Delete-SelectedItems {
    $selected = @($dgItems.SelectedItems)
    if ($selected.Count -eq 0) {
        [System.Windows.MessageBox]::Show('No items are selected to delete.', 'Nothing Selected', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        return
    }

    $firstPath = [string]$selected[0].FullPath
    $msg = "You are about to delete $($selected.Count) item(s). This operation cannot be undone.`n`nFirst item:`n$firstPath`n`nContinue?"
    $result = [System.Windows.MessageBox]::Show($msg, 'Confirm Delete', [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    if ($result -ne [System.Windows.MessageBoxResult]::Yes) { return }

    Show-Busy 'Deleting selected items...'
    try {
        foreach ($row in $selected) {
            $p = [string]$row.FullPath
            try {
                if (Test-Path -LiteralPath $p) {
                    $item = Get-Item -LiteralPath $p -ErrorAction Stop
                    if ($item.PSIsContainer) {
                        Remove-Item -LiteralPath $p -Recurse -Force -ErrorAction Stop
                    }
                    else {
                        Remove-Item -LiteralPath $p -Force -ErrorAction Stop
                    }
                }
            }
            catch {
                [System.Windows.MessageBox]::Show("Failed to delete:`n$p`n`n$($_.Exception.Message)", 'Delete Error', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
            }
        }
    }
    finally {
        Hide-Busy
    }

    Start-DirectoryScan -Path $txtPath.Text
}

function Export-CurrentViewToCsv {
    $data = @($dgItems.ItemsSource)
    if ($data.Count -eq 0) {
        [System.Windows.MessageBox]::Show('No data to export.', 'Export CSV', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        return
    }

    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
    $dlg.FileName = 'DirectoryExport.csv'
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $data | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
            [System.Windows.MessageBox]::Show("Exported $($data.Count) row(s) to:`n`n$($dlg.FileName)", 'Export CSV', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        }
        catch {
            [System.Windows.MessageBox]::Show((($_ | Out-String).Trim()), 'Export Error', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        }
    }
}

function Copy-SelectedPaths {
    $selected = @($dgItems.SelectedItems)
    if ($selected.Count -eq 0) {
        [System.Windows.MessageBox]::Show('No items selected to copy paths from.', 'Copy Path(s)', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        return
    }

    $paths = ($selected | ForEach-Object { [string]$_.FullPath }) -join "`r`n"
    try {
        [System.Windows.Clipboard]::SetText($paths)
        Set-Status ("Copied {0} path(s) to clipboard." -f $selected.Count)
    }
    catch {
        [System.Windows.MessageBox]::Show((($_ | Out-String).Trim()), 'Clipboard Error', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
    }
}

function Add-QuickItem {
    param(
        [System.Windows.Controls.ComboBox]$Combo,
        [string]$Label,
        [string]$Path
    )
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    if (-not (Test-Path -LiteralPath $Path -ErrorAction SilentlyContinue)) { return }

    $item = New-Object System.Windows.Controls.ComboBoxItem
    $item.Content = $Label
    $item.Tag = $Path
    [void]$Combo.Items.Add($item)
}

function Add-ProfileItem {
    param(
        [System.Windows.Controls.ComboBox]$Combo,
        [string]$Label,
        [string]$Path,
        [bool]$Recurse,
        [double]$MinSizeMB
    )
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    if (-not (Test-Path -LiteralPath $Path -ErrorAction SilentlyContinue)) { return }

    $obj = [pscustomobject]@{
        Path = $Path
        Recurse = $Recurse
        MinSizeMB = $MinSizeMB
    }

    $item = New-Object System.Windows.Controls.ComboBoxItem
    $item.Content = $Label
    $item.Tag = $obj
    [void]$Combo.Items.Add($item)
}

function Show-FolderPicker {
    param([string]$InitialPath)
    try {
        $shell = New-Object -ComObject Shell.Application
        $folder = $shell.BrowseForFolder(0, 'Select a folder', 0, 0)
        if ($folder -and $folder.Self -and $folder.Self.Path) {
            return [string]$folder.Self.Path
        }
    }
    catch {}

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'Select a folder'
    if (-not [string]::IsNullOrWhiteSpace($InitialPath) -and (Test-Path -LiteralPath $InitialPath -PathType Container)) {
        $dialog.SelectedPath = $InitialPath
    }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.SelectedPath
    }

    return $null
}

# Events
$btnLoad.Add_Click({ Start-DirectoryScan -Path $txtPath.Text })
$btnRefresh.Add_Click({ Start-DirectoryScan -Path $txtPath.Text })
$btnUp.Add_Click({ Go-UpDirectory })

$btnBrowse.Add_Click({
    $picked = Show-FolderPicker -InitialPath $txtPath.Text
    if (-not [string]::IsNullOrWhiteSpace($picked)) {
        $txtPath.Text = $picked
        Start-DirectoryScan -Path $picked
    }
})

$btnOpenExplorer.Add_Click({
    if (Test-Path -LiteralPath $txtPath.Text -PathType Container) {
        Start-Process explorer.exe $txtPath.Text
    }
    else {
        [System.Windows.MessageBox]::Show("Current path is not a valid folder:`n`n$($txtPath.Text)", 'Invalid Path', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
    }
})

$txtPath.Add_KeyDown({
    param($sender, $e)
    if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
        Start-DirectoryScan -Path $txtPath.Text
        $e.Handled = $true
    }
})

$txtFilter.Add_TextChanged({ if (-not $busyOverlay.IsVisible) { Apply-Filter -FilterText $txtFilter.Text } })
$txtMinSizeMB.Add_TextChanged({ if (-not $busyOverlay.IsVisible) { Apply-Filter -FilterText $txtFilter.Text } })
$chkShowSystem.Add_Click({ if (-not $busyOverlay.IsVisible) { Apply-Filter -FilterText $txtFilter.Text } })

$btnClearFilter.Add_Click({
    $txtFilter.Text = ''
    $txtMinSizeMB.Text = ''
    Apply-Filter -FilterText ''
})

$btnDeleteSelected.Add_Click({ Delete-SelectedItems })
$btnExportCsv.Add_Click({ Export-CurrentViewToCsv })
$btnCancelScan.Add_Click({
    if ($script:ScanState.Process -and -not $script:ScanState.Process.HasExited) {
        try { $script:ScanState.Process.Kill() } catch {}
    }
    try { if ($script:ScanState.Timer) { $script:ScanState.Timer.Stop() } } catch {}
    Hide-Busy
    Remove-ScanArtifacts
    Set-Status 'Scan cancelled.'
})

$dgItems.Add_MouseDoubleClick({
    if ($dgItems.SelectedItem) {
        $row = $dgItems.SelectedItem
        if ($row.ItemType -eq 'Directory') {
            Start-DirectoryScan -Path ([string]$row.FullPath)
        }
        else {
            Open-ItemPath -FullPath ([string]$row.FullPath)
        }
    }
})

$miOpen.Add_Click({
    if ($dgItems.SelectedItem) {
        $row = $dgItems.SelectedItem
        if ($row.ItemType -eq 'Directory') {
            Start-DirectoryScan -Path ([string]$row.FullPath)
        }
        else {
            Open-ItemPath -FullPath ([string]$row.FullPath)
        }
    }
})
$miOpenFolder.Add_Click({ if ($dgItems.SelectedItem) { Open-ContainingFolder -FullPath ([string]$dgItems.SelectedItem.FullPath) } })
$miCopyPaths.Add_Click({ Copy-SelectedPaths })
$miDelete.Add_Click({ Delete-SelectedItems })

$window.Add_KeyDown({
    param($sender, $e)

    if ($busyOverlay.IsVisible) {
        if ($e.Key -eq [System.Windows.Input.Key]::Escape) {
            $btnCancelScan.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
            $e.Handled = $true
        }
        return
    }

    if ($e.Key -eq [System.Windows.Input.Key]::F5) {
        Start-DirectoryScan -Path $txtPath.Text
        $e.Handled = $true
    }
    elseif ($e.Key -eq [System.Windows.Input.Key]::Delete -and $dgItems.IsKeyboardFocusWithin) {
        Delete-SelectedItems
        $e.Handled = $true
    }
    elseif ($e.Key -eq [System.Windows.Input.Key]::Back -and $dgItems.IsKeyboardFocusWithin) {
        Go-UpDirectory
        $e.Handled = $true
    }
})

# Quick paths and profiles
$userProfile = [Environment]::GetFolderPath('UserProfile')
$documents = [Environment]::GetFolderPath('MyDocuments')
$desktop = [Environment]::GetFolderPath('Desktop')
$downloads = Join-Path $userProfile 'Downloads'
$systemDrive = [System.IO.Path]::GetPathRoot([Environment]::GetFolderPath('System'))

Add-QuickItem -Combo $cmbQuick -Label 'Desktop' -Path $desktop
Add-QuickItem -Combo $cmbQuick -Label 'Documents' -Path $documents
Add-QuickItem -Combo $cmbQuick -Label 'Downloads' -Path $downloads
Add-QuickItem -Combo $cmbQuick -Label 'System drive' -Path $systemDrive

try {
    $mappedDrives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.DisplayRoot -or $_.Root -like '\\*' }
    foreach ($drive in $mappedDrives) {
        Add-QuickItem -Combo $cmbQuick -Label ($drive.Name + ':') -Path $drive.Root
    }
}
catch {}

$cmbQuick.Add_SelectionChanged({
    param($sender, $e)
    $item = $sender.SelectedItem
    if ($item -and $item.Tag) {
        $p = [string]$item.Tag
        if (Test-Path -LiteralPath $p -PathType Container) {
            $txtPath.Text = $p
            Start-DirectoryScan -Path $p
        }
    }
})

$noProf = New-Object System.Windows.Controls.ComboBoxItem
$noProf.Content = '(No profile)'
$noProf.Tag = $null
[void]$cmbProfile.Items.Add($noProf)

Add-ProfileItem -Combo $cmbProfile -Label 'User profile - big files (>20MB)' -Path $userProfile -Recurse $true -MinSizeMB 20
Add-ProfileItem -Combo $cmbProfile -Label 'Downloads - big files (>50MB)' -Path $downloads -Recurse $true -MinSizeMB 50
Add-ProfileItem -Combo $cmbProfile -Label 'System drive - very large (>200MB)' -Path $systemDrive -Recurse $true -MinSizeMB 200

$cmbProfile.SelectedIndex = 0
$cmbProfile.Add_SelectionChanged({
    param($sender, $e)
    $item = $sender.SelectedItem
    if (-not $item) { return }
    if (-not $item.Tag) { return }

    $prof = $item.Tag
    if (-not (Test-Path -LiteralPath $prof.Path -PathType Container)) {
        [System.Windows.MessageBox]::Show("Profile path is not available:`n`n$($prof.Path)", 'Profile Error', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
        return
    }

    $txtPath.Text = [string]$prof.Path
    $chkRecurse.IsChecked = [bool]$prof.Recurse
    $txtMinSizeMB.Text = [string]$prof.MinSizeMB
    $txtFilter.Text = ''
    Start-DirectoryScan -Path ([string]$prof.Path)
})

$chkShowSystem.IsChecked = $false
$btnCancelScan.IsEnabled = $false

$window.Add_Closed({
    try {
        if ($script:ScanState.Process -and -not $script:ScanState.Process.HasExited) {
            $script:ScanState.Process.Kill()
        }
    }
    catch {}
    Remove-ScanArtifacts
})

$window.Add_ContentRendered({
    if (Test-Path -LiteralPath $userProfile -PathType Container) {
        $txtPath.Text = $userProfile
        Start-DirectoryScan -Path $userProfile
    }
    else {
        Set-Status 'Ready.'
    }
})

$window.Topmost = $false
$window.ShowDialog() | Out-Null
