<#
Directory Explorer GUI – v4 + Busy Overlay
- Hide system/hidden items by default (toggle)
- Size filter (Min size MB)
- Path quick-jump + profiles for "what's eating this drive"
- Recursive scan, filter, delete, CSV export, copy paths
- Busy overlay during scans
#>

if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    powershell.exe -STA -ExecutionPolicy Bypass -File $PSCommandPath @args
    exit
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -AssemblyName System.Windows.Forms

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Directory Explorer - PowerShell"
        Height="600"
        Width="1100"
        WindowStartupLocation="CenterScreen"
        Background="White"
        Foreground="Black">

  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>   <!-- Path bar -->
      <RowDefinition Height="Auto"/>   <!-- Filter bar -->
      <RowDefinition Height="*"/>      <!-- DataGrid -->
      <RowDefinition Height="Auto"/>   <!-- Status bar -->
    </Grid.RowDefinitions>

    <!-- Top controls: path + nav -->
    <DockPanel Grid.Row="0" Margin="0,0,0,8">
      <TextBlock Text="Path:" VerticalAlignment="Center" Margin="0,0,6,0"/>
      <TextBox x:Name="txtPath"
               Width="520"
               Margin="0,0,6,0"
               VerticalAlignment="Center"
               ToolTip="Type or paste a folder path, then click Load or press Enter."/>
      <Button x:Name="btnBrowse"
              Content="Browse..."
              Margin="0,0,6,0"
              Padding="10,2"
              ToolTip="Open a folder picker dialog and select the directory to scan."/>
      <Button x:Name="btnUp"
              Content="Up"
              Margin="0,0,6,0"
              Padding="10,2"
              ToolTip="Go to the parent folder of the current directory."/>
      <Button x:Name="btnLoad"
              Content="Load"
              Margin="0,0,6,0"
              Padding="12,2"
              ToolTip="Load the contents of the specified directory (files and subfolders, if recursion is enabled)."/>
      <Button x:Name="btnOpenExplorer"
              Content="Explorer"
              Padding="10,2"
              ToolTip="Open the current folder in File Explorer."/>
    </DockPanel>

    <!-- Filter + recursion + size + quick paths + profiles -->
    <DockPanel Grid.Row="1" Margin="0,0,0,8">
      <CheckBox x:Name="chkRecurse"
                Content="Recurse subfolders"
                Margin="0,0,10,0"
                VerticalAlignment="Center"
                ToolTip="When checked, includes all files and folders under this directory (can be slow on large trees)."/>
      <CheckBox x:Name="chkShowSystem"
                Content="Show hidden/system"
                Margin="0,0,16,0"
                VerticalAlignment="Center"
                ToolTip="Include items marked Hidden or System. Off = hide them (default)."/>
      <TextBlock Text="Filter:" VerticalAlignment="Center" Margin="0,0,6,0"/>
      <TextBox x:Name="txtFilter"
               Width="180"
               Margin="0,0,6,0"
               VerticalAlignment="Center"
               ToolTip="Filter by name, extension, or full path (simple contains match)."/>
      <Button x:Name="btnClearFilter"
              Content="Clear Filter"
              Padding="8,2"
              Margin="0,0,10,0"
              ToolTip="Clear the filter text and reapply filters."/>
      <TextBlock Text="Min size (MB):" VerticalAlignment="Center" Margin="0,0,6,0"/>
      <TextBox x:Name="txtMinSizeMB"
               Width="60"
               Margin="0,0,16,0"
               VerticalAlignment="Center"
               ToolTip="Only show files at least this size (in MB). Leave blank for no size filter."/>
      <TextBlock Text="Quick:" VerticalAlignment="Center" Margin="0,0,6,0"/>
      <ComboBox x:Name="cmbQuick"
                Width="140"
                Margin="0,0,16,0"
                VerticalAlignment="Center"
                ToolTip="Jump quickly to common locations (Desktop, Documents, Downloads, etc.)"/>
      <TextBlock Text="Profile:" VerticalAlignment="Center" Margin="0,0,6,0"/>
      <ComboBox x:Name="cmbProfile"
                Width="220"
                VerticalAlignment="Center"
                ToolTip="Presets for 'what''s eating this drive' scans (sets path + recurse + min size)."/>
    </DockPanel>

    <!-- Main DataGrid -->
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
              GridLinesVisibility="Horizontal"
              AlternationCount="2"
              Background="White"
              Foreground="Black"
              RowBackground="White"
              AlternatingRowBackground="#FFF5F5F5">
      <DataGrid.Columns>
        <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="2*"/>
        <DataGridTextColumn Header="Extension" Binding="{Binding Extension}" Width="*"/>
        <DataGridTextColumn Header="Type" Binding="{Binding ItemType}" Width="*"/>
        <DataGridTextColumn Header="Size (KB)" Binding="{Binding SizeKB}" Width="*"/>
        <DataGridTextColumn Header="Created" Binding="{Binding Created}" Width="2*"/>
        <DataGridTextColumn Header="Modified" Binding="{Binding Modified}" Width="2*"/>
        <DataGridTextColumn Header="Attributes" Binding="{Binding Attributes}" Width="*"/>
        <DataGridTextColumn Header="Full Path" Binding="{Binding FullPath}" Width="3*"/>
      </DataGrid.Columns>
    </DataGrid>

    <!-- Bottom bar -->
    <DockPanel Grid.Row="3">
      <StackPanel Orientation="Horizontal" DockPanel.Dock="Left">
        <Button x:Name="btnDeleteSelected"
                Content="Delete Selected..."
                Padding="12,2"
                Margin="0,0,10,0"
                ToolTip="Delete all currently selected items (files/folders) after confirmation."/>
        <Button x:Name="btnExportCsv"
                Content="Export CSV..."
                Padding="12,2"
                ToolTip="Export the currently displayed items to a CSV file."/>
      </StackPanel>
      <TextBlock x:Name="txtStatus"
                 DockPanel.Dock="Right"
                 HorizontalAlignment="Right"
                 VerticalAlignment="Center"
                 Text="Ready."
                 Foreground="Black"
                 TextTrimming="CharacterEllipsis"/>
    </DockPanel>

    <!-- Busy overlay -->
    <Grid x:Name="BusyOverlay"
          Grid.RowSpan="4"
          Background="#80000000"
          Visibility="Collapsed">
      <Border HorizontalAlignment="Center"
              VerticalAlignment="Center"
              Background="White"
              CornerRadius="6"
              Padding="20"
              Opacity="0.95">
        <StackPanel HorizontalAlignment="Center">
          <TextBlock x:Name="txtBusyMessage"
                     Text="Scanning..."
                     Margin="0,0,0,10"
                     FontSize="14"
                     FontWeight="Bold"
                     Foreground="Black"
                     TextAlignment="Center"/>
          <ProgressBar IsIndeterminate="True"
                       Width="260"
                       Height="18"/>
        </StackPanel>
      </Border>
    </Grid>

  </Grid>
</Window>
'@

# Load XAML
$reader  = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window  = [Windows.Markup.XamlReader]::Load($reader)
if (-not $window) { Write-Error "Failed to load XAML window."; return }

# Get controls
$txtPath           = $window.FindName("txtPath")
$btnBrowse         = $window.FindName("btnBrowse")
$btnUp             = $window.FindName("btnUp")
$btnLoad           = $window.FindName("btnLoad")
$btnOpenExplorer   = $window.FindName("btnOpenExplorer")
$chkRecurse        = $window.FindName("chkRecurse")
$chkShowSystem     = $window.FindName("chkShowSystem")
$txtFilter         = $window.FindName("txtFilter")
$btnClearFilter    = $window.FindName("btnClearFilter")
$txtMinSizeMB      = $window.FindName("txtMinSizeMB")
$cmbQuick          = $window.FindName("cmbQuick")
$cmbProfile        = $window.FindName("cmbProfile")
$dgItems           = $window.FindName("dgItems")
$btnDeleteSelected = $window.FindName("btnDeleteSelected")
$btnExportCsv      = $window.FindName("btnExportCsv")
$txtStatus         = $window.FindName("txtStatus")
$busyOverlay       = $window.FindName("BusyOverlay")
$txtBusyMessage    = $window.FindName("txtBusyMessage")

if (-not $dgItems -or -not $txtPath -or -not $txtStatus) {
    Write-Error "One or more UI elements could not be found. Check XAML names."
    return
}

# Context menu
$cm          = New-Object System.Windows.Controls.ContextMenu
$miOpen      = New-Object System.Windows.Controls.MenuItem
$miOpen.Header     = "Open"
$miOpen.ToolTip    = "Open the selected file or folder with its default handler."
$miOpenFolder      = New-Object System.Windows.Controls.MenuItem
$miOpenFolder.Header  = "Open Containing Folder"
$miOpenFolder.ToolTip = "Open File Explorer at the item's location and select it."
$miCopyPaths        = New-Object System.Windows.Controls.MenuItem
$miCopyPaths.Header = "Copy Path(s)"
$miCopyPaths.ToolTip = "Copy the FullPath of all selected items to the clipboard."
$miDelete           = New-Object System.Windows.Controls.MenuItem
$miDelete.Header    = "Delete Selected..."
$miDelete.ToolTip   = "Delete all currently selected items (with confirmation)."

[void]$cm.Items.Add($miOpen)
[void]$cm.Items.Add($miOpenFolder)
[void]$cm.Items.Add($miCopyPaths)
[void]$cm.Items.Add((New-Object System.Windows.Controls.Separator))
[void]$cm.Items.Add($miDelete)
$dgItems.ContextMenu = $cm

# Global data
$script:AllItems = @()

function Set-Status {
    param([string]$Message)
    if ($txtStatus -and $txtStatus -is [System.Windows.Controls.TextBlock]) {
        $txtStatus.Text = $Message
    }
}

function Show-Busy {
    param([string]$Message = "Working...")
    if ($busyOverlay) {
        if ($txtBusyMessage) { $txtBusyMessage.Text = $Message }
        $busyOverlay.Visibility = 'Visible'
        # Force UI to render the overlay before heavy work
        $window.Dispatcher.Invoke({},
            [System.Windows.Threading.DispatcherPriority]::Render)
    }
}

function Hide-Busy {
    if ($busyOverlay) {
        $busyOverlay.Visibility = 'Collapsed'
    }
}

function Get-DirectoryItems {
    param(
        [string]$Path,
        [bool]$Recurse = $false
    )

    $items = @()

    try {
        if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
            [System.Windows.MessageBox]::Show(
                "The path does not exist or is not a directory.`n`n$Path",
                "Invalid Path",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            ) | Out-Null
            return $items
        }

        if ($Recurse) {
            Set-Status "Scanning $Path (including subfolders)..."
            $dirItems = Get-ChildItem -LiteralPath $Path -Force -Recurse -ErrorAction Stop
        } else {
            Set-Status "Scanning $Path ..."
            $dirItems = Get-ChildItem -LiteralPath $Path -Force -ErrorAction Stop
        }

        foreach ($i in $dirItems) {
            $sizeKB = $null
            if (-not $i.PSIsContainer) {
                try { $sizeKB = [math]::Round($i.Length / 1KB, 1) } catch {}
            }

            $items += [PSCustomObject]@{
                Name      = $i.Name
                Extension = $i.Extension
                ItemType  = if ($i.PSIsContainer) { "Directory" } elseif ($i -is [System.IO.FileInfo]) { "File" } else { "Other" }
                SizeKB    = $sizeKB
                Created   = $i.CreationTime
                Modified  = $i.LastWriteTime
                Attributes= $i.Attributes.ToString()
                FullPath  = $i.FullName
            }
        }
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Failed to list directory:`n`n$($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
        Set-Status "Error loading $Path"
    }

    return $items
}

function Apply-Filter {
    param([string]$FilterText)

    $source = $script:AllItems
    if (-not $source) {
        $dgItems.ItemsSource = $null
        Set-Status "No items loaded."
        return
    }

    $filtered = $source

    # Hidden/system filter
    $showSystem = ($chkShowSystem.IsChecked -eq $true)
    if (-not $showSystem) {
        $filtered = $filtered | Where-Object {
            $_.Attributes -notmatch 'Hidden' -and $_.Attributes -notmatch 'System'
        }
    }

    # Size filter
    $minSizeMB = 0.0
    $minSizeKB = 0.0
    if ($txtMinSizeMB -and -not [string]::IsNullOrWhiteSpace($txtMinSizeMB.Text)) {
        [void][double]::TryParse($txtMinSizeMB.Text, [ref]$minSizeMB)
        if ($minSizeMB -lt 0) { $minSizeMB = 0 }
        $minSizeKB = $minSizeMB * 1024
    }

    if ($minSizeKB -gt 0) {
        $filtered = $filtered | Where-Object {
            if ($_.ItemType -ne 'File') { return $true }
            if (-not $_.SizeKB)        { return $false }
            return $_.SizeKB -ge $minSizeKB
        }
    }

    # Text filter
    if (-not [string]::IsNullOrWhiteSpace($FilterText)) {
        $pattern = [System.Text.RegularExpressions.Regex]::Escape($FilterText)
        $filtered = $filtered | Where-Object {
            $_.Name -match $pattern -or
            $_.Extension -match $pattern -or
            $_.FullPath -match $pattern
        }
    }

    $filteredList = @($filtered)
    $dgItems.ItemsSource = $filteredList

    $totalCount = $source.Count
    $fileAll    = ($source | Where-Object { $_.ItemType -eq 'File' }).Count
    $dirAll     = ($source | Where-Object { $_.ItemType -eq 'Directory' }).Count

    $shownCount = $filteredList.Count
    $fileShown  = ($filteredList | Where-Object { $_.ItemType -eq 'File' }).Count
    $dirShown   = ($filteredList | Where-Object { $_.ItemType -eq 'Directory' }).Count

    $totalSizeKB = ($source       | Where-Object { $_.SizeKB } | Measure-Object -Property SizeKB -Sum).Sum
    $shownSizeKB = ($filteredList | Where-Object { $_.SizeKB } | Measure-Object -Property SizeKB -Sum).Sum
    if (-not $totalSizeKB) { $totalSizeKB = 0 }
    if (-not $shownSizeKB) { $shownSizeKB = 0 }

    $shownSizeMB = [math]::Round($shownSizeKB / 1024, 2)

    $path = $txtPath.Text
    Set-Status ("{0} total item(s) from {1} | Showing {2} (Files: {3}, Dirs: {4}) | Shown size: {5} MB | Filter='{6}' MinSize={7}MB ShowSystem={8}" -f `
        $totalCount, $path, $shownCount, $fileShown, $dirShown, $shownSizeMB, $FilterText, $minSizeMB, $showSystem)
}

function Load-Directory {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        [System.Windows.MessageBox]::Show(
            "Please enter a directory path.",
            "Missing Path",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
        return
    }

    $resolved = Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue
    if ($resolved) { $Path = $resolved.ProviderPath }

    $recurse = $chkRecurse.IsChecked -eq $true

    Show-Busy "Scanning $Path ..."
    try {
        $items   = Get-DirectoryItems -Path $Path -Recurse:$recurse
        $script:AllItems = @($items)
        $txtPath.Text    = $Path
        Apply-Filter -FilterText $txtFilter.Text
    }
    finally {
        Hide-Busy
    }
}

function Go-UpDirectory {
    $current = $txtPath.Text
    if ([string]::IsNullOrWhiteSpace($current)) { return }

    try {
        $resolved  = Resolve-Path -LiteralPath $current -ErrorAction Stop
        $dirPath   = $resolved.ProviderPath
        $parentDir = [System.IO.Directory]::GetParent($dirPath)

        if ($parentDir -and $parentDir.FullName) {
            Load-Directory -Path $parentDir.FullName
        } else {
            Set-Status "Already at top-level: $dirPath"
        }
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Cannot go up from:`n`n$current`n`n$($_.Exception.Message)",
            "Cannot Go Up",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        ) | Out-Null
    }
}

function Open-ItemPath {
    param([string]$FullPath)

    try {
        if (Test-Path -LiteralPath $FullPath) {
            Start-Process -FilePath $FullPath
        } else {
            [System.Windows.MessageBox]::Show(
                "The item no longer exists:`n`n$FullPath",
                "Not Found",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            ) | Out-Null
        }
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Failed to open:`n`n$($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
    }
}

function Open-ContainingFolder {
    param([string]$FullPath)

    try {
        if (Test-Path -LiteralPath $FullPath) {
            $arg = "/select,""$FullPath"""
            Start-Process explorer.exe $arg
        } else {
            [System.Windows.MessageBox]::Show(
                "The item no longer exists:`n`n$FullPath",
                "Not Found",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            ) | Out-Null
        }
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Failed to open containing folder:`n`n$($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
    }
}

function Delete-SelectedItems {
    $selected = @($dgItems.SelectedItems)
    if (-not $selected -or $selected.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "No items are selected to delete.",
            "Nothing Selected",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
        return
    }

    $paths = $selected.FullPath
    $msg = "You are about to delete $($paths.Count) item(s). This operation cannot be undone.`n`nFirst item:`n$($paths[0])`n`nContinue?"
    $result = [System.Windows.MessageBox]::Show(
        $msg,
        "Confirm Delete",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Warning
    )
    if ($result -ne [System.Windows.MessageBoxResult]::Yes) { return }

    Show-Busy "Deleting selected items..."
    try {
        foreach ($p in $paths) {
            try {
                if (Test-Path -LiteralPath $p) {
                    $item = Get-Item -LiteralPath $p -ErrorAction Stop
                    if ($item.PSIsContainer) {
                        Remove-Item -LiteralPath $p -Recurse -Force -ErrorAction Stop
                    } else {
                        Remove-Item -LiteralPath $p -Force -ErrorAction Stop
                    }
                }
            }
            catch {
                [System.Windows.MessageBox]::Show(
                    "Failed to delete:`n$($p)`n`n$($_.Exception.Message)",
                    "Delete Error",
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Error
                ) | Out-Null
            }
        }

        Load-Directory -Path $txtPath.Text
    }
    finally {
        Hide-Busy
    }
}

function Export-CurrentViewToCsv {
    $data = @($dgItems.ItemsSource)
    if (-not $data -or $data.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "No data to export.",
            "Export CSV",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
        return
    }

    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter   = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $dlg.FileName = "DirectoryExport.csv"

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $data | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
            [System.Windows.MessageBox]::Show(
                "Exported $($data.Count) row(s) to:`n`n$($dlg.FileName)",
                "Export CSV",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            ) | Out-Null
        }
        catch {
            [System.Windows.MessageBox]::Show(
                "Failed to export CSV:`n`n$($_.Exception.Message)",
                "Export Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            ) | Out-Null
        }
    }
}

function Copy-SelectedPaths {
    $selected = @($dgItems.SelectedItems)
    if (-not $selected -or $selected.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "No items selected to copy paths from.",
            "Copy Path(s)",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
        return
    }

    $paths = $selected.FullPath -join "`r`n"
    try {
        [System.Windows.Clipboard]::SetText($paths)
        Set-Status "Copied $($selected.Count) path(s) to clipboard."
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Failed to copy to clipboard:`n`n$($_.Exception.Message)",
            "Clipboard Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
    }
}

function Add-QuickItem {
    param($combo, [string]$label, [string]$path)
    if (-not (Test-Path -LiteralPath $path -ErrorAction SilentlyContinue)) { return }
    $item         = New-Object System.Windows.Controls.ComboBoxItem
    $item.Content = $label
    $item.Tag     = $path
    [void]$combo.Items.Add($item)
}

function Add-ProfileItem {
    param($combo, [string]$label, [string]$path, [bool]$recurse, [double]$minSizeMB)
    if (-not (Test-Path -LiteralPath $path -ErrorAction SilentlyContinue)) { return }
    $obj = [PSCustomObject]@{
        Path      = $path
        Recurse   = $recurse
        MinSizeMB = $minSizeMB
    }
    $item         = New-Object System.Windows.Controls.ComboBoxItem
    $item.Content = $label
    $item.Tag     = $obj
    [void]$combo.Items.Add($item)
}

# --- Event wiring ---

$btnLoad.Add_Click({ Load-Directory -Path $txtPath.Text })

$btnBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select a folder to browse"
    if ($txtPath.Text -and (Test-Path -LiteralPath $txtPath.Text -PathType Container)) {
        $dialog.SelectedPath = $txtPath.Text
    }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtPath.Text = $dialog.SelectedPath
        Load-Directory -Path $dialog.SelectedPath
    }
})

$btnUp.Add_Click({ Go-UpDirectory })

$btnOpenExplorer.Add_Click({
    $path = $txtPath.Text
    if (Test-Path -LiteralPath $path -PathType Container) {
        Start-Process explorer.exe $path
    } else {
        [System.Windows.MessageBox]::Show(
            "Current path is not a valid folder:`n`n$path",
            "Invalid Path",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        ) | Out-Null
    }
})

$txtPath.Add_KeyDown({
    param($sender, $e)
    if ($e.Key -eq [System.Windows.Input.Key]::Enter) {
        Load-Directory -Path $txtPath.Text
        $e.Handled = $true
    }
})

$txtFilter.Add_TextChanged({ Apply-Filter -FilterText $txtFilter.Text })
$txtMinSizeMB.Add_TextChanged({ Apply-Filter -FilterText $txtFilter.Text })
$chkShowSystem.Add_Click({ Apply-Filter -FilterText $txtFilter.Text })

$btnClearFilter.Add_Click({
    $txtFilter.Text = ""
    Apply-Filter -FilterText ""
})

$btnDeleteSelected.Add_Click({ Delete-SelectedItems })
$btnExportCsv.Add_Click({ Export-CurrentViewToCsv })

$dgItems.Add_MouseDoubleClick({
    if ($dgItems.SelectedItem -ne $null) {
        $item = $dgItems.SelectedItem
        if ($item.ItemType -eq 'Directory') {
            Load-Directory -Path $item.FullPath
        } else {
            Open-ItemPath -FullPath $item.FullPath
        }
    }
})

$miOpen.Add_Click({
    if ($dgItems.SelectedItem) {
        $i = $dgItems.SelectedItem
        if ($i.ItemType -eq 'Directory') { Load-Directory -Path $i.FullPath }
        else { Open-ItemPath -FullPath $i.FullPath }
    }
})

$miOpenFolder.Add_Click({
    if ($dgItems.SelectedItem) { Open-ContainingFolder -FullPath $dgItems.SelectedItem.FullPath }
})

$miCopyPaths.Add_Click({ Copy-SelectedPaths })
$miDelete.Add_Click({ Delete-SelectedItems })

$window.Add_KeyDown({
    param($sender, $e)
    $key = $e.Key

    if ($key -eq [System.Windows.Input.Key]::F5) {
        Load-Directory -Path $txtPath.Text
        $e.Handled = $true
    }
    elseif ($key -eq [System.Windows.Input.Key]::Delete -and $dgItems.IsKeyboardFocusWithin) {
        Delete-SelectedItems
        $e.Handled = $true
    }
    elseif ($key -eq [System.Windows.Input.Key]::Back -and $dgItems.IsKeyboardFocusWithin) {
        Go-UpDirectory
        $e.Handled = $true
    }
})

# Quick paths
Add-QuickItem -combo $cmbQuick -label "Desktop"      -path ([Environment]::GetFolderPath('Desktop'))
Add-QuickItem -combo $cmbQuick -label "Documents"    -path ([Environment]::GetFolderPath('MyDocuments'))
Add-QuickItem -combo $cmbQuick -label "Downloads"    -path ([Environment]::GetFolderPath('UserProfile') + '\Downloads')
Add-QuickItem -combo $cmbQuick -label "System drive" -path ([System.IO.Path]::GetPathRoot([Environment]::GetFolderPath('System')))

$cmbQuick.Add_SelectionChanged({
    param($sender, $e)
    $item = $sender.SelectedItem
    if ($item -and $item.Tag) {
        $p = [string]$item.Tag
        if (Test-Path -LiteralPath $p -PathType Container) {
            $txtPath.Text = $p
            Load-Directory -Path $p
        }
    }
})

# Profiles
$noProf         = New-Object System.Windows.Controls.ComboBoxItem
$noProf.Content = "(No profile)"
$noProf.Tag     = $null
[void]$cmbProfile.Items.Add($noProf)

$userProfile  = [Environment]::GetFolderPath('UserProfile')
$downloadsDir = Join-Path $userProfile 'Downloads'
$systemRoot   = [System.IO.Path]::GetPathRoot([Environment]::GetFolderPath('System'))

Add-ProfileItem -combo $cmbProfile -label "User profile - big files (>20MB)"    -path $userProfile  -recurse $true -minSizeMB 20
Add-ProfileItem -combo $cmbProfile -label "Downloads - big files (>50MB)"       -path $downloadsDir -recurse $true -minSizeMB 50
Add-ProfileItem -combo $cmbProfile -label "System drive - very large (>200MB)"  -path $systemRoot   -recurse $true -minSizeMB 200

$cmbProfile.SelectedIndex = 0

$cmbProfile.Add_SelectionChanged({
    param($sender, $e)

    $item = $sender.SelectedItem
    if (-not $item) { return }

    if (-not $item.Tag) {
        return
    }

    $prof = $item.Tag

    if (-not (Test-Path -LiteralPath $prof.Path -PathType Container)) {
        [System.Windows.MessageBox]::Show(
            "Profile path is not available:`n`n$($prof.Path)",
            "Profile Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        ) | Out-Null
        return
    }

    $txtPath.Text         = $prof.Path
    $chkRecurse.IsChecked = $prof.Recurse
    $txtMinSizeMB.Text    = [string]$prof.MinSizeMB
    $txtFilter.Text       = ""

    Load-Directory -Path $prof.Path
})

# Initial state
$chkShowSystem.IsChecked = $false
$defaultPath = $env:USERPROFILE
if (Test-Path -LiteralPath $defaultPath) {
    $txtPath.Text = $defaultPath
    Load-Directory -Path $defaultPath
} else {
    Set-Status "Ready."
}

$window.Topmost = $false
$window.ShowDialog() | Out-Null
