<#
.SYNOPSIS
    AD Computer Windows Update Scanner with GUI
.DESCRIPTION
    Queries Active Directory computers for Windows Update status including completed and available updates.
    Provides a GUI to select OUs and view results.
.NOTES
    Author: Claude
    Requires: ActiveDirectory PowerShell Module
    Requires: Run with appropriate AD permissions
#>

#Requires -Modules ActiveDirectory

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# XAML for GUI
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="AD Computer Windows Update Scanner"
    Height="700"
    Width="1200"
    WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="MinWidth" Value="100"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Padding" Value="5"/>
        </Style>
        <Style TargetType="Label">
            <Setter Property="Margin" Value="5,5,5,0"/>
        </Style>
    </Window.Resources>
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- OU Selection Section -->
        <GroupBox Header="Active Directory OU Selection" Grid.Row="0" Padding="10">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <Label Grid.Row="0" Grid.Column="0" Content="Domain:"/>
                <TextBox Name="txtDomain" Grid.Row="0" Grid.Column="1" IsReadOnly="True"/>

                <Label Grid.Row="1" Grid.Column="0" Content="Selected OU:"/>
                <TextBox Name="txtSelectedOU" Grid.Row="1" Grid.Column="1" IsReadOnly="True"/>
                <Button Name="btnBrowseOU" Grid.Row="1" Grid.Column="2" Content="Browse OU..."/>

                <CheckBox Name="chkIncludeSubOUs" Grid.Row="2" Grid.Column="1"
                          Content="Include Sub-OUs" IsChecked="True" Margin="5"/>
            </Grid>
        </GroupBox>

        <!-- Scan Options Section -->
        <GroupBox Header="Scan Options" Grid.Row="1" Padding="10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <Label Grid.Column="0" Content="Max Concurrent Scans:"/>
                <TextBox Name="txtMaxThreads" Grid.Column="1" Text="10" Width="60" HorizontalAlignment="Left"/>
                <Button Name="btnScan" Grid.Column="2" Content="Start Scan" FontWeight="Bold"/>
                <Button Name="btnStop" Grid.Column="3" Content="Stop Scan" IsEnabled="False"/>
            </Grid>
        </GroupBox>

        <!-- Results Grid -->
        <GroupBox Header="Scan Results" Grid.Row="2" Padding="5" Margin="0,5,0,5">
            <DataGrid Name="dgResults" AutoGenerateColumns="False" IsReadOnly="True"
                      CanUserResizeColumns="True" CanUserSortColumns="True"
                      GridLinesVisibility="All" AlternatingRowBackground="LightGray">
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Computer Name" Binding="{Binding ComputerName}" Width="150"/>
                    <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="100"/>
                    <DataGridTextColumn Header="OS Version" Binding="{Binding OSVersion}" Width="200"/>
                    <DataGridTextColumn Header="Last Boot" Binding="{Binding LastBoot}" Width="140"/>
                    <DataGridTextColumn Header="Installed Updates" Binding="{Binding InstalledCount}" Width="120"/>
                    <DataGridTextColumn Header="Available Updates" Binding="{Binding AvailableCount}" Width="120"/>
                    <DataGridTextColumn Header="Pending Reboot" Binding="{Binding PendingReboot}" Width="100"/>
                    <DataGridTextColumn Header="Error" Binding="{Binding Error}" Width="*"/>
                </DataGrid.Columns>
            </DataGrid>
        </GroupBox>

        <!-- Status and Progress Section -->
        <GroupBox Header="Progress" Grid.Row="3" Padding="10">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <TextBlock Name="txtStatus" Grid.Row="0" Text="Ready" Margin="0,0,0,5"/>
                <ProgressBar Name="progressBar" Grid.Row="1" Height="25" Minimum="0" Maximum="100" Value="0"/>
            </Grid>
        </GroupBox>

        <!-- Action Buttons -->
        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="btnExportCSV" Content="Export to CSV" IsEnabled="False"/>
            <Button Name="btnClear" Content="Clear Results"/>
            <Button Name="btnClose" Content="Close"/>
        </StackPanel>
    </Grid>
</Window>
"@

# Load XAML
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get controls
$txtDomain = $window.FindName("txtDomain")
$txtSelectedOU = $window.FindName("txtSelectedOU")
$btnBrowseOU = $window.FindName("btnBrowseOU")
$chkIncludeSubOUs = $window.FindName("chkIncludeSubOUs")
$txtMaxThreads = $window.FindName("txtMaxThreads")
$btnScan = $window.FindName("btnScan")
$btnStop = $window.FindName("btnStop")
$dgResults = $window.FindName("dgResults")
$txtStatus = $window.FindName("txtStatus")
$progressBar = $window.FindName("progressBar")
$btnExportCSV = $window.FindName("btnExportCSV")
$btnClear = $window.FindName("btnClear")
$btnClose = $window.FindName("btnClose")

# Initialize results collection
$script:results = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
$dgResults.ItemsSource = $script:results
$script:stopScan = $false

# Get current domain
try {
    $domain = Get-ADDomain
    $txtDomain.Text = $domain.DNSRoot
    $txtSelectedOU.Text = $domain.DistinguishedName
} catch {
    [System.Windows.MessageBox]::Show("Failed to connect to Active Directory. Please ensure you have the ActiveDirectory module installed and appropriate permissions.`n`nError: $_",
                                       "AD Connection Error",
                                       [System.Windows.MessageBoxButton]::OK,
                                       [System.Windows.MessageBoxImage]::Error)
    $window.Close()
    return
}

# Function to browse OU
function Show-OUBrowser {
    [xml]$ouXaml = @"
    <Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Select Organizational Unit"
        Height="500"
        Width="600"
        WindowStartupLocation="CenterOwner">
        <Grid Margin="10">
            <Grid.RowDefinitions>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <TreeView Name="tvOUs" Grid.Row="0" Margin="0,0,0,10"/>

            <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="btnOK" Content="OK" Width="80" Margin="5" IsDefault="True"/>
                <Button Name="btnCancel" Content="Cancel" Width="80" Margin="5" IsCancel="True"/>
            </StackPanel>
        </Grid>
    </Window>
"@

    $ouReader = New-Object System.Xml.XmlNodeReader $ouXaml
    $ouWindow = [Windows.Markup.XamlReader]::Load($ouReader)
    $ouWindow.Owner = $window

    $tvOUs = $ouWindow.FindName("tvOUs")
    $btnOK = $ouWindow.FindName("btnOK")
    $btnCancel = $ouWindow.FindName("btnCancel")

    # Populate OU tree
    try {
        $rootOU = Get-ADDomain
        $rootItem = New-Object System.Windows.Controls.TreeViewItem
        $rootItem.Header = $rootOU.DNSRoot
        $rootItem.Tag = $rootOU.DistinguishedName
        $rootItem.IsExpanded = $true

        # Add dummy item for lazy loading
        $dummyItem = New-Object System.Windows.Controls.TreeViewItem
        $dummyItem.Header = "Loading..."
        $rootItem.Items.Add($dummyItem) | Out-Null

        $tvOUs.Items.Add($rootItem) | Out-Null

        # Lazy load OUs
        $rootItem.add_Expanded({
            param($sender, $e)
            if ($sender.Items.Count -eq 1 -and $sender.Items[0].Header -eq "Loading...") {
                $sender.Items.Clear()
                $ous = Get-ADOrganizationalUnit -Filter * -SearchBase $sender.Tag -SearchScope OneLevel | Sort-Object Name
                foreach ($ou in $ous) {
                    $ouItem = New-Object System.Windows.Controls.TreeViewItem
                    $ouItem.Header = $ou.Name
                    $ouItem.Tag = $ou.DistinguishedName

                    # Check if this OU has children
                    $hasChildren = Get-ADOrganizationalUnit -Filter * -SearchBase $ou.DistinguishedName -SearchScope OneLevel -ErrorAction SilentlyContinue
                    if ($hasChildren) {
                        $dummy = New-Object System.Windows.Controls.TreeViewItem
                        $dummy.Header = "Loading..."
                        $ouItem.Items.Add($dummy) | Out-Null
                    }

                    $sender.Items.Add($ouItem) | Out-Null

                    # Add expand event recursively
                    $ouItem.add_Expanded({
                        param($s, $e)
                        if ($s.Items.Count -eq 1 -and $s.Items[0].Header -eq "Loading...") {
                            $s.Items.Clear()
                            $subOUs = Get-ADOrganizationalUnit -Filter * -SearchBase $s.Tag -SearchScope OneLevel | Sort-Object Name
                            foreach ($subOU in $subOUs) {
                                $subItem = New-Object System.Windows.Controls.TreeViewItem
                                $subItem.Header = $subOU.Name
                                $subItem.Tag = $subOU.DistinguishedName

                                $hasSubChildren = Get-ADOrganizationalUnit -Filter * -SearchBase $subOU.DistinguishedName -SearchScope OneLevel -ErrorAction SilentlyContinue
                                if ($hasSubChildren) {
                                    $d = New-Object System.Windows.Controls.TreeViewItem
                                    $d.Header = "Loading..."
                                    $subItem.Items.Add($d) | Out-Null
                                }

                                $s.Items.Add($subItem) | Out-Null
                            }
                        }
                        $e.Handled = $true
                    })
                }
            }
            $e.Handled = $true
        })

    } catch {
        [System.Windows.MessageBox]::Show("Error loading OUs: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }

    $btnOK.add_Click({
        $selectedItem = $tvOUs.SelectedItem
        if ($selectedItem -and $selectedItem.Tag) {
            $ouWindow.Tag = $selectedItem.Tag
            $ouWindow.DialogResult = $true
            $ouWindow.Close()
        } else {
            [System.Windows.MessageBox]::Show("Please select an OU.", "Selection Required", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        }
    })

    $btnCancel.add_Click({ $ouWindow.Close() })

    $result = $ouWindow.ShowDialog()
    if ($result) {
        return $ouWindow.Tag
    }
    return $null
}

# Function to get Windows Update information from remote computer
function Get-WindowsUpdateInfo {
    param(
        [string]$ComputerName
    )

    $result = [PSCustomObject]@{
        ComputerName = $ComputerName
        Status = "Unknown"
        OSVersion = ""
        LastBoot = ""
        InstalledCount = 0
        AvailableCount = 0
        PendingReboot = "Unknown"
        Error = ""
    }

    try {
        # Test connectivity
        if (-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
            $result.Status = "Offline"
            $result.Error = "Unable to ping computer"
            return $result
        }

        $result.Status = "Online"

        # Get OS Info and Last Boot Time
        try {
            $os = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction Stop
            $result.OSVersion = $os.Caption + " " + $os.Version
            $result.LastBoot = $os.ConvertToDateTime($os.LastBootUpTime).ToString("yyyy-MM-dd HH:mm:ss")
        } catch {
            $result.Error += "OS Info: $($_.Exception.Message); "
        }

        # Check for pending reboot
        try {
            $rebootPending = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                $reboot = $false

                # Check Component Based Servicing
                if (Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending") {
                    $reboot = $true
                }

                # Check Windows Update
                if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") {
                    $reboot = $true
                }

                # Check PendingFileRenameOperations
                $prop = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
                if ($prop) {
                    $reboot = $true
                }

                return $reboot
            } -ErrorAction Stop

            $result.PendingReboot = if ($rebootPending) { "Yes" } else { "No" }
        } catch {
            $result.PendingReboot = "Unknown"
        }

        # Get Windows Update information
        try {
            $updateInfo = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                $updateSession = New-Object -ComObject Microsoft.Update.Session
                $updateSearcher = $updateSession.CreateUpdateSearcher()

                # Search for installed updates (last 90 days)
                $installedCount = 0
                try {
                    $historyCount = $updateSearcher.GetTotalHistoryCount()
                    if ($historyCount -gt 0) {
                        $history = $updateSearcher.QueryHistory(0, [Math]::Min($historyCount, 1000))
                        $installedCount = ($history | Where-Object { $_.Operation -eq 1 }).Count
                    }
                } catch {
                    $installedCount = -1
                }

                # Search for available updates
                $availableCount = 0
                try {
                    $searchResult = $updateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
                    $availableCount = $searchResult.Updates.Count
                } catch {
                    $availableCount = -1
                }

                return @{
                    InstalledCount = $installedCount
                    AvailableCount = $availableCount
                }
            } -ErrorAction Stop

            $result.InstalledCount = if ($updateInfo.InstalledCount -ge 0) { $updateInfo.InstalledCount } else { "Error" }
            $result.AvailableCount = if ($updateInfo.AvailableCount -ge 0) { $updateInfo.AvailableCount } else { "Error" }

        } catch {
            $result.Error += "Update Query: $($_.Exception.Message); "
            $result.InstalledCount = "Error"
            $result.AvailableCount = "Error"
        }

    } catch {
        $result.Status = "Error"
        $result.Error += $_.Exception.Message
    }

    return $result
}

# Browse OU button click
$btnBrowseOU.add_Click({
    $selectedOU = Show-OUBrowser
    if ($selectedOU) {
        $txtSelectedOU.Text = $selectedOU
    }
})

# Start Scan button click
$btnScan.add_Click({
    if ([string]::IsNullOrWhiteSpace($txtSelectedOU.Text)) {
        [System.Windows.MessageBox]::Show("Please select an OU first.", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }

    # Validate max threads
    $maxThreads = 10
    if (-not [int]::TryParse($txtMaxThreads.Text, [ref]$maxThreads) -or $maxThreads -lt 1 -or $maxThreads -gt 50) {
        [System.Windows.MessageBox]::Show("Max Concurrent Scans must be between 1 and 50.", "Validation", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }

    # Clear previous results
    $script:results.Clear()
    $script:stopScan = $false

    # Disable/Enable buttons
    $btnScan.IsEnabled = $false
    $btnStop.IsEnabled = $true
    $btnBrowseOU.IsEnabled = $false
    $btnExportCSV.IsEnabled = $false
    $progressBar.Value = 0

    # Get computers from AD
    $searchBase = $txtSelectedOU.Text
    $searchScope = if ($chkIncludeSubOUs.IsChecked) { "Subtree" } else { "OneLevel" }

    $txtStatus.Text = "Retrieving computers from AD..."

    # Run scan in background
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("searchBase", $searchBase)
    $runspace.SessionStateProxy.SetVariable("searchScope", $searchScope)
    $runspace.SessionStateProxy.SetVariable("maxThreads", $maxThreads)
    $runspace.SessionStateProxy.SetVariable("window", $window)
    $runspace.SessionStateProxy.SetVariable("results", $script:results)
    $runspace.SessionStateProxy.SetVariable("stopScan", $script:stopScan)

    $powershell = [powershell]::Create()
    $powershell.Runspace = $runspace

    [void]$powershell.AddScript({
        param($sb, $ss, $mt)

        # Import required functions into runspace
        function Get-WindowsUpdateInfo {
            param([string]$ComputerName)

            $result = [PSCustomObject]@{
                ComputerName = $ComputerName
                Status = "Unknown"
                OSVersion = ""
                LastBoot = ""
                InstalledCount = 0
                AvailableCount = 0
                PendingReboot = "Unknown"
                Error = ""
            }

            try {
                if (-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
                    $result.Status = "Offline"
                    $result.Error = "Unable to ping computer"
                    return $result
                }

                $result.Status = "Online"

                try {
                    $os = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction Stop
                    $result.OSVersion = $os.Caption + " " + $os.Version
                    $result.LastBoot = $os.ConvertToDateTime($os.LastBootUpTime).ToString("yyyy-MM-dd HH:mm:ss")
                } catch {
                    $result.Error += "OS Info: $($_.Exception.Message); "
                }

                try {
                    $rebootPending = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                        $reboot = $false
                        if (Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending") { $reboot = $true }
                        if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") { $reboot = $true }
                        $prop = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
                        if ($prop) { $reboot = $true }
                        return $reboot
                    } -ErrorAction Stop

                    $result.PendingReboot = if ($rebootPending) { "Yes" } else { "No" }
                } catch {
                    $result.PendingReboot = "Unknown"
                }

                try {
                    $updateInfo = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                        $updateSession = New-Object -ComObject Microsoft.Update.Session
                        $updateSearcher = $updateSession.CreateUpdateSearcher()

                        $installedCount = 0
                        try {
                            $historyCount = $updateSearcher.GetTotalHistoryCount()
                            if ($historyCount -gt 0) {
                                $history = $updateSearcher.QueryHistory(0, [Math]::Min($historyCount, 1000))
                                $installedCount = ($history | Where-Object { $_.Operation -eq 1 }).Count
                            }
                        } catch { $installedCount = -1 }

                        $availableCount = 0
                        try {
                            $searchResult = $updateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
                            $availableCount = $searchResult.Updates.Count
                        } catch { $availableCount = -1 }

                        return @{ InstalledCount = $installedCount; AvailableCount = $availableCount }
                    } -ErrorAction Stop

                    $result.InstalledCount = if ($updateInfo.InstalledCount -ge 0) { $updateInfo.InstalledCount } else { "Error" }
                    $result.AvailableCount = if ($updateInfo.AvailableCount -ge 0) { $updateInfo.AvailableCount } else { "Error" }
                } catch {
                    $result.Error += "Update Query: $($_.Exception.Message); "
                    $result.InstalledCount = "Error"
                    $result.AvailableCount = "Error"
                }
            } catch {
                $result.Status = "Error"
                $result.Error += $_.Exception.Message
            }

            return $result
        }

        try {
            Import-Module ActiveDirectory -ErrorAction Stop

            $computers = Get-ADComputer -Filter * -SearchBase $sb -SearchScope $ss -Properties Name | Select-Object -ExpandProperty Name | Sort-Object

            if ($computers.Count -eq 0) {
                $window.Dispatcher.Invoke([action]{
                    $window.FindName("txtStatus").Text = "No computers found in selected OU"
                    $window.FindName("btnScan").IsEnabled = $true
                    $window.FindName("btnStop").IsEnabled = $false
                    $window.FindName("btnBrowseOU").IsEnabled = $true
                })
                return
            }

            $window.Dispatcher.Invoke([action]{
                $window.FindName("txtStatus").Text = "Found $($computers.Count) computers. Starting scan..."
            })

            $jobs = @()
            $completed = 0
            $total = $computers.Count

            foreach ($computer in $computers) {
                # Check if scan should stop
                if ($script:stopScan) { break }

                # Wait if we've hit max concurrent jobs
                while (($jobs | Where-Object { $_.State -eq 'Running' }).Count -ge $mt) {
                    Start-Sleep -Milliseconds 100

                    # Collect completed jobs
                    $completedJobs = $jobs | Where-Object { $_.State -ne 'Running' }
                    foreach ($job in $completedJobs) {
                        $jobResult = Receive-Job -Job $job
                        Remove-Job -Job $job

                        $window.Dispatcher.Invoke([action]{
                            $results.Add($jobResult)
                        })

                        $jobs = $jobs | Where-Object { $_.Id -ne $job.Id }
                        $completed++

                        $window.Dispatcher.Invoke([action]{
                            $progress = [math]::Round(($completed / $total) * 100)
                            $window.FindName("progressBar").Value = $progress
                            $window.FindName("txtStatus").Text = "Scanned $completed of $total computers..."
                        })
                    }
                }

                # Start new job
                $job = Start-Job -ScriptBlock {
                    param($comp, $funcDef)

                    # Define function in job scope
                    Invoke-Expression $funcDef

                    return Get-WindowsUpdateInfo -ComputerName $comp
                } -ArgumentList $computer, ${function:Get-WindowsUpdateInfo}.ToString()

                $jobs += $job
            }

            # Wait for remaining jobs
            while ($jobs.Count -gt 0) {
                $completedJobs = $jobs | Where-Object { $_.State -ne 'Running' }
                foreach ($job in $completedJobs) {
                    $jobResult = Receive-Job -Job $job
                    Remove-Job -Job $job

                    $window.Dispatcher.Invoke([action]{
                        $results.Add($jobResult)
                    })

                    $jobs = $jobs | Where-Object { $_.Id -ne $job.Id }
                    $completed++

                    $window.Dispatcher.Invoke([action]{
                        $progress = [math]::Round(($completed / $total) * 100)
                        $window.FindName("progressBar").Value = $progress
                        $window.FindName("txtStatus").Text = "Scanned $completed of $total computers..."
                    })
                }

                if ($jobs.Count -gt 0) {
                    Start-Sleep -Milliseconds 100
                }
            }

            $window.Dispatcher.Invoke([action]{
                $window.FindName("txtStatus").Text = "Scan complete. Scanned $total computers."
                $window.FindName("progressBar").Value = 100
                $window.FindName("btnScan").IsEnabled = $true
                $window.FindName("btnStop").IsEnabled = $false
                $window.FindName("btnBrowseOU").IsEnabled = $true
                $window.FindName("btnExportCSV").IsEnabled = $true
            })

        } catch {
            $window.Dispatcher.Invoke([action]{
                $window.FindName("txtStatus").Text = "Error: $($_.Exception.Message)"
                $window.FindName("btnScan").IsEnabled = $true
                $window.FindName("btnStop").IsEnabled = $false
                $window.FindName("btnBrowseOU").IsEnabled = $true
            })
        }
    }).AddArgument($searchBase).AddArgument($searchScope).AddArgument($maxThreads)

    $asyncResult = $powershell.BeginInvoke()
})

# Stop Scan button click
$btnStop.add_Click({
    $script:stopScan = $true
    $txtStatus.Text = "Stopping scan..."
    $btnStop.IsEnabled = $false
})

# Export CSV button click
$btnExportCSV.add_Click({
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveDialog.FileName = "WindowsUpdateScan_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $script:results | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
            [System.Windows.MessageBox]::Show("Results exported successfully to:`n$($saveDialog.FileName)",
                                               "Export Complete",
                                               [System.Windows.MessageBoxButton]::OK,
                                               [System.Windows.MessageBoxImage]::Information)
        } catch {
            [System.Windows.MessageBox]::Show("Error exporting results: $_",
                                               "Export Error",
                                               [System.Windows.MessageBoxButton]::OK,
                                               [System.Windows.MessageBoxImage]::Error)
        }
    }
})

# Clear Results button click
$btnClear.add_Click({
    $script:results.Clear()
    $progressBar.Value = 0
    $txtStatus.Text = "Ready"
    $btnExportCSV.IsEnabled = $false
})

# Close button click
$btnClose.add_Click({ $window.Close() })

# Show the window
$window.ShowDialog() | Out-Null
