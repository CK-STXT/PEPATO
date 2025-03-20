# Ensure PowerShell is running in Single Threaded Apartment (STA) mode
# Needed for WPF GUI to function properly
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne "STA") {
    Start-Process PowerShell -ArgumentList "-sta -File `"$PSCommandPath`"" -NoNewWindow -Wait
    exit
}

# Load WPF (Windows Presentation Framework) for GUI support
Add-Type -AssemblyName PresentationFramework

# Check if script is running as Administrator (needed for firewall and registry commands)
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    [System.Windows.MessageBox]::Show("This script must be run as administrator!", "Error", "OK", "Error")
    exit
}

# Define the XAML layout for the GUI
[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
        Title="PerPac Tool" Height="350" Width="500" ResizeMode="NoResize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Firewall Section -->
        <GroupBox Header="Firewall" Grid.Row="0" Margin="10">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                <Button Name="btnFirewallSnapshotBefore" Content="Snapshot Before" Width="120" Margin="5"/>
                <Button Name="btnFirewallSnapshotAfter" Content="Snapshot After" Width="120" Margin="5"/>
                <Button Name="btnCompareFirewall" Content="Compare" Width="100" Margin="5"/>
            </StackPanel>
        </GroupBox>

        <!-- Registry Section -->
        <GroupBox Header="Registry" Grid.Row="1" Margin="10">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                <Button Name="btnRegistrySnapshotBefore" Content="Snapshot Before" Width="120" Margin="5"/>
                <Button Name="btnRegistrySnapshotAfter" Content="Snapshot After" Width="120" Margin="5"/>
                <Button Name="btnCompareRegistry" Content="Compare" Width="100" Margin="5"/>
            </StackPanel>
        </GroupBox>

        <!-- Text Output for status messages -->
        <TextBox Name="txtOutput" Width="450" Height="80" Margin="10" Grid.Row="2" IsReadOnly="True" VerticalAlignment="Top" HorizontalAlignment="Center" TextWrapping="Wrap"/>

        <!-- Bottom Buttons -->
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Center">
            <Button Name="btnOpenFolder" Content="Open Folder" Width="120" Margin="10"/>
            <Button Name="btnClose" Content="Close" Width="120" Margin="10"/>
        </StackPanel>
    </Grid>
</Window>
"@

# Load the WPF GUI from the XAML definition
$reader = (New-Object System.Xml.XmlNodeReader $XAML)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Find UI Elements by Name
$btnFirewallSnapshotBefore = $window.FindName("btnFirewallSnapshotBefore")
$btnFirewallSnapshotAfter = $window.FindName("btnFirewallSnapshotAfter")
$btnCompareFirewall = $window.FindName("btnCompareFirewall")
$btnRegistrySnapshotBefore = $window.FindName("btnRegistrySnapshotBefore")
$btnRegistrySnapshotAfter = $window.FindName("btnRegistrySnapshotAfter")
$btnCompareRegistry = $window.FindName("btnCompareRegistry")
$btnOpenFolder = $window.FindName("btnOpenFolder")
$btnClose = $window.FindName("btnClose")
$txtOutput = $window.FindName("txtOutput")

# Define File Paths for Snapshots
$snapshotFolder = "$env:TEMP\PerPacTool"
$snapshotBeforeFirewall = "$snapshotFolder\FirewallSnapshot_Before.txt"
$snapshotAfterFirewall = "$snapshotFolder\FirewallSnapshot_After.txt"
$snapshotBeforeRegistry = "$snapshotFolder\RegistrySnapshot_Before.reg"
$snapshotAfterRegistry = "$snapshotFolder\RegistrySnapshot_After.reg"

# Ensure that the snapshot folder exists
if (!(Test-Path $snapshotFolder)) { New-Item -ItemType Directory -Path $snapshotFolder -Force }

# Function to update the output message in the GUI
function Update-Output {
    param($text)
    $window.Dispatcher.Invoke([action] { $txtOutput.Text = $text })
}

# Function: Take Firewall Snapshot (Synchronous Execution)
function Take-FirewallSnapshot {
    param($snapshotPath, $type)
    
    # Display a "Started..." message
    Update-Output "Firewall snapshot ($type) started..."
    
    # Execute the command to take a firewall snapshot and save to file
    netsh advfirewall firewall show rule name=all | Out-File -FilePath $snapshotPath -Encoding UTF8

    # Wait until the file exists (ensures snapshot completed before moving on)
    do { Start-Sleep -Milliseconds 500 } while (!(Test-Path $snapshotPath))

    # Display "Saved to..." message
    Update-Output "Firewall snapshot ($type) saved to: $snapshotPath"
}

# Function: Take Registry Snapshot (Synchronous Execution)
function Take-RegistrySnapshot {
    param($snapshotPath, $type)
    
    # Display a "Started..." message
    Update-Output "Registry snapshot ($type) started..."
    
    # Execute the command to take a registry snapshot and save to file
    reg export "HKLM\Software" $snapshotPath /y

    # Wait until the file exists (ensures snapshot completed before moving on)
    do { Start-Sleep -Milliseconds 500 } while (!(Test-Path $snapshotPath))

    # Display "Saved to..." message
    Update-Output "Registry snapshot ($type) saved to: $snapshotPath"
}

# Button Click Events
$btnFirewallSnapshotBefore.Add_Click({ Take-FirewallSnapshot $snapshotBeforeFirewall "Before" })
$btnFirewallSnapshotAfter.Add_Click({ Take-FirewallSnapshot $snapshotAfterFirewall "After" })
$btnRegistrySnapshotBefore.Add_Click({ Take-RegistrySnapshot $snapshotBeforeRegistry "Before" })
$btnRegistrySnapshotAfter.Add_Click({ Take-RegistrySnapshot $snapshotAfterRegistry "After" })
$btnOpenFolder.Add_Click({ Start-Process explorer.exe -ArgumentList $snapshotFolder })  # Opens the snapshot folder
$btnClose.Add_Click({ $window.Close() })  # Closes the GUI window

# Show the GUI
$window.ShowDialog()
