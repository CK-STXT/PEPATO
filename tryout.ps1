# Ensure PowerShell is running in Single Threaded Apartment (STA) mode
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne "STA") {
    Start-Process PowerShell -ArgumentList "-sta -File `"$PSCommandPath`"" -NoNewWindow -Wait
    exit
}

# Load WPF Assembly
Add-Type -AssemblyName PresentationFramework

# Check if script is running as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    [System.Windows.MessageBox]::Show("This script must be run as administrator!", "Error", "OK", "Error")
    exit
}

# Define Paths for Snapshots
$snapshotFolder = "$env:TEMP\PerPacTool"
$snapshotBeforeFirewall = "$snapshotFolder\FirewallSnapshot_Before.txt"
$snapshotAfterFirewall = "$snapshotFolder\FirewallSnapshot_After.txt"
$snapshotBeforeRegistry = "$snapshotFolder\RegistrySnapshot_Before.reg"
$snapshotAfterRegistry = "$snapshotFolder\RegistrySnapshot_After.reg"

# Ensure the snapshot folder exists
if (!(Test-Path $snapshotFolder)) { New-Item -ItemType Directory -Path $snapshotFolder -Force }

# Define the XAML layout for the GUI
[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
        Title="PerPac Tool (Running as SYSTEM)" Height="700" Width="600" ResizeMode="NoResize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
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

        <!-- Installer Handling Section -->
        <GroupBox Header="Installer Handling" Grid.Row="2" Margin="10">
            <StackPanel Orientation="Vertical" HorizontalAlignment="Center">
                <Button Name="btnSelectExe" Content="Select Installer (EXE)" Width="200" Margin="5"/>
                <TextBox Name="txtExePath" Width="450" Height="25" Margin="5" IsReadOnly="True"/>
                <Button Name="btnRunInstaller" Content="Run Installer (As SYSTEM)" Width="200" Margin="5"/>
                <Button Name="btnFindMSI" Content="Find Extracted MSIs" Width="200" Margin="5"/>
            </StackPanel>
        </GroupBox>

        <!-- Text Output for status messages -->
        <TextBox Name="txtOutput" Width="550" Height="100" Margin="10" Grid.Row="3" IsReadOnly="True" VerticalAlignment="Top" HorizontalAlignment="Center" TextWrapping="Wrap"/>

        <!-- Bottom Buttons -->
        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Center">
            <Button Name="btnOpenFolder" Content="Open Folder" Width="120" Margin="10"/>
            <Button Name="btnClose" Content="Close" Width="120" Margin="10"/>
        </StackPanel>
    </Grid>
</Window>
"@

# Load the WPF GUI
$reader = (New-Object System.Xml.XmlNodeReader $XAML)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Find UI Elements
$btnFirewallSnapshotBefore = $window.FindName("btnFirewallSnapshotBefore")
$btnFirewallSnapshotAfter = $window.FindName("btnFirewallSnapshotAfter")
$btnCompareFirewall = $window.FindName("btnCompareFirewall")
$btnRegistrySnapshotBefore = $window.FindName("btnRegistrySnapshotBefore")
$btnRegistrySnapshotAfter = $window.FindName("btnRegistrySnapshotAfter")
$btnCompareRegistry = $window.FindName("btnCompareRegistry")
$btnSelectExe = $window.FindName("btnSelectExe")
$btnRunInstaller = $window.FindName("btnRunInstaller")
$btnFindMSI = $window.FindName("btnFindMSI")
$btnOpenFolder = $window.FindName("btnOpenFolder")
$btnClose = $window.FindName("btnClose")
$txtExePath = $window.FindName("txtExePath")
$txtOutput = $window.FindName("txtOutput")

# Function: Update Output Text
function Update-Output {
    param($text)
    $window.Dispatcher.Invoke([action] { $txtOutput.Text = $text })
}

# Firewall Snapshot Function
function Take-FirewallSnapshot {
    param($snapshotPath, $type)

    Update-Output "Firewall snapshot ($type) started..."
    netsh advfirewall firewall show rule name=all | Out-File -FilePath $snapshotPath -Encoding UTF8
    Update-Output "Firewall snapshot ($type) saved to: $snapshotPath"
}

# Registry Snapshot Function
function Take-RegistrySnapshot {
    param($snapshotPath, $type)

    Update-Output "Registry snapshot ($type) started..."
    reg export "HKLM\Software" $snapshotPath /y
    Update-Output "Registry snapshot ($type) saved to: $snapshotPath"
}

# Function: Compare Firewall Snapshots
function Compare-FirewallSnapshots {
    Update-Output "Comparing Firewall snapshots..."
    $before = Get-Content $snapshotBeforeFirewall
    $after = Get-Content $snapshotAfterFirewall
    $diff = Compare-Object -ReferenceObject $before -DifferenceObject $after
    Update-Output "Firewall comparison complete!"
}

# Function: Compare Registry Snapshots
function Compare-RegistrySnapshots {
    Update-Output "Comparing Registry snapshots..."
    $before = Get-Content $snapshotBeforeRegistry
    $after = Get-Content $snapshotAfterRegistry
    $diff = Compare-Object -ReferenceObject $before -DifferenceObject $after
    Update-Output "Registry comparison complete!"
}

# Button Click Events
$btnFirewallSnapshotBefore.Add_Click({ Take-FirewallSnapshot $snapshotBeforeFirewall "Before" })
$btnFirewallSnapshotAfter.Add_Click({ Take-FirewallSnapshot $snapshotAfterFirewall "After" })
$btnCompareFirewall.Add_Click({ Compare-FirewallSnapshots })
$btnRegistrySnapshotBefore.Add_Click({ Take-RegistrySnapshot $snapshotBeforeRegistry "Before" })
$btnRegistrySnapshotAfter.Add_Click({ Take-RegistrySnapshot $snapshotAfterRegistry "After" })
$btnCompareRegistry.Add_Click({ Compare-RegistrySnapshots })
$btnSelectExe.Add_Click({ Select-Exe })
$btnRunInstaller.Add_Click({ Run-Installer })
$btnFindMSI.Add_Click({ Find-MSI })
$btnOpenFolder.Add_Click({ Start-Process explorer.exe -ArgumentList $snapshotFolder })
$btnClose.Add_Click({ $window.Close() })

# Show the GUI
$window.ShowDialog()
