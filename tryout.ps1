# Ensure PowerShell is running in STA mode (needed for WPF)
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne "STA") {
    Start-Process PowerShell -ArgumentList "-sta -File `"$PSCommandPath`"" -NoNewWindow -Wait
    exit
}

# Load WPF Assembly
Add-Type -AssemblyName PresentationFramework

# Check if script is running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    [System.Windows.MessageBox]::Show("This script must be run as administrator!", "Error", "OK", "Error")
    exit
}

# XAML for WPF GUI
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

        <!-- Text Output -->
        <TextBox Name="txtOutput" Width="450" Height="60" Margin="10" Grid.Row="2" IsReadOnly="True" VerticalAlignment="Top" HorizontalAlignment="Center" TextWrapping="Wrap"/>

        <!-- Progress Bar -->
        <ProgressBar Name="progressBar" Width="450" Height="20" Grid.Row="3" Margin="10"/>

        <!-- Bottom Buttons -->
        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Center">
            <Button Name="btnOpenFolder" Content="Open Folder" Width="120" Margin="10"/>
            <Button Name="btnClose" Content="Close" Width="120" Margin="10"/>
        </StackPanel>
    </Grid>
</Window>
"@

# Load WPF
$reader = (New-Object System.Xml.XmlNodeReader $XAML)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Find UI Elements
$btnFirewallSnapshotBefore = $window.FindName("btnFirewallSnapshotBefore")
$btnFirewallSnapshotAfter = $window.FindName("btnFirewallSnapshotAfter")
$btnCompareFirewall = $window.FindName("btnCompareFirewall")
$btnRegistrySnapshotBefore = $window.FindName("btnRegistrySnapshotBefore")
$btnRegistrySnapshotAfter = $window.FindName("btnRegistrySnapshotAfter")
$btnCompareRegistry = $window.FindName("btnCompareRegistry")
$btnOpenFolder = $window.FindName("btnOpenFolder")
$btnClose = $window.FindName("btnClose")
$txtOutput = $window.FindName("txtOutput")
$progressBar = $window.FindName("progressBar")

# Define File Paths
$snapshotFolder = "$env:TEMP\PerPacTool"
$snapshotBeforeFirewall = "$snapshotFolder\FirewallSnapshot_Before.txt"
$snapshotAfterFirewall = "$snapshotFolder\FirewallSnapshot_After.txt"
$snapshotBeforeRegistry = "$snapshotFolder\RegistrySnapshot_Before.reg"
$snapshotAfterRegistry = "$snapshotFolder\RegistrySnapshot_After.reg"

# Ensure snapshot folder exists
if (!(Test-Path $snapshotFolder)) { New-Item -ItemType Directory -Path $snapshotFolder -Force }

# Function: Update UI elements dynamically
function Update-UI {
    param($text, $progress)
    $window.Dispatcher.Invoke([action]{
        $txtOutput.Text = $text
        $progressBar.Value = $progress
    })
}

# Function: Take Firewall Snapshot (Now Works!)
function Take-FirewallSnapshot {
    param($snapshotPath)
    Update-UI "Starting Firewall Snapshot..." 10
    Start-Sleep -Seconds 1
    netsh advfirewall firewall show rule name=all | Out-File -FilePath $snapshotPath -Encoding UTF8
    Update-UI "Firewall snapshot saved to: $snapshotPath" 100
    Start-Sleep -Seconds 1
    Update-UI "" 0
}

# Function: Take Registry Snapshot (Now Works & Uses reg export for Speed!)
function Take-RegistrySnapshot {
    param($snapshotPath)
    Update-UI "Starting Registry Snapshot... (This can take time)" 10
    Start-Sleep -Seconds 1
    reg export "HKLM\Software" $snapshotPath /y
    if (Test-Path $snapshotPath) {
        Update-UI "Registry snapshot saved to: $snapshotPath" 100
    } else {
        Update-UI "Failed to save registry snapshot!" 0
    }
    Start-Sleep -Seconds 1
    Update-UI "" 0
}

# Function: Compare Firewall Snapshots
function Compare-FirewallSnapshots {
    if (!(Test-Path $snapshotBeforeFirewall) -or !(Test-Path $snapshotAfterFirewall)) {
        Update-UI "Snapshots not found! Take both snapshots first." 0
        return
    }
    Update-UI "Comparing Firewall Snapshots..." 50
    $diff = Compare-Object -ReferenceObject (Get-Content $snapshotBeforeFirewall) -DifferenceObject (Get-Content $snapshotAfterFirewall)
    if ($diff) {
        $diff | Out-File -FilePath "$snapshotFolder\Firewall_Differences.txt" -Encoding UTF8
        Update-UI "Differences saved in Firewall_Differences.txt" 100
    } else {
        Update-UI "No differences found." 100
    }
    Start-Sleep -Seconds 1
    Update-UI "" 0
}

# Button Events
$btnFirewallSnapshotBefore.Add_Click({ Take-FirewallSnapshot $snapshotBeforeFirewall })
$btnFirewallSnapshotAfter.Add_Click({ Take-FirewallSnapshot $snapshotAfterFirewall })
$btnCompareFirewall.Add_Click({ Compare-FirewallSnapshots })

$btnRegistrySnapshotBefore.Add_Click({ Take-RegistrySnapshot $snapshotBeforeRegistry })
$btnRegistrySnapshotAfter.Add_Click({ Take-RegistrySnapshot $snapshotAfterRegistry })

$btnOpenFolder.Add_Click({ Start-Process explorer.exe -ArgumentList $snapshotFolder })
$btnClose.Add_Click({ $window.Close() })

# Show GUI
$window.ShowDialog()
