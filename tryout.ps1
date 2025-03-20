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
        <TextBox Name="txtOutput" Width="450" Height="80" Margin="10" Grid.Row="2" IsReadOnly="True" VerticalAlignment="Top" HorizontalAlignment="Center" TextWrapping="Wrap"/>

        <!-- Bottom Buttons -->
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Center">
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

# Define File Paths
$snapshotFolder = "$env:TEMP\PerPacTool"
$snapshotBeforeFirewall = "$snapshotFolder\FirewallSnapshot_Before.txt"
$snapshotAfterFirewall = "$snapshotFolder\FirewallSnapshot_After.txt"
$snapshotBeforeRegistry = "$snapshotFolder\RegistrySnapshot_Before.reg"
$snapshotAfterRegistry = "$snapshotFolder\RegistrySnapshot_After.reg"
$diffFirewallPath = "$snapshotFolder\Firewall_Differences.txt"
$diffRegistryPath = "$snapshotFolder\Registry_Differences.txt"

# Ensure snapshot folder exists
if (!(Test-Path $snapshotFolder)) { New-Item -ItemType Directory -Path $snapshotFolder -Force }

# Function: Update Text Output
function Update-Output {
    param($text)
    $window.Dispatcher.Invoke([action]{
        $txtOutput.Text = $text
    })
}

# Function: Take Firewall Snapshot
function Take-FirewallSnapshot {
    param($snapshotPath, $type)
    Update-Output "Firewall snapshot ($type) started..."
    netsh advfirewall firewall show rule name=all | Out-File -FilePath $snapshotPath -Encoding UTF8
    Update-Output "Firewall snapshot ($type) saved to: $snapshotPath"
}

# Function: Compare Firewall Snapshots
function Compare-FirewallSnapshots {
    Update-Output "Firewall comparison started..."
    if (!(Test-Path $snapshotBeforeFirewall) -or !(Test-Path $snapshotAfterFirewall)) {
        Update-Output "Snapshots not found! Take both snapshots first."
        return
    }
    $diff = Compare-Object -ReferenceObject (Get-Content $snapshotBeforeFirewall) -DifferenceObject (Get-Content $snapshotAfterFirewall)
    if ($diff) {
        $diff | Out-File -FilePath $diffFirewallPath -Encoding UTF8
        Update-Output "Firewall comparison ended. Differences saved in: $diffFirewallPath"
    } else {
        Update-Output "Firewall comparison ended. No differences found."
    }
}

# Function: Take Registry Snapshot
function Take-RegistrySnapshot {
    param($snapshotPath, $type)
    Update-Output "Registry snapshot ($type) started..."
    reg export "HKLM\Software" $snapshotPath /y
    if (Test-Path $snapshotPath) {
        Update-Output "Registry snapshot ($type) saved to: $snapshotPath"
    } else {
        Update-Output "Failed to save registry snapshot!"
    }
}

# Function: Compare Registry Snapshots
function Compare-RegistrySnapshots {
    Update-Output "Registry comparison started..."
    if (!(Test-Path $snapshotBeforeRegistry) -or !(Test-Path $snapshotAfterRegistry)) {
        Update-Output "Snapshots not found! Take both snapshots first."
        return
    }
    $diff = Compare-Object -ReferenceObject (Get-Content $snapshotBeforeRegistry) -DifferenceObject (Get-Content $snapshotAfterRegistry)
    if ($diff) {
        $diff | Out-File -FilePath $diffRegistryPath -Encoding UTF8
        Update-Output "Registry comparison ended. Differences saved in: $diffRegistryPath"
    } else {
        Update-Output "Registry comparison ended. No differences found."
    }
}

# Button Events
$btnFirewallSnapshotBefore.Add_Click({ Take-FirewallSnapshot $snapshotBeforeFirewall "Before" })
$btnFirewallSnapshotAfter.Add_Click({ Take-FirewallSnapshot $snapshotAfterFirewall "After" })
$btnCompareFirewall.Add_Click({ Compare-FirewallSnapshots })

$btnRegistrySnapshotBefore.Add_Click({ Take-RegistrySnapshot $snapshotBeforeRegistry "Before" })
$btnRegistrySnapshotAfter.Add_Click({ Take-RegistrySnapshot $snapshotAfterRegistry "After" })
$btnCompareRegistry.Add_Click({ Compare-RegistrySnapshots })

$btnOpenFolder.Add_Click({ Start-Process explorer.exe -ArgumentList $snapshotFolder })
$btnClose.Add_Click({ $window.Close() })

# Show GUI
$window.ShowDialog()
