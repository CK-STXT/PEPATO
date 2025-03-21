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
        Title="PerPac Tool (Running as SYSTEM)" Height="750" Width="600" ResizeMode="NoResize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
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

        <!-- MSI Analysis Section -->
        <GroupBox Header="MSI Analysis" Grid.Row="3" Margin="10">
            <StackPanel Orientation="Vertical" HorizontalAlignment="Center">
                <Button Name="btnSelectMSI" Content="Select MSI" Width="200" Margin="5"/>
                <TextBox Name="txtMSIPath" Width="450" Height="25" Margin="5" IsReadOnly="True"/>
                <Button Name="btnAnalyzeMSI" Content="Analyze MSI" Width="200" Margin="5"/>
            </StackPanel>
        </GroupBox>

        <!-- Text Output for status messages -->
        <TextBox Name="txtOutput" Width="550" Height="120" Margin="10" Grid.Row="4" IsReadOnly="True" VerticalAlignment="Top" HorizontalAlignment="Center" TextWrapping="Wrap"/>

        <!-- Bottom Buttons -->
        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Center">
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
$btnSelectMSI = $window.FindName("btnSelectMSI")
$btnAnalyzeMSI = $window.FindName("btnAnalyzeMSI")
$txtMSIPath = $window.FindName("txtMSIPath")
$txtOutput = $window.FindName("txtOutput")

# Function: Update Output Text
function Update-Output {
    param($text)
    $window.Dispatcher.Invoke([action] { $txtOutput.Text = $text })
}

# Function: Select MSI File
function Select-MSI {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "MSI Files (*.msi)|*.msi"
    $dialog.ShowDialog() | Out-Null
    if ($dialog.FileName) {
        $txtMSIPath.Text = $dialog.FileName
        Update-Output "Selected MSI: $($dialog.FileName)"
    }
}

# Function: Analyze MSI File
function Analyze-MSI {
    if ([string]::IsNullOrEmpty($txtMSIPath.Text)) {
        Update-Output "Please select an MSI file first."
        return
    }

    $msiPath = $txtMSIPath.Text
    Update-Output "Analyzing MSI properties..."

    $installer = New-Object -ComObject WindowsInstaller.Installer
    $database = $installer.OpenDatabase($msiPath, 0)

    # Query for ProductCode, REBOOT, and ALLUSERS properties
    $properties = @("ProductCode", "REBOOT", "ALLUSERS")
    $results = @()

    foreach ($property in $properties) {
        $view = $database.OpenView("SELECT Value FROM Property WHERE Property = '$property'")
        $view.Execute()
        $record = $view.Fetch()
        if ($record) {
            $value = $record.StringData(1)
            $results += "${property}: $value"
        } else {
            $results += "${property}: Not Found"
        }
        $view.Close()
    }

    # Display results
    Update-Output "MSI Analysis Results:`n$($results -join "`n")"
}

# Button Click Events
$btnSelectMSI.Add_Click({ Select-MSI })
$btnAnalyzeMSI.Add_Click({ Analyze-MSI })

# Show the GUI
$window.ShowDialog()
