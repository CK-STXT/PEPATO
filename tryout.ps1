# Ensure PowerShell is running in STA mode
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne "STA") {
    Start-Process PowerShell -ArgumentList "-sta -File `"$PSCommandPath`"" -NoNewWindow -Wait
    exit
}

# Load WPF assembly
Add-Type -AssemblyName PresentationFramework

# Check for Administrator rights
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    [System.Windows.MessageBox]::Show("This script must be run as administrator!", "Error", "OK", "Error")
    exit
}

# Setup base folder and timestamp
$baseFolder = "C:\\temp\\PrePackTool"
if (!(Test-Path $baseFolder)) { New-Item -ItemType Directory -Path $baseFolder -Force }
$timestamp = Get-Date -Format "dd_MM_yy"

# Define snapshot paths
$snapshotBeforeFirewall = "$baseFolder\FirewallSnapshot_Before_$timestamp.txt"
$snapshotAfterFirewall = "$baseFolder\FirewallSnapshot_After_$timestamp.txt"
$snapshotBeforeRegistry = "$baseFolder\RegistrySnapshot_Before_$timestamp.reg"
$snapshotAfterRegistry = "$baseFolder\RegistrySnapshot_After_$timestamp.reg"

# GUI layout (XAML)
[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="PerPac Tool (Running as SYSTEM)" Height="850" Width="600" ResizeMode="NoResize">
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

        <GroupBox Header="Firewall" Grid.Row="0" Margin="10">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                <Button Name="FirewallSnapshotBefore" Content="Snapshot Before" Width="120" Margin="5"/>
                <Button Name="FirewallSnapshotAfter" Content="Snapshot After" Width="120" Margin="5"/>
                <Button Name="FirewallCompare" Content="Compare" Width="100" Margin="5"/>
            </StackPanel>
        </GroupBox>

        <GroupBox Header="Registry" Grid.Row="1" Margin="10">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                <Button Name="RegistrySnapshotBefore" Content="Snapshot Before" Width="120" Margin="5"/>
                <Button Name="RegistrySnapshotAfter" Content="Snapshot After" Width="120" Margin="5"/>
                <Button Name="RegistryCompare" Content="Compare" Width="100" Margin="5"/>
            </StackPanel>
        </GroupBox>

        <GroupBox Header="Installer Handling" Grid.Row="2" Margin="10">
            <StackPanel Orientation="Vertical" HorizontalAlignment="Center">
                <Button Name="SelectExe" Content="Select Installer (EXE)" Width="200" Margin="5"/>
                <TextBox Name="ExePath" Width="450" Height="25" Margin="5" IsReadOnly="True"/>
                <Button Name="RunInstaller" Content="Run Installer (As SYSTEM)" Width="200" Margin="5"/>
                <Button Name="FindMSI" Content="Find Extracted MSIs" Width="200" Margin="5"/>
            </StackPanel>
        </GroupBox>

        <GroupBox Header="MSI Analysis" Grid.Row="3" Margin="10">
            <StackPanel Orientation="Vertical" HorizontalAlignment="Center">
                <Button Name="SelectMSI" Content="Select MSI" Width="200" Margin="5"/>
                <TextBox Name="MSIPath" Width="450" Height="25" Margin="5" IsReadOnly="True"/>
                <Button Name="AnalyzeMSI" Content="Analyze MSI" Width="200" Margin="5"/>
            </StackPanel>
        </GroupBox>

        <TextBox Name="Output" Width="550" Height="150" Margin="10" Grid.Row="4" IsReadOnly="True" VerticalAlignment="Top" HorizontalAlignment="Center" TextWrapping="Wrap"/>

        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Center">
            <Button Name="OpenFolder" Content="Open Folder" Width="120" Margin="10"/>
            <Button Name="Close" Content="Close" Width="120" Margin="10"/>
        </StackPanel>
    </Grid>
</Window>
"@

# Load GUI
$reader = (New-Object System.Xml.XmlNodeReader $XAML)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get UI elements
$elements = @{
    FirewallSnapshotBefore = $window.FindName("FirewallSnapshotBefore")
    FirewallSnapshotAfter  = $window.FindName("FirewallSnapshotAfter")
    FirewallCompare        = $window.FindName("FirewallCompare")
    RegistrySnapshotBefore = $window.FindName("RegistrySnapshotBefore")
    RegistrySnapshotAfter  = $window.FindName("RegistrySnapshotAfter")
    RegistryCompare        = $window.FindName("RegistryCompare")
    SelectExe              = $window.FindName("SelectExe")
    ExePath                = $window.FindName("ExePath")
    RunInstaller           = $window.FindName("RunInstaller")
    FindMSI                = $window.FindName("FindMSI")
    SelectMSI              = $window.FindName("SelectMSI")
    MSIPath                = $window.FindName("MSIPath")
    AnalyzeMSI             = $window.FindName("AnalyzeMSI")
    Output                 = $window.FindName("Output")
    OpenFolder             = $window.FindName("OpenFolder")
    Close                  = $window.FindName("Close")
}

function Update-Output($text) {
    $window.Dispatcher.Invoke([action] { $elements.Output.Text = $text })
}

function Take-FirewallSnapshot($path, $type) {
    Update-Output "Firewall snapshot ($type) started..."
    netsh advfirewall firewall show rule name=all | Out-File -FilePath $path -Encoding UTF8
    Update-Output "Firewall snapshot ($type) saved to: $path"
}

function Take-RegistrySnapshot($path, $type) {
    Update-Output "Registry snapshot ($type) started..."
    reg export "HKLM\Software" $path /y
    Update-Output "Registry snapshot ($type) saved to: $path"
}

function Compare-Snapshots($beforePath, $afterPath, $label) {
    if (!(Test-Path $beforePath) -or !(Test-Path $afterPath)) {
        Update-Output "$label snapshots not found!"
        return
    }
    $before = Get-Content $beforePath
    $after = Get-Content $afterPath
    $diff = Compare-Object $before $after
    if ($diff) {
        $diffPath = "$baseFolder\${label}_Differences_$timestamp.txt"
        $diff | Out-File -FilePath $diffPath -Encoding UTF8
        Update-Output "Differences saved to: $diffPath"
    } else {
        Update-Output "No differences found in $label."
    }
}

function Select-File($filter) {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = $filter
    if ($dialog.ShowDialog() -eq 'OK') { return $dialog.FileName }
    return $null
}

function Analyze-MSI {
    if ([string]::IsNullOrEmpty($elements.MSIPath.Text)) {
        Update-Output "Please select an MSI file first."
        return
    }
    $msiPath = $elements.MSIPath.Text
    Update-Output "Analyzing MSI: $msiPath"
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $db = $installer.OpenDatabase($msiPath, 0)
        $props = @("ProductCode", "REBOOT", "ALLUSERS")
        $results = foreach ($p in $props) {
            $v = $db.OpenView("SELECT Value FROM Property WHERE Property='$p'")
            $v.Execute()
            $r = $v.Fetch()
            $v.Close()
            if ($r) { "${p}: $($r.StringData(1))" } else { "${p}: Not Found" }
        }
        $shortcuts = @()
        $view = $db.OpenView("SELECT Shortcut, Target, Directory_ FROM Shortcut")
        $view.Execute()
        while ($rec = $view.Fetch()) {
            $shortcuts += "Shortcut: $($rec.StringData(1)) → Target: $($rec.StringData(2)) (Dir: $($rec.StringData(3)))"
        }
        $view.Close()
        $out = $results -join "`n"
        if ($shortcuts.Count -gt 0) {
            $out += "`n`nShortcuts:`n" + ($shortcuts -join "`n")
        } else {
            $out += "`n`nNo shortcuts found."
        }
        Update-Output $out
    } catch {
        Update-Output "Error analyzing MSI: $_"
    }
}

# Wire up buttons
$elements.FirewallSnapshotBefore.Add_Click({ Take-FirewallSnapshot $snapshotBeforeFirewall "Before" })
$elements.FirewallSnapshotAfter.Add_Click({ Take-FirewallSnapshot $snapshotAfterFirewall "After" })
$elements.FirewallCompare.Add_Click({ Compare-Snapshots $snapshotBeforeFirewall $snapshotAfterFirewall "Firewall" })
$elements.RegistrySnapshotBefore.Add_Click({ Take-RegistrySnapshot $snapshotBeforeRegistry "Before" })
$elements.RegistrySnapshotAfter.Add_Click({ Take-RegistrySnapshot $snapshotAfterRegistry "After" })
$elements.RegistryCompare.Add_Click({ Compare-Snapshots $snapshotBeforeRegistry $snapshotAfterRegistry "Registry" })
$elements.SelectExe.Add_Click({ $path = Select-File "Executable Files (*.exe)|*.exe"; if ($path) { $elements.ExePath.Text = $path; Update-Output "Selected EXE: $path" } })
$elements.SelectMSI.Add_Click({ $path = Select-File "MSI Files (*.msi)|*.msi"; if ($path) { $elements.MSIPath.Text = $path; Update-Output "Selected MSI: $path" } })
$elements.AnalyzeMSI.Add_Click({ Analyze-MSI })
$elements.OpenFolder.Add_Click({ Start-Process explorer.exe -ArgumentList $baseFolder })
$elements.Close.Add_Click({ $window.Close() })

# Show GUI
$window.ShowDialog()