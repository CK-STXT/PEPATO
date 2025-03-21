# Ensure PowerShell is running in STA mode
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne "STA") {
    Start-Process PowerShell -ArgumentList "-sta -File `"$PSCommandPath`"" -NoNewWindow -Wait
    exit
}

# Load WPF
Add-Type -AssemblyName PresentationFramework

# Admin check
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    [System.Windows.MessageBox]::Show("This script must be run as administrator!", "Error", "OK", "Error")
    exit
}

# Output folder & timestamp
$baseFolder = "C:\temp\PrePackTool"
if (!(Test-Path $baseFolder)) { New-Item -ItemType Directory -Path $baseFolder -Force }
$timestamp = Get-Date -Format "dd_MM_yy"

# Snapshot paths
$snapshotBeforeFirewall = "$baseFolder\FirewallSnapshot_Before_$timestamp.txt"
$snapshotAfterFirewall  = "$baseFolder\FirewallSnapshot_After_$timestamp.txt"
$snapshotBeforeRegistry = "$baseFolder\RegistrySnapshot_Before_$timestamp.reg"
$snapshotAfterRegistry  = "$baseFolder\RegistrySnapshot_After_$timestamp.reg"

# XAML GUI layout
[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="🛠 PerPac Tool" Height="850" Width="650" ResizeMode="CanResizeWithGrip" Background="#1e1e1e" Foreground="White">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#333"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Padding" Value="5,2"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#2d2d30"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderBrush" Value="#555"/>
            <Setter Property="Margin" Value="5"/>
        </Style>
        <Style TargetType="GroupBox">
            <Setter Property="Margin" Value="10"/>
            <Setter Property="Foreground" Value="White"/>
        </Style>
    </Window.Resources>
    <ScrollViewer VerticalScrollBarVisibility="Auto">
        <StackPanel Margin="10">
            <GroupBox Header="🔥 Firewall Rules Snapshot">
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                    <Button Name="FirewallSnapshotBefore" Content="Snapshot Before" Width="150"/>
                    <Button Name="FirewallSnapshotAfter" Content="Snapshot After" Width="150"/>
                    <Button Name="FirewallCompare" Content="Compare" Width="120"/>
                </StackPanel>
            </GroupBox>
            <GroupBox Header="🧠 Registry Snapshot">
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                    <Button Name="RegistrySnapshotBefore" Content="Snapshot Before" Width="150"/>
                    <Button Name="RegistrySnapshotAfter" Content="Snapshot After" Width="150"/>
                    <Button Name="RegistryCompare" Content="Compare" Width="120"/>
                </StackPanel>
            </GroupBox>
            <GroupBox Header="📦 EXE Installer Handling">
                <StackPanel Orientation="Vertical">
                    <Button Name="SelectExe" Content="Select Installer (EXE)" Width="200" HorizontalAlignment="Center"/>
                    <TextBox Name="ExePath" Height="25" IsReadOnly="True"/>
                    <Button Name="RunInstaller" Content="Run Installer (As SYSTEM)" Width="200" HorizontalAlignment="Center"/>
                    <Button Name="FindMSI" Content="Find Extracted MSIs" Width="200" HorizontalAlignment="Center"/>
                </StackPanel>
            </GroupBox>
            <GroupBox Header="🧪 MSI Analysis">
                <StackPanel Orientation="Vertical">
                    <Button Name="SelectMSI" Content="Select MSI" Width="200" HorizontalAlignment="Center"/>
                    <TextBox Name="MSIPath" Height="25" IsReadOnly="True"/>
                    <Button Name="AnalyzeMSI" Content="Analyze MSI" Width="200" HorizontalAlignment="Center"/>
                </StackPanel>
            </GroupBox>
            <GroupBox Header="📋 Output Log">
                <TextBox Name="Output" Height="150" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto"/>
            </GroupBox>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                <Button Name="OpenFolder" Content="Open Folder" Width="120"/>
                <Button Name="Close" Content="Close" Width="120"/>
            </StackPanel>
        </StackPanel>
    </ScrollViewer>
</Window>
"@

# Load GUI
$reader = (New-Object System.Xml.XmlNodeReader $XAML)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Map controls
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

function Parse-NetshRuleToAddCommand($ruleOutput) {
    $rule = @{}
    foreach ($line in $ruleOutput) {
        if ($line -match "^\s*(.+?):\s*(.+)$") {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            $rule[$key] = $value
        }
    }

    if ($rule["Rule Name"]) {
        $cmd = "netsh advfirewall firewall add rule"
        $cmd += " name=`"$($rule["Rule Name"])`""
        if ($rule["Direction"])     { $cmd += " dir=$($rule["Direction"].ToLower())" }
        if ($rule["Action"])        { $cmd += " action=$($rule["Action"].ToLower())" }
        if ($rule["Enabled"])       { $cmd += " enable=$($rule["Enabled"].ToLower())" }
        if ($rule["Protocol"])      { $cmd += " protocol=$($rule["Protocol"])" }
        if ($rule["LocalPort"])     { $cmd += " localport=$($rule["LocalPort"])" }
        if ($rule["RemotePort"])    { $cmd += " remoteport=$($rule["RemotePort"])" }
        if ($rule["LocalIP"])       { $cmd += " localip=$($rule["LocalIP"])" }
        if ($rule["RemoteIP"])      { $cmd += " remoteip=$($rule["RemoteIP"])" }
        if ($rule["Profile"])       { $cmd += " profile=$($rule["Profile"].ToLower())" }
        if ($rule["InterfaceType"]) { $cmd += " interfacetype=$($rule["InterfaceType"].ToLower())" }
        return $cmd
    }
    return $null
}

function Compare-Snapshots($beforePath, $afterPath, $label) {
    if (!(Test-Path $beforePath) -or !(Test-Path $afterPath)) {
        Update-Output "$label snapshots not found!"
        return
    }

    $before = Get-Content $beforePath
    $after = Get-Content $afterPath
    $diff = Compare-Object $before $after -PassThru -IncludeEqual:$false

    if ($label -eq "Firewall" -and $diff) {
        $addedLines = $diff | Where-Object { $_.SideIndicator -eq "=>" }
        $ruleNames = $addedLines | Where-Object { $_ -match "^Rule Name:\s*(.+)" } | ForEach-Object {
            $matches[1].Trim()
        } | Sort-Object -Unique

        $commandLines = @()
        foreach ($ruleName in $ruleNames) {
            $ruleOutput = netsh advfirewall firewall show rule name="$ruleName"
            $cmd = Parse-NetshRuleToAddCommand $ruleOutput
            if ($cmd) { $commandLines += $cmd }
        }

        $diffPath = "$baseFolder\${label}_Differences_$timestamp.txt"
        if ($commandLines.Count -gt 0) {
            $commandLines | Out-File -FilePath $diffPath -Encoding UTF8
            Update-Output "Firewall 'add rule' commands saved to: $diffPath"
        } else {
            Update-Output "No new firewall rules found."
        }
    }
    elseif ($diff) {
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
    try {
        $msiPath = $elements.MSIPath.Text
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $db = $installer.OpenDatabase($msiPath, 0)

        $props = @("ProductCode", "REBOOT", "ALLUSERS")
        $results = foreach ($p in $props) {
            $view = $db.OpenView("SELECT Value FROM Property WHERE Property='$p'")
            $view.Execute()
            $record = $view.Fetch()
            $view.Close()
            if ($record) { "${p}: $($record.StringData(1))" } else { "${p}: Not Found" }
        }

        $shortcuts = @()
        $view = $db.OpenView("SELECT Shortcut, Target, Directory_ FROM Shortcut")
        $view.Execute()
        while ($rec = $view.Fetch()) {
            $shortcuts += "Shortcut: $($rec.StringData(1)) → $($rec.StringData(2)) (Dir: $($rec.StringData(3)))"
        }
        $view.Close()

        $output = ($results -join "`n")
        if ($shortcuts.Count -gt 0) {
            $output += "`n`nShortcuts:`n" + ($shortcuts -join "`n")
        } else {
            $output += "`n`nNo shortcuts found."
        }

        Update-Output $output
    } catch {
        Update-Output "Error analyzing MSI: $_"
    }
}

# Wire buttons
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
