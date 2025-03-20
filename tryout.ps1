[System.Reflection.Assembly]::LoadWithPartialName('presentationframework')

# XAML for WPF GUI
[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
        Title="PerPac Tool" Height="250" Width="400">
    <Grid>
        <Button Name="btnFirewallSnapshot" Content="Take Firewall Snapshot" Width="200" Height="30" Margin="100,20,0,0" HorizontalAlignment="Left"/>
        <Button Name="btnRegistrySnapshot" Content="Take Registry Snapshot" Width="200" Height="30" Margin="100,70,0,0" HorizontalAlignment="Left"/>
        <TextBox Name="txtOutput" Width="360" Height="70" Margin="10,120,10,10" IsReadOnly="True"/>
    </Grid>
</Window>
"@

# Load WPF
$reader = (New-Object System.Xml.XmlNodeReader $XAML)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Define button actions
$btnFirewallSnapshot = $window.FindName("btnFirewallSnapshot")
$btnRegistrySnapshot = $window.FindName("btnRegistrySnapshot")
$txtOutput = $window.FindName("txtOutput")

# Firewall Snapshot Function
function Take-FirewallSnapshot {
    $snapshot = Get-NetFirewallRule | Select-Object DisplayName, Direction, Action, Enabled
    $snapshot | Out-File -FilePath "FirewallSnapshot.txt"
    $txtOutput.Text = "Firewall snapshot saved!"
}

# Registry Snapshot Function
function Take-RegistrySnapshot {
    $regSnapshot = Get-ChildItem -Recurse HKLM:\Software, HKCU:\Software | Select-Object Name
    $regSnapshot | Out-File -FilePath "RegistrySnapshot.txt"
    $txtOutput.Text = "Registry snapshot saved!"
}

# Button Events
$btnFirewallSnapshot.Add_Click({ Take-FirewallSnapshot })
$btnRegistrySnapshot.Add_Click({ Take-RegistrySnapshot })

# Show Window
$window.ShowDialog()
