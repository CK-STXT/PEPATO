[System.Reflection.Assembly]::LoadWithPartialName('presentationframework')

# XAML for WPF GUI
[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
        Title="PerPac Tool" Height="300" Width="400">
    <Grid>
        <Button Name="btnFirewallSnapshotBefore" Content="Firewall Snapshot (Before)" Width="200" Height="30" Margin="100,20,0,0" HorizontalAlignment="Left"/>
        <Button Name="btnFirewallSnapshotAfter" Content="Firewall Snapshot (After)" Width="200" Height="30" Margin="100,60,0,0" HorizontalAlignment="Left"/>
        <Button Name="btnCompareFirewall" Content="Compare Firewall Snapshots" Width="200" Height="30" Margin="100,100,0,0" HorizontalAlignment="Left"/>
        <TextBox Name="txtOutput" Width="360" Height="70" Margin="10,150,10,10" IsReadOnly="True"/>
    </Grid>
</Window>
"@

# Load WPF
$reader = (New-Object System.Xml.XmlNodeReader $XAML)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Define button actions
$btnFirewallSnapshotBefore = $window.FindName("btnFirewallSnapshotBefore")
$btnFirewallSnapshotAfter = $window.FindName("btnFirewallSnapshotAfter")
$btnCompareFirewall = $window.FindName("btnCompareFirewall")
$txtOutput = $window.FindName("txtOutput")

# Firewall Snapshot Function (Before Installation)
function Take-FirewallSnapshotBefore {
    netsh advfirewall firewall show rule name=all | Out-File -FilePath "$(pwd)\FirewallSnapshot_Before.txt" -Encoding UTF8
    $txtOutput.Text = "Firewall snapshot (before) saved!"
}

# Firewall Snapshot Function (After Installation)
function Take-FirewallSnapshotAfter {
    netsh advfirewall firewall show rule name=all | Out-File -FilePath "$(pwd)\FirewallSnapshot_After.txt" -Encoding UTF8
    $txtOutput.Text = "Firewall snapshot (after) saved!"
}

# Compare Firewall Snapshots
function Compare-FirewallSnapshots {
    if (!(Test-Path "$(pwd)\FirewallSnapshot_Before.txt") -or !(Test-Path "$(pwd)\FirewallSnapshot_After.txt")) {
        $txtOutput.Text = "Snapshots not found! Take both snapshots first."
        return
    }
    $before = Get-Content "$(pwd)\FirewallSnapshot_Before.txt"
    $after = Get-Content "$(pwd)\FirewallSnapshot_After.txt"
    $diff = Compare-Object -ReferenceObject $before -DifferenceObject $after
    if ($diff) {
        $diff | Out-File -FilePath "$(pwd)\Firewall_Differences.txt" -Encoding UTF8
        $txtOutput.Text = "Differences saved to Firewall_Differences.txt"
    } else {
        $txtOutput.Text = "No differences found."
    }
}

# Button Events
$btnFirewallSnapshotBefore.Add_Click({ Take-FirewallSnapshotBefore })
$btnFirewallSnapshotAfter.Add_Click({ Take-FirewallSnapshotAfter })
$btnCompareFirewall.Add_Click({ Compare-FirewallSnapshots })

# Show Window
$window.ShowDialog()
