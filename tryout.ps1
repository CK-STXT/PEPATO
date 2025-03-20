# Ensure PowerShell is running in STA mode
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne "STA") {
    Start-Process PowerShell -ArgumentList "-sta -File `"$PSCommandPath`"" -NoNewWindow -Wait
    exit
}

# Load WPF Assembly
Add-Type -AssemblyName PresentationFramework

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

# Validate Window Loading
if ($null -eq $window) {
    Write-Host "Failed to load the XAML UI. Check your XAML syntax."
    exit
}

# Find Controls
$btnFirewallSnapshotBefore = $window.FindName("btnFirewallSnapshotBefore")
$btnFirewallSnapshotAfter = $window.FindName("btnFirewallSnapshotAfter")
$btnCompareFirewall = $window.FindName("btnCompareFirewall")
$txtOutput = $window.FindName("txtOutput")

# Validate UI Element Existence
if ($null -eq $btnFirewallSnapshotBefore -or $null -eq $btnFirewallSnapshotAfter -or $null -eq $btnCompareFirewall -or $null -eq $txtOutput) {
    Write-Host "One or more UI elements were not found. Check the XAML control names."
    exit
}

# Define File Paths
$snapshotBeforePath = "$env:TEMP\FirewallSnapshot_Before.txt"
$snapshotAfterPath = "$env:TEMP\FirewallSnapshot_After.txt"
$diffPath = "$env:TEMP\Firewall_Differences.txt"

# Function: Take Firewall Snapshot (Before)
function Take-FirewallSnapshotBefore {
    param([System.Windows.Controls.TextBox]$OutputBox)
    netsh advfirewall firewall show rule name=all | Out-File -FilePath $snapshotBeforePath -Encoding UTF8
    $OutputBox.Text = "Firewall snapshot (before) saved to: $snapshotBeforePath"
}

# Function: Take Firewall Snapshot (After)
function Take-FirewallSnapshotAfter {
    param([System.Windows.Controls.TextBox]$OutputBox)
    netsh advfirewall firewall show rule name=all | Out-File -FilePath $snapshotAfterPath -Encoding UTF8
    $OutputBox.Text = "Firewall snapshot (after) saved to: $snapshotAfterPath"
}

# Function: Compare Snapshots
function Compare-FirewallSnapshots {
    param([System.Windows.Controls.TextBox]$OutputBox)
    if (!(Test-Path $snapshotBeforePath) -or !(Test-Path $snapshotAfterPath)) {
        $OutputBox.Text = "Snapshots not found! Take both snapshots first."
        return
    }
    $before = Get-Content $snapshotBeforePath
    $after = Get-Content $snapshotAfterPath
    $diff = Compare-Object -ReferenceObject $before -DifferenceObject $after
    if ($diff) {
        $diff | Out-File -FilePath $diffPath -Encoding UTF8
        $OutputBox.Text = "Differences saved to: $diffPath"
    } else {
        $OutputBox.Text = "No differences found."
    }
}

# Button Events
$btnFirewallSnapshotBefore.Add_Click({
    Take-FirewallSnapshotBefore -OutputBox $txtOutput
})
$btnFirewallSnapshotAfter.Add_Click({
    Take-FirewallSnapshotAfter -OutputBox $txtOutput
})
$btnCompareFirewall.Add_Click({
    Compare-FirewallSnapshots -OutputBox $txtOutput
})

# Show Window
$window.ShowDialog()
