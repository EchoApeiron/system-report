### Variables for Reporting File ### 
$GeneratedDate = Get-Date -Format "MM-dd-yyyy HH:mm"
$ReportName = 'Windows_System_Report.html'

#### System Information Collection ####
$ComputerInfo = Get-ComputerInfo
### CPU Statistics ###
$InstalledProcessors = $ComputerInfo.CsProcessors # This is an array of objects and needs to be iterated through
### Hard Drive Statistics 
$HardDrives = Get-CimInstance win32_logicaldisk # This is an array of objects and needs to be iterated through
### Video Card Statistics 
$VideoCards = Get-CimInstance Win32_VideoController # This is an array of objects and needs to be iterated 
### Process Information 
$Processes = Get-Process | Sort-Object PagedMemorySize -Descending | Select-Object -First 20

# Check if report was ran previously, if so remove old report so we can recreate the file
if ([System.IO.File]::Exists($ReportName)) {
    Remove-Item -Path $ReportName
}

# Start building our HTML Report with our Header Template
$Header = @"
<header class="masthead">
    <h1>System Statistics</h1>
</header>
<main class="report">
"@

# Construct the Body of the HTML Report and start adding already gathered system info
$Report += @"
        <h2>System Summary</h2>
        <p>Computer Name: <span>$($ComputerInfo.CsName)</span></p>
        <p>Operating System: <span>$($ComputerInfo.WindowsProductName)</span></p>
        <p>OS Version: <span>$($ComputerInfo.OsVersion)</span></p>
        <p>OS Build: <span>$($ComputerInfo.OsBuildNumber)</span></p>
        <h3>Hardware Profile</h3>
        <p>Motherboard Manufacturer: <span>$($ComputerInfo.CsManufacturer)</span></p>
        <p class="indent">Motherboard Model: <span>$($ComputerInfo.CsModel)</span></p>
        <p>Installed Processors: <span>$($InstalledProcessors.Length)</span></p>
        <table class="system-table">
            <tr>
                <th>CPU</th><th>Frequency</th><th>Cores</th><th>Threads</th>
            </tr>
"@

# Loop through our CPU Array and Print out Relevant Information About Them 
for ($p=0; $p -le $InstalledProcessors.Length-1; $p++) {
    if (!($p % 2)) {
        $Report += @"
            <tr class="alt-row">
"@
    }
    else {
        $Report += @"
            <tr>
"@
    }

    $Report += @"
                <td><span>$($InstalledProcessors[$p].Name)</span></td>
                <td><span>$($InstalledProcessors[$p].MaxClockSpeed / 1000) GHz</span></td>
                <td><span>$($InstalledProcessors[$p].NumberOfCores)</span></td>
                <td><span>$($InstalledProcessors[$p].NumberOfLogicalProcessors)</span></td>
            </tr>
"@
}

$Report += @"
        </table>
"@

# Continue Entering other relevant information we have already collected 
$Report += @"
        <p>Installed Memory: <span>$($ComputerInfo.CsPhysicallyInstalledMemory / 1024 / 1024) GB</span></p>
        <p>Installed Hard Drives: <span>$($HardDrives.Length)</span></p>
        <table class="system-table">
            <tr>
                <th>Drive Letter</th><th>Drive Space</th><th>Free Space</th><th>File System</th>
            </tr>
"@

# Loop through our Disk Array and Print out Relevant Information About Them
for ($c=0; $c -le $HardDrives.Length-1; $c++) {
    if (!($c % 2)) {
        $Report += @"
            <tr class="alt-row">
"@
    }
    else {
        $Report += @"
            <tr>
"@
    }

    $Report += @"
                <td><span>$($HardDrives[$c].DeviceID)</span></td>
                <td><span>$([math]::Round(($HardDrives[$c].Size / 1024 / 1024 / 1024), 2)) GB</span></td>
                <td><span>$([math]::Round(($HardDrives[$c].FreeSpace / 1024 / 1024 / 1024), 2)) GB</span></td>
                <td><span>$($HardDrives[$c].FileSystem)</span></td>
            </tr>
"@
}

$Report += @"
        </table>
"@

# Loop through our installed Video Cards and add their information to the report now
$Report += @"
        <p>Video Cards Installed: <span>$($VideoCards.Length)</span></p>
        <table class="system-table">
            <tr>
                <th>Video Card</th><th>Video Card Memory</th>
            </tr>
"@

for ($v=0; $v -le $VideoCards.Length-1; $v++) {
    if (!($v % 2)) {
        $Report += @"
            <tr class="alt-row">
"@
    }
    else {
        $Report += @"
            <tr>
"@
    }

    $Report += @"
                <td><span>$($VideoCards[$v].Name)</span></td>
                <td><span>$([math]::Round($VideoCards[$v].AdapterRam / 1024 / 1024 / 1024)) GB</span></td>
            </tr>
"@
}

$Report += @"
        </table>
"@

# Process information about the current system 
$Report += @"
    <h3>Heavy System Processes</h3>
    <table class="indent-table">
        <tr>
            <th>Process Name</th><th>Process ID</th><th>Memory Used</th>
        </tr>
"@

for ($p=0; $p -le $Processes.Length-1; $p++) {
    if (!($p % 2)) {
        $Report += @"
            <tr class="alt-row">
"@
    }
    else {
        $Report += @"
            <tr>
"@
    }

        $Report += @"
        <td><span>$($Processes[$p].Name)</span></td><td><span>$($Processes[$p].Id)</span></td><td><span>$([math]::Round(($Processes[$p].PagedMemorySize / 1024 / 1024), 2)) MB</span></td>
        </tr>
"@
}

$Report += @"
        </table>
"@

# Export a custom footer for our page 
$Footer += @"
</main>
<footer class="footer">
    <p>Report Ran On: $GeneratedDate</p>
</footer>
"@

# Finally get data converted to a valid HTML page 
$Report = ConvertTo-Html -Head "<style>$(Get-Content .\styles.css)</style>" -Title 'Windows System Report' -Body $Header,$Report,$Footer 
$Report | Out-File -FilePath $ReportName