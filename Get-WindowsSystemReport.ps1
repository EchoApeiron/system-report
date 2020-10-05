### Variables for Reporting File ### 
$GeneratedDate = Get-Date -Format "MM-dd-yyyy HH:mm"
$ReportName = 'Windows_System_Report.html'
$ReportStyles = Get-Content .\styles.css

#### System Information Collection ####
$ComputerInfo = Get-ComputerInfo
### CPU Statistics ###
$InstalledProcessors = $ComputerInfo.CsProcessors # This is an array of objects and needs to be iterated through
### Hard Drive Statistics 
$HardDrives = Get-CimInstance win32_logicaldisk # This is an array of objects and needs to be iterated through
### Video Card Statistics 
$VideoCards = Get-CimInstance Win32_VideoController # This is an array of objects and needs to be iterated 
### Process Information 
$Processes = Get-Process | Sort-Object PagedMemorySize -Descending | Select-Object -First 10

# Check if report was ran previously, if so remove old report so we can recreate the file
if ([System.IO.File]::Exists($ReportName)) {
    Remove-Item -Path $ReportName
}

# Start building our HTML Report with our Header Template
$Header = @"
<header class="masthead">
    <h1>System Report</h1>
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
        <p>Motherboard Model: <span>$($ComputerInfo.CsModel)</span></p>
        <p>Installed Processors: <span>$($InstalledProcessors.Length)</span></p>
"@

# Loop through our CPU Array and Print out Relevant Information About Them 
for($p=0; $p -le $InstalledProcessors.Length-1; $p++) {
    $Report += @"
        <p class="indent">CPU #$($p+1): <span>$($InstalledProcessors[$p].Name)</span></p>
        <p class="indent">CPU #$($p+1) Speed: <span>$($InstalledProcessors[$p].MaxClockSpeed)</span></p>
        <p class="indent">CPU #$($p+1) Cores: <span>$($InstalledProcessors[$p].NumberOfCores)</span></p>
        <p class="indent">CPU #$($p+1) Threads: <span>$($InstalledProcessors[$p].NumberOfLogicalProcessors)</span></p>
"@
}



# Continue Entering other relevant information we have already collected 
$Report += @"
        <p>Installed Memory: <span>$($ComputerInfo.CsPhysicallyInstalledMemory / 1024 / 1024) GB</span></p>
        <p>Installed Hard Drives: <span>$($HardDrives.Length)</span></p>
"@

# Loop through our Disk Array and Print out Relevant Information About Them
for($c=0; $c -le $HardDrives.Length-1; $c++) {
    $Report += @"
        <p class="indent">Disk #$($c+1) Drive Letter: <span>$($HardDrives[$c].DeviceID)</span></p>
        <p class="indent">Disk #$($c+1) Drive Space: <span>$([math]::Round(($HardDrives[$c].Size / 1024 / 1024 / 1024), 2)) GB</span></p>
        <p class="indent">Disk #$($c+1) Free Space: <span>$([math]::Round(($HardDrives[$c].FreeSpace / 1024 / 1024 / 1024), 2)) GB</span></p>
        <p class="indent">Disk #$($c+1) File System: <span>$($HardDrives[$c].FileSystem)</span></p>
"@
}

# Loop through our installed Video Cards and add their information to the report now
$Report += @"
        <p>Video Cards Installed: <span>$($VideoCards.Length)</span></p>
"@

for($v=0; $v -le $VideoCards.Length-1; $v++) {
    $Report += @"
        <p class="indent">Video Card #$($v+1): <span>$($VideoCards[$v].Name)</span></p>
        <p class="indent">Video Card #$($v+1) Memory: <span>$([math]::Round($VideoCards[$v].AdapterRam / 1024 / 1024 / 1024)) GB</span></p>
"@
}

# Process information about the current system 
$Report += @"
    <h3>Top Ten System Processes</h3>
    <table class="indent-table">
        <tr>
            <th>Process Name</th><th>Process ID</th><th>Memory Used</th>
        </tr>
"@

foreach($Process in $Processes) {
    $Report += @"
        <tr>
            <td><span>$($Process.Name)</span></td><td><span>$($Process.Id)</span></td><td><span>$([math]::Round(($Process.PagedMemorySize / 1024 / 1024), 2)) MB</span></td>
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
$Report = ConvertTo-Html -Head "<style>$($ReportStyles)</style>" -Title 'Windows System Report' -Body $Header,$Report,$Footer 
$Report | Out-File -FilePath $ReportName