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
$Processes = Get-Process | Sort PagedMemorySize -Descending | Select -First 10

echo $Processes

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
$p = 1
foreach ($cpu in $InstalledProcessors) {
    $Report += @"
        <p class="indent">CPU #$($p): <span>$($cpu.Name)</span></p>
        <p class="indent">CPU #$($p) Speed: <span>$($cpu.MaxClockSpeed)</span></p>
        <p class="indent">CPU #$($p) Cores: <span>$($cpu.NumberOfCores)</span></p>
        <p class="indent">CPU #$($p) Threads: <span>$($cpu.NumberOfLogicalProcessors)</span></p>
"@
    $p++
}

# Continue Entering other relevant information we have already collected 
$Report += @"
        <p>Installed Memory: <span>$($ComputerInfo.CsPhysicallyInstalledMemory / 1024 / 1024) GB</span></p>
        <p>Installed Hard Drives: <span>$($HardDrives.Length)</span></p>
"@

# Loop through our Disk Array and Print out Relevant Information About Them
$c = 1
foreach ($drive in $HardDrives) {
    $Report += @"
        <p class="indent">Disk #$($c) Drive Letter: <span>$($drive.DeviceID)</span></p>
        <p class="indent">Disk #$($c) Drive Space: <span>$([math]::Round(($drive.Size / 1024 / 1024 / 1024), 2)) GB</span></p>
        <p class="indent">Disk #$($c) Free Space: <span>$([math]::Round(($drive.FreeSpace / 1024 / 1024 / 1024), 2)) GB</span></p>
        <p class="indent">Disk #$($c) File System: <span>$($drive.FileSystem)</span></p>
"@
    $c++
}

# Loop through our installed Video Cards and add their information to the report now
$Report += @"
        <p>Video Cards Installed: <span>$($VideoCards.Length)</span></p>
"@

$v = 1
foreach ($card in $VideoCards) {
    $Report += @"
        <p class="indent">Video Card #$($v): <span>$($card.Name)</span></p>
        <p class="indent">Video Card #$($v) Memory: <span>$([math]::Round($card.AdapterRam / 1024 / 1024 / 1024)) GB</span></p>
"@
    $v++
}

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




$Report = ConvertTo-Html -Title 'Windows System Report' -Body $Header,$Report,$Footer -CssUri .\styles.css


$Report | Out-File -FilePath $ReportName
