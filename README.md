# system-report

This is a Powershell script that uses an external CSS file to gather System Information about a Windows computer and then exports the data to an HTML report for a user to later review. 

Currently this is a limited functioning script that just goes and processes onto a local computer. It will only iterate Windows System information for now, and uses a lot of .NET and CIM instances to gather hardware information at this time. 

The script will currently only report on the OS information and the following hardware statistics: 
 - Motherboard Information
 - All Installed CPUs
 - Memory 
 - All Intalled Hard Drives 
 - All Installed Video Cards

