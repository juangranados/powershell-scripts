# PrintNightmare - Install printer drivers on remote computers using printerExport file

[Right click here and select "Save link as" to download](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Bulk%20Change%20PrinterDriverAttributes%20for%20non%20Package-Aware%20printer%20drivers/ChangePrinterDriverAttributes.ps1)

On August 10, Microsoft posted a [blog post](https://msrc-blog.microsoft.com/2021/08/10/point-and-print-default-behavior-change/) about changes to the point and print.

After the August patches, standard users cant add any printers. This  means that you need to pre-install all drivers on your workstations or Remote Session Host servers.

In [KB5005652](https://support.microsoft.com/topic/873642bf-2634-49c5-a23b-6d8e9a302872) documentation, Microsoft recommends this four solutions to install print drivers:

This script allows remote printer drivers installation on a group of computers using printerExport file.

To generate a printer export file with all printer drivers, run `PrintbrmUI.exe` from the computer or server you want to export them.

## Parameters

**printerExportFile**: Path to the printerExport file.

*Example: \\\SRVFS01\Drivers\drivers.printerExport*

**ComputerList**: List of computers in install printer drivers. You can only use one source of target computers: ComputerList, OU or CSV.

*Example: Computer001,Computer002,Computer003 (Without quotation marks)*

**OU**: OU containing computers in which install printer drivers. You can only use one source of target computers: ComputerList, OU or CSV.
SAT for AD module for PowerShell must be installed in order to query AD.
If you run script from a Domain Controller, AD module for PowerShell is already enabled.

*Example: : 'OU=RDSH,OU=Servers,DC=CONTOSO,DC=COM'*

**CSV**: CSV file containing computers in which install printer drivers. You can only use one source of target computers: ComputerList, OU or CSV.

*Example: 'C:\Scripts\Computers.csv'*
CSV Format:

```CSV
Name
RDSH01
RDSH03
RDSH02
```

**LogPath:** Path where save log file.
Default: My Documents

*Example: C:\Logs*

**Credential**: Script will ask for an account to perform remote installation.
