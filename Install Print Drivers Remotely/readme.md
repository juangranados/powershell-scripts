# PrintNightmare - Install printer drivers on remote computers using printerExport file

[Right click here and select "Save link as" to download](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Install%20Print%20Drivers%20Remotely/Install-PrinterDriversRemotely.ps1)

![Screenshot](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Install%20Print%20Drivers%20Remotely/3.png)

On August 10, Microsoft posted a [blog post](https://msrc-blog.microsoft.com/2021/08/10/point-and-print-default-behavior-change/) about changes to the point and print.

After the August patches, standard users cant add any printers. This  means that you need to pre-install all drivers on your workstations or Remote Session Host servers.

In [KB5005652](https://support.microsoft.com/topic/873642bf-2634-49c5-a23b-6d8e9a302872) documentation, Microsoft recommends this four possible solutions after install KB5005652 patch:

![Solutions to KB5005652](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Install%20Print%20Drivers%20Remotely/1.PNG)

This script allows remote printer drivers installation on a group of computers using printerExport file and [PSExec](https://docs.microsoft.com/en-us/sysinternals/downloads/pstools). If PSExec is not found on computer, script asks to the user for download it and extract in system folder. I recommend use PSExec latest version because v2.2 does not launch printbrm.exe properly.

To generate a printer export file with all printer drivers, run `PrintbrmUI.exe` from the computer or server you want to export them. 

Be carefully because this tool exports all printers and ports too. I recommend install all drivers in a test computer without printers and export them using `PrintbrmUI.exe`. Alternatively, you can export all from print server, import in a test computer, delete printers and ports and export again to obtain a printerExport file with only drivers.

![PrintbrmUI screenshot](https://github.com/juangranados/powershell-scripts/blob/main/Install%20Print%20Drivers%20Remotely/2.PNG?raw=true)

## Parameters

**printerExportFile**: Path to the printerExport file.

*Example: \\\SRVFS01\Drivers\drivers.printerExport*

**ComputerList**: List of computers in install printer drivers. You can only use one source of target computers: ComputerList, OU or CSV.

*Example: Computer001,Computer002,Computer003 (Without quotation marks)*

**OU**: OU containing computers in which install printer drivers. You can only use one source of target computers: ComputerList, OU or CSV.
RSAT for AD module for PowerShell must be installed in order to query AD.
If you run script from a Domain Controller, AD module for PowerShell is already enabled.

To install it from Windows 10 computer

```powershell
Get-WindowsCapability -Online |? {$_.Name -like "*RSAT.ActiveDirectory*" -and $_.State -eq "NotPresent"} | Add-WindowsCapability -Online
```

To install it from server

```powershell
Install-WindowsFeature RSAT-AD-PowerShell
```

*Example: : "OU=RDSH,OU=Servers,DC=CONTOSO,DC=COM"*

**CSV**: CSV file containing computers in which install printer drivers. You can only use one source of target computers: ComputerList, OU or CSV.

*Example: "C:\Scripts\Computers.csv"*
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

## Remote installation examples
```powershell
Install-PrinterExportRemoteComputers.ps1 -printerExportFile "\\MV-SRV-PR01\Drivers\print_drivers.printerExport" -OU "OU=RDS,OU=Datacenter,DC=CONTOSO,DC=COM"  
Install-PrinterExportRemoteComputers.ps1 -printerExportFile "\\MV-SRV-PR01\Drivers\print_drivers.printerExport" -ComputerList SRVRSH-001,SRVRSH-002,SRVRSH-003 -Credential -LogPath C:\Temp\Logs
Install-PrinterExportRemoteComputers.ps1 -printerExportFile "\\MV-SRV-PR01\Drivers\print_drivers.printerExport" -CSV "C:\scripts\computers.csv"
```
