# Windows Maintenance

[Right click here and select "Save link as" to download](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Windows%20Mainteinance/Win-Mnt.ps1)

Performs several commands to check and repair a Windows Computer, Server or Workstation.

## Parameters

- **all** runs all commands.
- **antivirus** runs all antivirus: Windows Defender, Karspersky, McAfee, ClamAV and Adaware.
- **sfc** runs SFC /scannow.
- **dism** runs DISM /Online /Cleanup-Image /RestoreHealth
- **wmi** runs Winmgmt /salvagerepository
- **mof** runs mofcomp.exe from C:\Windows\System32\wbem\AutoRecover
- **defrag** defrag drives if required (Fragmentation > 10%)
- **update** install Microsoft updates (except drivers)
- **defender** runs Windows Defender Update and Quick Scan
- **adaware** runs adaware update and quick/boot Scan
- **kas** runs Kaspersky Virus Removal Tool
- **clamav** runs ClamAV full scan of C:\
- **mcafee** runs McAfee Stinger
- **LogPath** path where save log file.
  *Default: My Documents*

## Examples

Runs all commands.

```powershell
Win-Mnt.ps1 -all
```

Runs all antivirus.

```powershell
Win-Mnt.ps1 -antivirus
```

 Runs sfc and defrag with custom log.

```powershell
Win-Mnt.ps1 -sfc -defrag -logPath "\\INFSRV001\Scripts$\Mainteinance\Logs"
```

Runs sfc, dism and adaware scan.

```powershell
Win-Mnt.ps1 -sfc -dism -adaware
```

