# Install or update Java Runtime Environment 32 or 64 bits on domain computers using GPO and uninstall old versions.

Install Java Runtime Environment (JRE) if previous or no version was detected and uninstall previous versions for security reasons.

This script install Java only if a previous version or no version was detected.

You can skip uninstallation of old versions with -NoUninstall switch.

[Install-JRE](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Install%20Print%20Drivers%20Remotely/Install-JRE.ps1): to install latest JRE on Windows Computers with GPO.

[Download-JRE](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Install%20Print%20Drivers%20Remotely/Download-JRE.ps1): to run scheduled task on server in order to keep JRE updated in file share.

## Instructions

### Download and update daily

To download and keep Java Runtime Environment Updated

Create a scheduled task on server that nightly executes 'Download-JRE.ps1' to your server path.

```powershell
Download-JRE.ps1 -downloadPath "D:\Software\JRE" -logPath "C:\Logs"
```
Will download jrex64.exe and jrex32.exe to D:\Software\JRE

### Instructions for 64 bit installation

1. Create shared folder to deploy Java Runtime Environment: ```\\FILESERVER-01\JRE```
2. Grant 'Authenticated users' read access to ```\\FILESERVER-01\JRE```
3. Copy ```Install-JRE.ps1``` to ```\\FILESERVER-01\JRE\Install-JRE.ps1```
4. Create shared folder for logs: ```\\FILESERVER-01\JRE\logs```
5. Grant 'Authenticated users' write access to ```\\FILESERVER-01\JRE\logs```
6. Create a computer GPO that runs PowerShell Script:
```
Name: \\FILESERVER-01\JRE\Install-JRE.ps1
Parameters: \\FILESERVER-01\JRE\jrex64.exe \\FILESERVER-01\JRE\logs -x64
```

### Instructions for 32 bit installation

1. Create shared folder to deploy Java Runtime Environment: ```\\FILESERVER-01\JRE```
2. Grant 'Authenticated users' read access to ```\\FILESERVER-01\JRE```
3. Copy ```Install-JRE.ps1``` to ```\\FILESERVER-01\JRE\Install-JRE.ps1```
4. Create shared folder for logs: ```\\FILESERVER-01\JRE\logs```
5. Grant 'Authenticated users' ```write access to \\FILESERVER-01\JRE\logs```
6. Create a computer GPO that runs PowerShell Script:
```
Name: \\FILESERVER-01\JRE\Install-JRE.ps1
Parameters: \\FILESERVER-01\JRE\jrex32.exe \\FILESERVER-01\JRE\logs
```

## Parameters

### Install-JRE

**InstallPath:** JRE installer path.

**LogPath:** Path where save log file.

**NoUninstall:** Will not uninstall JRE previous versions.

**x64:** Install JRE 64 bits version. By default 32 bits version is installed.

```powershell
# Install JRE 32 bit
Install-JRE.ps1 -InstallPath \\FS01\JRE\jrex32.exe -LogPath \\FS01\JRE\Logs
```
```powershell
# Install JRE 32 bit without uninstall previous versions
Install-JRE.ps1 -InstallPath \\FS01\JRE\jrex32.exe -LogPath \\FS01\JRE\Logs -NoUninstall
```
```powershell
# Install JRE 64 bit
Install-JRE.ps1 -InstallPath \\FS01\JRE\jrex32.exe -LogPath \\FS01\JRE\Logs -x64
```
```powershell
# Install JRE 64 bit without uninstall previous versions
Install-JRE.ps1 -InstallPath \\FS01\JRE\jrex32.exe -LogPath \\FS01\JRE\Logs -x64 -NoUninstall
```

### Download-JRE

**DownloadPath:** Path where save jrex32.exe and jrex64.exe.

**LogPath:** Path where save log file.

```powershell
# Download JRE 32 and 64 bit to D:\Software\JRE and save log in C:\Logs
Download-JRE.ps1 -downloadPath "D:\Software\JRE" -logPath "C:\Logs"
```

