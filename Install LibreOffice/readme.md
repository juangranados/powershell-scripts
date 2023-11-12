# Install or update LibreOffice on domain computers using GPO and PowerShell.

Install LibreOffice if previous or no version was detected.

This script install LibreOffice only if a previous version or no version was detected.

[Install-LibreOffice](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Install%20LibreOffice/Install-LibreOffice.ps1): script to install LibreOffice with GPO.

## Instructions

1. Create shared folder to deploy LibreOffice: ```\\FILESERVER-01\LibreOffice```
2. Grant 'Authenticated users' read access to ```\\FILESERVER-01\LibreOffice```
3. Copy ```Install-LibreOffice.ps1``` to ```\\FILESERVER-01\LibreOffice\Install-LibreOffice.ps1```
4. Create shared folder for logs: ```\\FILESERVER-01\LibreOffice\logs```
5. Grant 'Authenticated users' write access to ```\\FILESERVER-01\LibreOffice\logs```
6. Create a computer GPO that runs PowerShell Script:
```
Name: \\FILESERVER-01\JRE\Install-LibreOffice.ps1
Parameters: -InstallPath "\\INFSRV003\Software$\LibreOffice\LibreOffice_7.5.8_Win_x86-64.msi" -LogPath"\\INFSRV003\Software$\LibreOffice\Logs"
```
## Parameters

**InstallPath:** LibreOffice installer path.

**LogPath:** Path where save log file.

**MSIArguments:** LibreOffice installation options (optional).

```powershell
# Install LibreOffice
\\INFSRV003\Software$\LibreOffice\Install-LibreOffice.ps1 -InstallPath "\\INFSRV003\Software$\LibreOffice\LibreOffice_7.5.8_Win_x86-64.msi" -LogPath"\\INFSRV003\Software$\LibreOffice\Logs" 
```
