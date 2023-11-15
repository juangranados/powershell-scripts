# Install or update any MSI on domain computers using GPO and PowerShell.

Install any MSI if previous or no version detected

Allows to update installed software and check the result of the installation using a log file.

Is a more complete alternative to the msi installation via gpo.

[Install-MSI](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Install%20MSI/Install-MSI.ps1): script to install MSI with GPO.

## Parameters

**InstallPath**  
MSI full installer path  
Example: \\FILESERVER-01\LibreOffice\LibreOffice_7.5.8_Win_x86-64.msi  

**SearchName**  
Name of the application to search for it in the registry in order to get the version installed.  
It does not need to be the exact name, but search by this name must return only one item or nothing.  
You can simulate the search using the command:  
Get-WmiObject  Win32_Product | Where-Object {$_.Name -like '*Office*'}  
  
**LogPath**  
Log path (optional). ComputerName.log file will be created.  
Example: \\FILESERVER-01\LibreOffice\Logs (Log will be saved to \\FILESERVER-01\JRE\computername.log)  

**MSIArguments**  
Parameters of MSI file.  
Warning! There seems to be a maximum number of 256 characters that can be used in the Script Parameters setting in a GPO Startup/Shutdown/Logon/Logoff PowerShell script.  
Sometimes scripts do not run even with fewer characters, so you can create a script that calls this script with all its parameters and run it via GPO.  
Optional, /qn is already applied.  

## Instructions for deploying LibreOffice MSI

1. Create shared folder to deploy LibreOffice: ```\\FILESERVER-01\LibreOffice```
2. Grant 'Authenticated users' read access to ```\\FILESERVER-01\LibreOffice```
3. Copy ```Install-MSI.ps1``` to your script folder ```\\FILESERVER-01\Scripts$\Install-MSI.ps1```
4. Create shared folder for logs: ```\\FILESERVER-01\LibreOffice\logs```
5. Grant 'Authenticated users' write access to ```\\FILESERVER-01\LibreOffice\logs```
6. Create a powershell script named _Install-LibreOffice.ps1_ to avoid GPO parameter limit and run the installation on background:
```
Start-Job -ScriptBlock {\\FILESERVER-01\Scripts$\Install-MSI.ps1 -InstallPath "\\FILESERVER-01\Software$\LibreOffice\LibreOffice_7.5.8_Win_x86-64.msi" -LogPath "\\FILESERVER-01\Software$\LibreOffice\Logs" -SearchName "LibreOffice" -MSIArguments "/norestart ALLUSERS=1 CREATEDESKTOPLINK=0 REGISTER_ALL_MSO_TYPES=0 REGISTER_NO_MSO_TYPES=1 ISCHECKFORPRODUCTUPDATES=0 QUICKSTART=0 ADDLOCAL=ALL UI_LANGS=en_US,ca,es"}
```
7. Create a GPO that runs _Install-LibreOffice.ps1_:
