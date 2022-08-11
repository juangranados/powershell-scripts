# Install or update or uninstall software on computers using GPO/Intune/PSExec with PowerShell.

This script Install/Update/Uninstall software using RZGet repository from https://ruckzuck.tools/.

[Right click here and select "Save link as" to download](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Install%20Software/Install-Software.ps1)

![screenshot](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Install%20Software/locally.png)

## Parameters

### tempFolder
* Folder to download installers
* Default: "C:\temp\InstallSoftware"

### software
* List of software to install. Check https://ruckzuck.tools/Home/Repository
* Default: None
* Example: "7-Zip","Notepad++","Edge","3CXPhone for Windows","Google Chrome","Teams","Postman"

###  logFolder
* Log file path.
* Default: "C:\temp\InstallSoftware"
* Example: "\\ES-CPD-BCK02\scripts\InstallSoftware\Log"

###  uninstall
* Uninstall software if it is already installed

###  checkOnly
* Check if software list is already installed

###  runAsAdmin
* Check if script is elevated and exit if false.

###  sleep
* Amount of seconds to sleep.
* When running as logon script sometimes Windows profile is not ready and installation fails.
* Default 0 

## Examples 

Install ```7-Zip, Notepad++, Edge, 3CXPhone for Windows, Google Chrome, Teams and Postman``` if not installed or if is outdated and save log in ```\\ES-CPD-BCK02\scripts\InstallSoftware\Log``` using ```C:\temp\InstallSoftware``` as temp folder.

```powershell
Install-Software.ps1 -tempFolder C:\temp\InstallSoftware -software "7-Zip","Notepad++","Edge","3CXPhone for Windows","Google Chrome","Teams","Postman" -logFolder "\\ES-CPD-BCK02\scripts\InstallSoftware\Log"
```
Uninstall ```Teams and Postman``` if installed and save log in ```\\ES-CPD-BCK02\scripts\InstallSoftware\Log```.

```powershell
Install-Software.ps1 -software "Teams","Postman" -logFolder "\\ES-CPD-BCK02\scripts\InstallSoftware\Log" -uninstall
```

Check if ```Teams and Postman``` are installed and save log in ```\\ES-CPD-BCK02\scripts\InstallSoftware\Log```.

```powershell
Install-Software.ps1 -software "Teams","Postman" -logFolder "\\ES-CPD-BCK02\scripts\InstallSoftware\Log" -checkOnly
```
### Instructions for GPO deployment

1. Create shared folder to save script file Install-Software.ps1: ```\\FILESERVER-01\Install-Software```
2. Grant 'Authenticated users' read access to ```\\FILESERVER-01\Install-Software```
3. Copy ```Install-Software.ps1``` to ```\\FILESERVER-01\Install-Software\Install-Software.ps1```
4. Create shared folder for logs: ```\\FILESERVER-01\Install-Software\logs```
5. Grant 'Authenticated users' write access to ```\\FILESERVER-01\Install-Software\logs```
6. Create a computer GPO that runs PowerShell Script:
```
Name: \\FILESERVER-01\Install-Software\Install-Software.ps1
Parameters: -software "7-Zip","Notepad++","Edge" -logFolder '\\FILESERVER-01\Install-Software\logs' -sleep 60
```
### Instructions for Intune deployment
As Intune does not allow parameters, you must harcode ```software``` parameter.

Change line 35 by:
```powershell
[Parameter(Mandatory = $false)]
```

And line 36 with the software to install
```powershell
[string[]]$software="7-Zip","Notepad++","Edge"
```

### Instructions for Remote Execution
You can download PSExec from [here](https://docs.microsoft.com/en-us/sysinternals/downloads/psexec).

**Run PSExec**

psexec.exe -s \\```COMPUTER_NAME``` powershell.exe "-Command" "\\```NETWORKSHARE\Install-Software.ps1 -parameters```"

Example
```
psexec.exe -s \\WK-MARKETING01 powershell.exe "-Command"  "\\FILESERVER-01\Install-Software\Install-Software.ps1 -software '7-Zip','Notepad++','Edge' -logFolder '\\ES-CPD-BCK02\scripts\InstallSoftware\Log'"
```
![screenshot](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Install%20Software/psexec.png)