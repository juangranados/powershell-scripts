# Install software on multiple computers remotely with PowerShell

This script install software remotely in a group of computers and retry the installation in case of error. It uses PowerShell to perform remote installation.

*Screenshot of TightVNC remote installation on 77 computers.*

![screenshot](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Install%20Software%20Remotely/Screenshot.png)

As it uses Powershell to perform the remote installation, target computer must allow Windows PowerShell Remoting. Script can try to enable Windows PowerShell Remoting using Microsoft Sysinternals Psexec with the paramenter -EnablePSRemoting. If PSExec is not found on computer, script asks to the user for download it and extract to system folder.

Please, read parameter description carefully before running.

**AppPath:** Path to the application executable, It can be a network or local path because entire folder will be copied to remote computer before installing and deleted after installation.   

​	*Example:* `C:\Software\TeamViewer\TeamvieverHost.msi` *(Folder TeamViewer will be copied to remote computer before run ejecutable).*

**AppArgs:** Application arguments to perform silent installation.

​	*Example:* `/S /R settings.reg`

**LocalPath:** Local path of the remote computer where copy application directory.

​	*Default:* `C:\temp`

**Retries:** Number of times to retry failed installations.

​	*Default: 5.*

**TimeBetweenRetries:** Seconds to wait before retrying failed installations.

​	*Default: 60*

**ComputerList:** List of computers in install software. You can only use one source of target computers: ComputerList, OU or CSV.

​	*Example:* `Computer001,Computer002,Computer003` *(Without quotation marks)*

**OU:** OU containing computers in which install software. RSAT for AD module for Powershell must be installed in order to query AD. If you run script from a Domain Controller, AD module for PowerShell is already enabled.

​	*Example:* `OU=Test,OU=Computers,DC=CONTOSO,DC=COM`

**CSV:** CSV file containing computers in which install software.

​	*Example:* `C:\Scripts\Computers.csv`

​	*CSV Format:*

​		`Name`

​		`Computer001`

​		`Computer002`

​		`Computer003.`

**LogPath:** Path where save log file.

*Default: My Documents.*

**Credential:** Script will ask for an account to perform remote installation.

**EnablePSRemoting:** Try to enable PSRemoting on failed computers using Psexec. Psexec has to be on system path. If PSExec is not found. Script ask to download automatically PSTools and copy them to C:\Windows\System32.

**AppName:** App name as shown in registry to check if app is installed on remote computer and not reinstall it.

You can check app name on a computer with it installed looking at:  

​	`HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\`

​	`HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\`

​	*Example: 'TightVNC'*

​	*Default: None*

**AppVersion:** App name as shown in registry to check if app is installed on remote computer and not reinstall it. If not specified and AppName has a value, version will be ignored.

You can check app version on a computer with it installed looking at:

​	`HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\`

​	`	HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\`

​	*Example: '2.0.8.1'*

​	*Default: all*

**WMIQuery:** WMI Query to execute in remote computers. Software will be installed if query returns any value.

​	*Example:* `select * from Win32_Processor where DeviceID="CPU0" and AddressWidth="64"' (64 bit computers)`

​	*Example:* `select * from Win32_Processor where DeviceID="CPU0" and AddressWidth="32"' (32 bit computers)`

​	*Default: None*

## Examples

Thanks Terence Luk for his amazing post:

[http://terenceluk.blogspot.com/2019/02/using-installsoftwareremotelyps1-to.html](https://web.archive.org/web/20200318211458/http://terenceluk.blogspot.com/2019/02/using-installsoftwareremotelyps1-to.html)

Other examples:

```powershell
#Install TightVNC mirage Driver using computer list with different credentials checking before if it is installed and computers have 32 bits, enabling PSRemoting on connection error. 
.\InstallSoftwareRemotely.ps1 ` 
-AppPath 'C:\Scripts\TightVNC\dfmirage-setup-2.0.301.exe' ` 
-AppArgs '/verysilent /norestart' ` 
-ComputerList PC01,PC03,PC12,PC34,PC43,PC50 ` 
-Retries 2 ` 
-AppName 'DemoForge Mirage Driver for TightVNC 2.0' ` 
-AppVersion '2.0' ` 
-WMIQuery 'select * from Win32_Processor where DeviceID="CPU0" and AddressWidth="32"' ` 
-EnablePSRemoting ` 
-Credential 
 
#Install TightVNC on 64 bits computers in a OU checking before if it is installed and enablig PSRemoting on connection error. 
.\InstallSoftwareRemotely.ps1 ` 
-AppPath 'C:\Scripts\TightVNC\tightvnc-2.8.8-gpl-setup-64bit.msi' ` 
-AppArgs '/quiet /norestart ADDLOCAL="Server" SERVER_REGISTER_AS_SERVICE=1 SERVER_ADD_FIREWALL_EXCEPTION=1 SERVER_ALLOW_SAS=1 SET_USEVNCAUTHENTICATION=1 VALUE_OF_USEVNCAUTHENTICATION=1 SET_PASSWORD=1 VALUE_OF_PASSWORD=P@ssw0rd SET_USECONTROLAUTHENTICATION=1 VALUE_OF_USECONTROLAUTHENTICATION=1 SET_CONTROLPASSWORD=1 VALUE_OF_CONTROLPASSWORD=P@ssw0rd' ` 
-OU 'OU=Central,OU=Computers,DC=Contoso,DC=local' ` 
-Retries 2 ` 
-AppName 'TightVNC' ` 
-AppVersion '2.8.8.0' ` 
-WMIQuery 'select * from Win32_Processor where DeviceID="CPU0" and AddressWidth="64"' ` 
-EnablePSRemoting 
 
#Upgrade VMware Tools in datacenter OU 
.\InstallSoftwareRemotely.ps1 -AppPath "\\mv-srv-fs01\Software\VMware Tools\setup64.exe" -AppArgs '/s /v "/qn reboot=r"' -OU "OU=Datacenter,DC=CONTOSO,DC=COM" 
```
