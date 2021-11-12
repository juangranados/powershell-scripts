# Remote Computer Update

<a href="https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Remote%20Computer%20Update/Update-Computer.ps1" download>Right click here and select "Save link as" to download</a>

Runs Windows Update and software update using RuckZuck on local or remote computer. 

In order to run in remote computer it has to be executed from [PsExec](https://docs.microsoft.com/en-us/sysinternals/downloads/psexec). See examples below.

It uses [RZGet](https://github.com/rzander/ruckzuck/releases) to update computer software.

![Screenshot](https://github.com/juangranados/powershell-scripts/raw/main/Remote%20Computer%20Update/screenshot.png)

## Parameters

### logPath

Log file path.
Default Documents
Example: "\\ES-CPD-BCK02\scripts\ComputerUpdate\Log"

### scheduleReboot

Reboot wil be scheduled if needed.
Default: false

### rebootHours

Number of hours after finish update to reboot computer.
Default: 2

### rebootNow

Reboots after finish update.
Default: false

### rebootMessage

Shows a message to user.
Default: none

### RZGetPath

RZGet.exe path.
Example: \\SRV-FS05\RZGet\rzget.exe
If path not found RZGet will not be called unless you set -downloadRZGet switch and it will be downloaded to this path.
Default: none

### RZGetArguments

RZGet.exe Arguments.
Default: update --all

### downloadRZGet

Download RZGet.exe latest version to RZGetPath
Default: false

RZGet.exe Arguments.
Default: update --all

## Examples

### Run remotely

#### Harcode parameters on script and run PsExec. 

```bash
psexec.exe -s \\ComputerName powershell.exe -ExecutionPolicy Bypass -file \\ES-CPD-BCK02\scripts\WindowsUpdate\Update-Computer.ps1
```

#### Use an script to run with parameters.

[LaunchRemote.cmd](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Remote%20Computer%20Update/LaunchRemote.cmd): run [LaunchRemote.ps1](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Remote%20Computer%20Update/LaunchRemote.ps1) with PsExec.

[LaunchRemote.ps1](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Remote%20Computer%20Update/LaunchRemote.ps1): run [Update-Computer.ps1](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Remote%20Computer%20Update/Update-Computer.ps1) with arguments.

### Run locally

```powershell
Update-Computer.ps1 -$logPath "\\ES-CPD-BCK02\scripts\WindowsUpdate\Log" -scheduleReboot -rebootHours 2 -rebootMessage "Computer will reboot in two hours. You can reboot now or it will reboot later" -RZGetPath "\\ES-CPD-BCK02\scripts\WindowsUpdate\RZGet.exe" $RZGetArguments 'update "7-Zip" "Google Chrome" "Notepad++(x64)" "AdobeReader DC"'
```

