# Windows Server Backup Email Report of Several Servers

[Right click here and select "Save link as" to download](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Delete%20reappearing%20printers%20that%20keeps%20comming%20back/DeletePrintersRDS.cmd)

Script to delete ghost printers that keeps coming back after deletion. 

**Warning: this script will delete all printers.**

You need to add psexec to system path. [Download PsTools](https://docs.microsoft.com/en-us/sysinternals/downloads/pstools) and extract it to C:\Windows\System32. Please, take a registry backup first!

*Example of undeletable printers on a RDS server*

![Example of ghost printers](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Delete%20reappearing%20printers%20that%20keeps%20comming%20back/printers.png)

