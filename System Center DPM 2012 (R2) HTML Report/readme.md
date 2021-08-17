# Bulk Change PrinterDriverAttributes for non Package-Aware printer drivers

[Right click here and select "Save link as" to download](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Bulk%20Change%20PrinterDriverAttributes%20for%20non%20Package-Aware%20printer%20drivers/ChangePrinterDriverAttributes.ps1)

This script Create and send a protection group backup report of System Center Data Protection Manager (DPM) 2012 (R2) Servers based on the recovery points and protection groups.

![Report Screenshot](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/System%20Center%20DPM%202012%20(R2)%20HTML%20Report/report_screenshot.PNG)

*Usage*

`DPMProtectionGroupReport.ps1 [-DPMServers <string[]>] [-ProtectionGroups <string[]>] [-MinLastBackupHours <int>] [-MinRecoveryPoints <int>] [ReportLocation <string>] [SMTPServer <string>] [Recipient <string[]>] [Sender <string>] [Username <string>] [Password <string>] [-SSL <True | False>] [Port <int>]`

