# Bulk Change PrinterDriverAttributes for non Package-Aware printer drivers

[Right click here and select "Save link as" to download](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Bulk%20Change%20PrinterDriverAttributes%20for%20non%20Package-Aware%20printer%20drivers/ChangePrinterDriverAttributes.ps1)

![Error non Pacaked driver](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Bulk%20Change%20PrinterDriverAttributes%20for%20non%20Package-Aware%20printer%20drivers/do_you_trust_this_printer.png)

Since [KB3170455](https://support.microsoft.com/en-us/topic/ms16-087-security-update-for-windows-print-spooler-components-july-12-2016-afceb380-914b-500f-5aa2-904fe6d13817), deployed Printers via Group Policy does not install  non Package-Aware printer drivers automatically. Users are prompted with a message saying "Do yo trust this printer?"

More info and manual fix: [Group Policy Printer Issue – Print and Point Restrictions – KB3170455](https://www.richardwalz.com/group-policy-printer-issue-print-and-point-restrictions-kb3170455/)

You have to install printer driver manually in each computer. But,  for many old drivers, you can force GPO driver installation changing  PrinterDriverAttributes registry value to a odd number in print server.

This script, scan registry and changes all even PrinterDriverAttributes to allow driver installation by GPO. You have to restart print  server spooler after running it.

