# File Server Access Audit Report with PowerShell

<a href="https://raw.githubusercontent.com/juangranados/powershell-scripts/main/File%20Server%20Access%20Audit%20Report%20with%20PowerShell/Get-AuditReport.ps1" download>Right click here and select "Save link as" to download</a>

This PowerShell script allows to audit several file servers and send a report in CSV and HTML by mail.

CSV file can be import on Excel to generate a File Audit Report.

HTML report can filter and sorting rows by server, time, user, file or operation (read, delete or write).

![HTML Report example](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/File%20Server%20Access%20Audit%20Report%20with%20PowerShell/7.png)

## How to configure auditing in Windows Server

### Enable audit in Windows Server

From local policy or group policy, navigate to Computer Configuration → Policies → Windows Settings → Security Settings → Local Policies → Audit Policy → Open Audit object access and select Success and Failure

![Enable Audit](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/File%20Server%20Access%20Audit%20Report%20with%20PowerShell/1.PNG)

### Enable audit in folder and subfolders

Right click on folder →Properties → Security → Advanced → Auditing → Add

![Enable auditing](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/File%20Server%20Access%20Audit%20Report%20with%20PowerShell/2.PNG)

Select domain users or user group and mark:

- Type: All.

- Basic Permissions: Read & execute, List folder contents, Read.

![Configure auditing](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/File%20Server%20Access%20Audit%20Report%20with%20PowerShell/5.PNG)

### Increase Security Log size

From local policy or group policy, browse to Computer Configuration → Policies → Administrative Templates → Windows Components → Event Log Service → Security → Specify the maximum log file size
Set the maximum log file size setting, for example 4194240 KB (4 GB).

![Increase security log screenshot](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/File%20Server%20Access%20Audit%20Report%20with%20PowerShell/6.PNG)
