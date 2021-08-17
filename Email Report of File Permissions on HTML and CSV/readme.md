# Email Report of File Permissions on HTML and CSV

[Right click here and select "Save link as" to download](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Email%20Report%20of%20File%20Permissions%20on%20HTML%20and%20CSV/Get-FolderPermissions.ps1)

Starting with a root folder, it generates a folders permissions  report. Number of sub folders examined depends on FolderDeep parameter.

Report is generated in CSV format and can be send attached via mail with a HTML report in the body.

![Screenshot](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Email%20Report%20of%20File%20Permissions%20on%20HTML%20and%20CSV/screenshot.png)

```powershell
<#
.SYNOPSIS
    Generate a folders permissions report.
.DESCRIPTION
    Starting with a root folder, it generates a folders permissions report. Number of subfolders examined depends on FolderDeep parameter.
    Report is generated in CSV format and can be send attached via mail with a html report in the body. 
.PARAMETER OutFile
    Path to store CSV file.
    Default .\Permissions.csv
.PARAMETER RootPath
    Folder to start checking permissions.
.PARAMETER FolderDeep
    Number of subfolders levels to check.
    Default 99.
.PARAMETER ObjectsIgnored
    Users or groups to ignore in report.
    Default NT AUTHORITY\SYSTEM,BUILTIN\Administrator
.PARAMETER InspectGroups
    List only users in report.
    Default $False
.PARAMETER SMTPServer
    Sets smtp server in order to sent an email with backup result. If leave blank, no email will be send.
.PARAMETER SMTPRecipient
    List of emails addresses which will receive the backup result separated by commas.
.PARAMETER SMTPSender
    Email address which will send the backup result.
.PARAMETER SMTPUser
    Username in case of smtp server requires authentication.
.PARAMETER SMTPPassword
    Password in case of smtp server requires authentication.
.PARAMETER SMTPSSL
    Use of SSL in case of smtp server requires SSL.
    Default: $False
.PARAMETER SMTPPort
    Port to connect to smtp server.
    Default: 25
.EXAMPLE
    Get-FoldersPermissions -RootPath "D:\Data\Departments" -FolderDeep 2 -SMTPServer "mail.server.com" -SMTPRecipient "megaboss@server.com","support@server.com" -SMTPSender "reports@server.com"
.NOTES 
    Author: Juan Granados 
#>
```

