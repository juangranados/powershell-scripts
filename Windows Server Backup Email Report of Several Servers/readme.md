# Windows Server Backup Email Report of Several Servers

[Right click here and select "Save link as" to download](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Windows%20Server%20Backup%20Email%20Report%20of%20Several%20Servers/Get-WSBReport.ps1)

Performs a Windows Server Backup HTML report of a servers list and send it via email.

![Screenshot](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Windows%20Server%20Backup%20Email%20Report%20of%20Several%20Servers/wsb.PNG)

It gets server list from a file called 'Servers.txt'. [Right click here and select "Save link as" to download an example of 'Servers.txt'](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Windows%20Server%20Backup%20Email%20Report%20of%20Several%20Servers/Servers.txt)

Requires Windows Server Backup Command Line Tools installed on remote servers. On each server, open PowerShell console as administrator and run `Add-WindowsFeature Backup-Tools`

*Example*

```powershell
Get-WSBReport.ps1 -ServerList C:\Scripts\servers_contoso.txt -HtmlReport \\SERVER1\Reports\ -SMTPServer mail.contoso.com -Sender soporte@contoso.com -Recipient jgranados@contoso.com,administrador@contoso.com -Username contoso\jgranados -Password P@ssw0rd
```

*Full description*

```powershell
<#
.SYNOPSIS
    This script collect information about Windows Server Backup on a list of servers.
    Requires Windows Server Backup Command Line Tools installed on remote servers (Add-WindowsFeature Backup-Tools)
.DESCRIPTION
    This script collect information about Windows Server Backup on a list of servers and show results on console. It has the posibility of generate an send an html report with backup results.
    Usage: Get-WSBReport.ps1 [-Servers <string>] [-HtmlReport <string>] [SMTPServer <string>] [Recipient <string[]>] [Sender <string>] [Username <string>] [Password <string>] [-SSL <True | False>] [Port <int>]
.PARAMETER Servers
   Full path of a file containing the list of servers to check Windows Server Backup Status   
   Default "C:\Scripts\Servers.txt"
.PARAMETER HtmlReport
   Folder to store html report file.
   Default "C:\Scripts\"
.PARAMETER SMTPServer
    Sets smtp server in order to sent an email with backup result.
    Default: None
.PARAMETER Recipient
    List of emails addresses which will receive the backup result separated by commas.
    Default: None
.PARAMETER Sender
    Email address which will send the backup result.
    Default: None
.PARAMETER Username
    Username in case of smtp server requires authentication.
    Default: None
.PARAMETER Password
    Password in case of smtp server requires authentication.
    Default: None
.PARAMETER SSL
    Use of SSL in case of smtp server requires SSL.
    Default: False
.PARAMETER Port
    Port to connect to smtp server.
    Default: 25
.EXAMPLE
    .\Get-WSBReport.ps1 -ServerList C:\Scripts\servers_contoso.txt -HtmlReport \\SERVER1\Reports\ -SMTPServer mail.contoso.com -Sender soporte@contoso.com -Recipient jgranados@contoso.com,administrador@contoso.com -Username contoso\jgranados -Password P@ssw0rd
.NOTES
    Author: Juan Granados
#>
```

