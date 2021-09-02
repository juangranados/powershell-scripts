# SQL Server Backup

[Right click here and select "Save link as" to download](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/SQL%20Server%20Backup/BackupSQL.ps1)

This script perform full and log backup of a SQL Server databases into a zip file.

You can specify what databases will be saved and the remote location of backup, also it can delete old backup files to save disk space, e.g keeping one week of backups. It can send an email report with backup result.

This script can be executed in the SQL Server you want to backup, but if you want to run it from a computer without SQL Server installed:

1. Download and install SQLSysClrTypes.msi from: [Microsoft® SQL Server® 2017 Feature Pack](https://www.microsoft.com/en-US/download/details.aspx?id=55992)
2. Run from PowerShell: 
  - `Register-PackageSource -provider NuGet -name nugetRepository -location https://www.nuget.org/api/v2`
  - `Install-Package Microsoft.SqlServer.SqlManagementObjects` in case of error of circular dependency add `-SkipDependencies`
3. Run from PowerShell: `Install-Module -Name SqlServer`

#### *Examples*

Backup default instance databases to a network share

`BackupSQL.ps1 -BackupDirectory \\FS-SERVER01\BackupSQL`

Backup Windows Internal Database (WID) to a network share

`BackupSQL.ps1 -BackupDirectory "\\MV0SRV-C01\Backups" -Instance "\\.\pipe\MICROSOFT##WID\tsql\query"`

Backup default instance databases to a network share and send an email with result using gmail

`BackupSQL.ps1 -BackupDirectory \\FS-SERVER01\BackupSQL -SMTPServer smtp.gmail.com -Recipient jgranados@contoso.com,administrator@contoso.com -Sender backupSQL@gmail.com -Username backupSQL@gmail.com -Password Pa$$W0rd -SSL True -Port 587`

Backup named instance databases to a network share

`BackupSQL.ps1 -BackupDirectory \\FS-SERVER01\BackupSQL -Instance SQLSVR01\BKUPEXEC`

Backup default instance databases to a network share, delete from network share files older than a week and write result in Windows Application Event

`BackupSQL.ps1 -BackupDirectory \\FS-SERVER01\BackupSQL -RetainDays 7 -WriteEvent True`

Backup only specified databases of a named instance

`BackupSQL.ps1 -Instance SQLSVR01\BKUPEXEC -DataBases BEDB,msdb,model`
