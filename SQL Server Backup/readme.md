# SQL Server Backup

[Right click here and select "Save link as" to download](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/SQL%20Server%20Backup/BackupSQL.ps1)

This script perform full and log backup of a SQL Server databases into a zip file. It needs pscx module because native command `Compress-Archive` maximum file size is 2 GB.

To install pscx run as administrator `Install-Module -Name Pscx`

You can specify what databases will be saved and the remote location of backup, also it can delete old backup files to save disk space, e.g keeping one week of backups. It can send an email report with backup result.

This script must be executed in the SQL Server you want to backup.

#### *Examples*

Backup default instance databases to a network share
`BackupSQL.ps1 -BackupDirectory \\FS-SERVER01\BackupSQL`

Backup default instance databases to a network share and send an email with result using gmail
`BackupSQL.ps1 -BackupDirectory \\FS-SERVER01\BackupSQL -SMTPServer smtp.gmail.com -Recipient jgranados@contoso.com,administrator@contoso.com -Sender backupSQL@gmail.com -Username backupSQL@gmail.com -Password Pa$$W0rd -SSL True -Port 587`

Backup named instance databases to a network share
`BackupSQL.ps1 -BackupDirectory \\FS-SERVER01\BackupSQL -Instance SQLSVR01\BKUPEXEC`

Backup default instance databases to a network share, delete from network share files older than a week and write result in Windows Application Event
`BackupSQL.ps1 -BackupDirectory \\FS-SERVER01\BackupSQL -RetainDays 7 -WriteEvent True`

Backup only specified databases of a named instance
`BackupSQL.ps1 -Instance SQLSVR01\BKUPEXEC -DataBases BEDB,msdb,model`