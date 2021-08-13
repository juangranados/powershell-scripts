# Optimize and cleanup of WSUS on Windows Server

Optimize WSUS DB using [official Microsoft SQL script](https://docs.microsoft.com/en-us/troubleshoot/mem/configmgr/reindex-the-wsus-database) and performs server maintenance.

WsusDBMaintenance.sql and script must be on the same path.

It saves log in script path: yyyyMMdd__Optimize-WSUS.log

Prerequisites for running Invoke-Sqlcmd

Download: [Microsoft® SQL Server® 2017 Feature Pack](https://www.microsoft.com/en-US/download/details.aspx?id=55992)

Install packages in order:

1. *SQLSysClrTypes.msi*
2. *SharedManagementObjects.msi*  
3. Run from PowerShell: `Install-Module -Name SqlServer`