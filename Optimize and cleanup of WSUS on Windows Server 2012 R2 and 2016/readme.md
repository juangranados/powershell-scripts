# Optimize and cleanup of WSUS on Windows Server

Optimize WSUS DB using [official Microsoft SQL script](https://docs.microsoft.com/en-us/troubleshoot/mem/configmgr/reindex-the-wsus-database) and performs server maintenance.

WsusDBMaintenance.sql and script must be on the same path.

- <a href="https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Optimize%20and%20cleanup%20of%20WSUS%20on%20Windows%20Server%202012%20R2%20and%202016/Optimize-WSUS.ps1" download>Right click here and select "Save link as" to download script</a>

- <a href="https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Optimize%20and%20cleanup%20of%20WSUS%20on%20Windows%20Server%202012%20R2%20and%202016/WsusDBMaintenance.sql" download>Right click here and select "Save link as" to download WsusDBMaintenance.sql</a>

It saves log in script path: yyyyMMdd__Optimize-WSUS.log

Prerequisites for running Invoke-Sqlcmd

Download: [Microsoft® SQL Server® 2017 Feature Pack](https://www.microsoft.com/en-US/download/details.aspx?id=55992)

Install packages in order:

1. *SQLSysClrTypes.msi*
2. *SharedManagementObjects.msi*  
3. Run from PowerShell: `Install-Module -Name SqlServer`
