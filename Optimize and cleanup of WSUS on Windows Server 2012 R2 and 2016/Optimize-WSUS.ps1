
<#PSScriptInfo

.VERSION 1.0

.GUID c90bbc5c-9a05-471c-9020-741b5fc59d51

.AUTHOR Juan Granados

.COPYRIGHT 2021 Juan Granados

.TAGS WSUS Optimice Database SQL WID DB

.LICENSEURI https://raw.githubusercontent.com/juangranados/powershell-scripts/main/LICENSE

.PROJECTURI https://github.com/juangranados/powershell-scripts/tree/main/Optimize%20and%20cleanup%20of%20WSUS%20on%20Windows%20Server%202012%20R2%20and%202016

.RELEASENOTES Initial release

#>

<#
.SYNOPSIS
	Optimize WSUS DB and performs maintenance.

.DESCRIPTION
	Optimize WSUS and its DB using official Microsoft SQL script and performs server maintenance.
      
    Prerequisites for running Invoke-Sqlcmd

    1. Save T-SQL script as WsusDBMaintenance.sql from: https://docs.microsoft.com/en-us/troubleshoot/mem/configmgr/reindex-the-wsus-database
    
    2. Navigate to: https://www.microsoft.com/en-US/download/details.aspx?id=55992 and install SQLSysClrTypes.msi
    
    3. Run from PowerShell
        - Register-PackageSource -provider NuGet -name nugetRepository -location https://www.nuget.org/api/v2
        - Install-Package Microsoft.SqlServer.SqlManagementObjects  
    
    4. Run from PowerShell: Install-Module -Name SqlServer
    
    .NOTES
	Author: Juan Granados 
#>
#Requires -RunAsAdministrator
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition

Start-Transcript "$scriptPath\$(Get-Date -format "yyyyMMdd")_Optimize-WSUS.log"
function check-requisites {
    if (-not (Test-Path "$scriptPath\WsusDBMaintenance.sql")) {
        Write-Error "WsusDBMaintenance.sql missing. Save T-SQL script as WsusDBMaintenance.sql from: https://docs.microsoft.com/en-us/troubleshoot/mem/configmgr/reindex-the-wsus-database"
        Stop-Transcript
        Exit 1              
    }
    if (-not (Get-Module -ListAvailable SQLServer)) {
        Write-Error "Module SQLServer missing. Install it running: Install-Module -Name SqlServer"
        Stop-Transcript
        Exit 1
    }
    if (-not (Test-Path "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server 2017 Redist\SQL Server System CLR Types\CurrentVersion")) {
        Write-Error "SQLSysClrTypes 2017 is misssing. Install it from https://www.microsoft.com/en-US/download/details.aspx?id=55992"
        Stop-Transcript
        Exit 1
    }
    if (-not(Get-Package Microsoft.SqlServer.SqlManagementObjects)) {
        Write-Error "Microsoft.SqlServer.SqlManagementObjects is missing. Install it running 'Register-PackageSource -provider NuGet -name nugetRepository -location https://www.nuget.org/api/v2' and 'Install-Package Microsoft.SqlServer.SqlManagementObjects'"
        Stop-Transcript
        Exit 1
    }
}

check-requisites
Get-WsusServer | Invoke-WsusServerCleanup -CleanupObsoleteComputers -CleanupObsoleteUpdates -CleanupUnneededContentFiles -CompressUpdates -DeclineExpiredUpdates -DeclineSupersededUpdates
Invoke-Sqlcmd -ServerInstance "\\.\pipe\microsoft##WID\tsql\query" -InputFile "$scriptPath\WsusDBMaintenance.sql" -Verbose
Stop-Transcript