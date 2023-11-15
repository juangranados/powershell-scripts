<#
.SYNOPSIS
    Install any MSI if previous or no version detected
    Allows to update installed software and check the result of the installation using a log file.
    Is a more complete alternative to the msi installation via gpo.
.DESCRIPTION
    This script any MSI only if a previous version or no version was detected.
    Allows to update installed software and check the result of the installation using a log file.
    Is a more complete alternative to the msi installation via gpo.
.PARAMETER InstallPath
    MSI full installer path
    Example: \\FILESERVER-01\LibreOffice\LibreOffice_7.5.8_Win_x86-64.msi
.PARAMETER SearchName
    Name of the application to search for it in the registry in order to get the version installed. 
    It does not need to be the exact name, but search by this name must return only one item or nothing.
    You can simulate the search using the command:
    Get-WmiObject  Win32_Product | Where-Object {$_.Name -like '*Office*'}
.PARAMETER LogPath
    Log path (optional). ComputerName.log file will be created.
    Example: \\FILESERVER-01\LibreOffice\Logs (Log will be saved to \\FILESERVER-01\JRE\computername.log)
.PARAMETER MSIArguments
    Parameters of MSI file. 
    Warning! There seems to be a maximum number of 256 characters that can be used in the Script Parameters setting in a GPO Startup/Shutdown/Logon/Logoff PowerShell script.
    Sometimes scripts do not run even with fewer characters, so you can create a script that calls this script with all its parameters and run it via GPO.
    Optional, /qn is already applied.
.EXAMPLE
    Install LibreOffice from network share, saving log in Log folder of network share.
    Note: network share must have read permissions on "\\FILESERVER-01\LibreOffice\" and write on "\\FILESERVER-01\LibreOffice\Logs" for "Authenticated Users" group in order to run it with GPO.
    Install-MSI.ps1 -InstallPath "\\FILESERVER-01\LibreOffice\LibreOffice_7.5.8_Win_x86-64.msi" -SearchName "LibreOffice" -LogPath "\\FILESERVER-01\LibreOffice\Logs" -MSIArguments "UI_LANGS=en_US,es"
.NOTES 
	Author: Juan Granados 
	Date:   November 2023
#>
Param(
    [Parameter(Mandatory = $true, Position = 0)] 
    [ValidateNotNullOrEmpty()]
    [string]$InstallPath,
    [Parameter(Mandatory = $true, Position = 1)] 
    [ValidateNotNullOrEmpty()]
    [string]$SearchName,
    [Parameter(Mandatory = $false, Position = 2)] 
    [ValidateNotNullOrEmpty()]
    [string]$LogPath,
    [Parameter(Mandatory = $false, Position = 3)] 
    [ValidateNotNullOrEmpty()]
    [string]$MSIArguments
)
#Requires -RunAsAdministrator
function Get-InstalledApps {
    if ([IntPtr]::Size -eq 4) {
        $regpath = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    }
    else {
        $regpath = @(
            'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
            'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
        )
    }
    Get-ItemProperty $regpath | . { process { if ($_.DisplayName -and $_.UninstallString) { $_ } } } | 
    Select-Object DisplayName, Publisher, InstallDate, DisplayVersion, UninstallString | 
    Sort-Object DisplayVersion
}

function StringVersionToFloat($version) {
    while (($version.ToCharArray() | Where-Object { $_ -eq '.' } | Measure-Object).Count -gt 1) {
        $aux = $version.Substring($version.LastIndexOf('.') + 1)
        $version = $version.Substring(0, $version.LastIndexOf('.')) + $aux
    }
    return [float]$version
}

function Get-MSIFileVersion ([System.IO.FileInfo]$msiFilePath) {
    try { 
        $WindowsInstaller = New-Object -com WindowsInstaller.Installer 
        $Database = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $Null, $WindowsInstaller, @($msiFilePath.FullName, 0)) 
        $Query = "SELECT Value FROM Property WHERE Property = 'ProductVersion'"
        $View = $database.GetType().InvokeMember("OpenView", "InvokeMethod", $Null, $Database, ($Query)) 
        $View.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $View, $Null) | Out-Null
        $Record = $View.GetType().InvokeMember( "Fetch", "InvokeMethod", $Null, $View, $Null ) 
        $Version = $Record.GetType().InvokeMember( "StringData", "GetProperty", $Null, $Record, 1 ) 
        return $Version
    }
    catch { 
        Write-Host "Failed to get MSI file version: {0}." -f $_
        Stop-Transcript
        Exit
    }     
}
function Get-MSIFileName ([System.IO.FileInfo]$msiFilePath) {
    try { 
        $WindowsInstaller = New-Object -com WindowsInstaller.Installer 
        $Database = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $Null, $WindowsInstaller, @($msiFilePath.FullName, 0)) 
        $Query = "SELECT Value FROM Property WHERE Property = 'ProductName'"
        $View = $database.GetType().InvokeMember("OpenView", "InvokeMethod", $Null, $Database, ($Query)) 
        $View.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $View, $Null) | Out-Null
        $Record = $View.GetType().InvokeMember( "Fetch", "InvokeMethod", $Null, $View, $Null ) 
        $ProductName = $Record.GetType().InvokeMember( "StringData", "GetProperty", $Null, $Record, 1 ) 
        return $ProductName
    }
    catch { 
        Write-Host "Failed to get MSI file ProductName: {0}." -f $_
        Stop-Transcript
        Exit
    }     
}

if (-not [string]::IsNullOrWhiteSpace($LogPath) -and $logPath.Chars($logPath.Length - 1) -eq '\') {
    $logPath = ($logPath.TrimEnd('\'))
}
if (Test-Path -Path $logPath) {
    if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
        try { Stop-Transcript | Out-Null } catch {}
        $logFilename = "$($env:COMPUTERNAME)_$($SearchName).txt" 
        $logFilename.Split([IO.Path]::GetInvalidFileNameChars()) -join '_' | Out-Null
        Start-Transcript -Path "$($LogPath)\$($logFilename)" | Out-Null
    }
}
else {
    Write-Host "Log path does not exists"
    Exit
}

$ErrorActionPreference = "Stop"
$Install = $true

try {
    $MSIFile = Get-Item -Path $InstallPath
}
catch {
    Write-Error "Error accessing $($InstallPath)."
    Write-Error "$($Error[0])"
    if (-not [string]::IsNullOrWhiteSpace($LogPath)) { Stop-Transcript }
    Exit 1
}

$MSIVersion = StringVersionToFloat $(Get-MSIFileversion $MSIFile)
$MSIName = $(Get-MSIFileName $MSIFile).ToString()

Write-Host "$($MSIName) version $($MSIVersion) selected for installation."

$widcardSearch = "*$($SearchName)*"
$installedVersion = Get-InstalledApps | Where-Object { $_.DisplayName -like $widcardSearch }

if ($installedVersion) {

    foreach ($installation in $installedVersion) {

        $version = StringVersionToFloat $installation.DisplayVersion
        
        if (($version -lt $MSIVersion)) {
            Write-Host "Version $($version) installed."
        }
        else {
            Write-Host "Version $($version) or greater already installed."
            $Install = $false
        }
    }
} 
else {
    Write-Host "No $($MSIName) detected."
}

if ($Install) {
    $msiLog = "$($env:LOCALAPPDATA)\msi_Installation.txt"
    if (Test-Path $msiLog) {
        Remove-Item $msiLog -Force
    }
    if (-not ([string]::IsNullOrEmpty($MSIArguments))) {
    	$MSIArgs = "/i $($InstallPath) /qn " + $MSIArguments + " /l $($msiLog)"
    } else {
    	$MSIArgs = "/i $($InstallPath) /qn /l $($msiLog)"
    }
    Write-Host "Starting installation of $($MSIName) version $($MSIVersion) and exiting."
    Write-Host "Running: msiexec.exe $($MSIArgs)"
    $exitCode = (Start-Process "msiexec.exe" -ArgumentList $MSIArgs -Wait -NoNewWindow -PassThru).ExitCode
    get-content $msiLog
    if ($exitCode -eq 0) {
        Write-Host "$($MSIName) installed sucessfully"
    }
    else {
        Write-Host "Error $exitCode installing $($MSIName)"
    }
}

if (-not [string]::IsNullOrWhiteSpace($LogPath)) { Stop-Transcript }
