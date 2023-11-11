<#
.SYNOPSIS
    Install LibreOffice for Windows if previous or no version detected and uninstall previous versions.
.DESCRIPTION
    This script LibreOffice only if a previous version or no version was detected.
    You can avoid uninstall old versions with -NoUninstall switch.
.PARAMETER InstallPath
    LibreOffice full installer path. It could be download from https://www.libreoffice.org/download/download-libreoffice/
    Authenticated users must have read permissions over shared folder.
    Example: \\FILESERVER-01\LibreOffice\LibreOffice_7.5.8_Win_x86-64.msi
.PARAMETER LogPath
    Log path (optional). ComputerName.log file will be created.
    Authenticated users must have write permissions over log shared folder.
    Example: \\FILESERVER-01\LibreOffice\Logs (Log will be saved to \\FILESERVER-01\JRE\computername.log)
.PARAMETER MSIArguments
    Parameters of LibreOffice installation options (optional).
    Default "/qn /norestart ALLUSERS=1 CREATEDESKTOPLINK=0 REGISTER_ALL_MSO_TYPES=0 REGISTER_NO_MSO_TYPES=1 ISCHECKFORPRODUCTUPDATES=0 QUICKSTART=0 ADDLOCAL=ALL UI_LANGS=en_US,ca,es"
.PARAMETER NoUninstall
    Default: false.
    Avoid uninstall previous versions before install.
.EXAMPLE
    Install LibreOffice from network share, saving log in Log folder of network share.
    Note: network share must have read permissions on "\\FILESERVER-01\LibreOffice\" and write on "\\FILESERVER-01\LibreOffice\Logs" for "Authenticated Users" group.
    InstallLibreOffice.ps1 "\\FILESERVER-01\LibreOffice\LibreOffice_7.5.8_Win_x86-64.msi" "\\FILESERVER-01\LibreOffice\Logs"
.NOTES 
	Author: Juan Granados 
	Date:   November 2023
#>
Param(
    [Parameter(Mandatory = $true, Position = 0)] 
    [ValidateNotNullOrEmpty()]
    [string]$InstallPath,
    [Parameter(Mandatory = $false, Position = 1)] 
    [ValidateNotNullOrEmpty()]
    [string]$LogPath,
    [Parameter(Mandatory = $false, Position = 2)] 
    [ValidateNotNullOrEmpty()]
    [string]$MSIArguments= "/qn /norestart ALLUSERS=1 CREATEDESKTOPLINK=0 REGISTER_ALL_MSO_TYPES=0 REGISTER_NO_MSO_TYPES=1 ISCHECKFORPRODUCTUPDATES=0 QUICKSTART=0 ADDLOCAL=ALL UI_LANGS=en_US"
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
    } catch { 
        throw "Failed to get MSI file version: {0}." -f $_
    }     
}

if (-not [string]::IsNullOrWhiteSpace($LogPath) -and $logPath.Chars($logPath.Length - 1) -eq '\') {
    $logPath = ($logPath.TrimEnd('\'))
}
try{Stop-Transcript} catch{}
if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
        Start-Transcript -Path "$($LogPath)\$($env:COMPUTERNAME).txt" | Out-Null
}

$ErrorActionPreference = "Stop"
$Install = $true

try {
    $loExeFile = Get-Item -Path $InstallPath
}
catch {
    Write-Error "Error accessing $($InstallPath)."
    Write-Error "$($Error[0])"
    if (-not [string]::IsNullOrWhiteSpace($LogPath)) { Stop-Transcript }
    Exit 1
}

$loExeVersion = StringVersionToFloat $(Get-MSIFileversion $loExeFile)

Write-Host "LibreOffice $($loExeVersion) selected for installation."

$loInstalledVersion = Get-InstalledApps | Where-Object { $_.DisplayName -like '*LibreOffice*' }

if ($loInstalledVersion) {

    foreach ($installation in $loInstalledVersion) {

        $version = StringVersionToFloat $installation.DisplayVersion
        
        if (($version -lt $loExeVersion)) {
            Write-Host "LibreOffice $($version) installed."
        }
        else {
            Write-Host "LibreOffice $($version) or greater already installed."
            $Install = $false
        }
    }
} 
else {
    Write-Host "No LibreOffice version detected."
}

if ($Install) {
    Write-Host "Starting installation of LibreOffice $($loExeVersion) and exiting."
    $MSIArgs = "/i $($InstallPath) " + $MSIArguments + " /l $($env:LOCALAPPDATA)\InstallLibreOfficeInstallation.txt"
    $exitCode = (Start-Process "msiexec.exe" -ArgumentList $MSIArgs -Wait -NoNewWindow -PassThru).ExitCode
    get-content "$($env:LOCALAPPDATA)\InstallLibreOfficeInstallation.txt"
    if ($exitCode -eq 0) {
        Write-Host "LibreOffice installed sucessfully"
    }
    else {
        Write-Host "Error $exitCode installing LibreOffice"
    }
}

if (-not [string]::IsNullOrWhiteSpace($LogPath)) { Stop-Transcript }
