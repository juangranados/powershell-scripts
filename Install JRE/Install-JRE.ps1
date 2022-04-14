<#
.SYNOPSIS
	Install Java Runtime Environment (jre) for Windows if previous or no version detected and uninstall previous versions.
.DESCRIPTION
	This script install Java Runtime Environment only if a previous version or no version was detected.
    You can avoid uninstall old versions with -NoUninstall switch.
.PARAMETER InstallPath
	Java Runtime Environment full installer path. It could be download from https://www.java.com/en/download/manual.jsp.
    Authenticated users must have read permissions over shared folder.
    Example: \\FILESERVER-01\Skype\jre-10.0.2_windows-x64_bin.exe
.PARAMETER LogPath
    Log path (optional). ComputerName.log file will be created.
    Authenticated users must have write permissions over log shared folder.
    Example: \\FILESERVER-01\JRE\Logs (Log will be saved to \\FILESERVER-01\JRE\computername.log)
.PARAMETER NoUninstall
    Default: false.
    Avoid uninstall previous versions before install Java Runtime Environment.
.PARAMETER x64
    Default: false.
    64 bit version is selected for installing: InstallPath contains a 64 Bit JRE Installer.
.EXAMPLE
    Install 64 bit Java Runtime Environment from network share, saving log in Log folder of network share and uninstall previous versions.
    Note: network share must have read permissions on "\\FILESERVER-01\JRE\" and write on "\\FILESERVER-01\JRE\Logs" for "Authenticated Users" group.
	InstallJRE.ps1 "\\FILESERVER-01\JRE\jre-10.0.2_windows-x64_bin.exe" "\\FILESERVER-01\JRE\Logs" -x64
.EXAMPLE
    Install 32 bit Java Runtime Environment from network share, saving log in Log folder of network share and uninstall previous versions.
    Note: network share must have read permissions on "\\FILESERVER-01\JRE\" and write on "\\FILESERVER-01\JRE\Logs" for "Authenticated Users" group.
	InstallJRE.ps1 "\\FILESERVER-01\JRE\jre-8u181-windows-i586.exe" "\\FILESERVER-01\JRE\Logs"
.NOTES 
	Author: Juan Granados 
	Date:   April 2022
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
    [switch]$NoUninstall,
    [Parameter(Mandatory = $false, Position = 3)] 
    [ValidateNotNullOrEmpty()]
    [switch]$x64
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

if (-not [string]::IsNullOrWhiteSpace($LogPath) -and $logPath.Chars($logPath.Length - 1) -eq '\') {
    $logPath = ($logPath.TrimEnd('\'))
}

if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
    if ($x64) {
        Start-Transcript -Path "$($LogPath)\$($env:COMPUTERNAME)_x64.log" | Out-Null
    }
    else {
        Start-Transcript -Path "$($LogPath)\$($env:COMPUTERNAME).log" | Out-Null
    }
}

if (-not [Environment]::Is64BitOperatingSystem -and $x64) {
    Write-Error "Error: 64 Bit version can not be installed on 32 Bit Operating System."
    if (-not [string]::IsNullOrWhiteSpace($LogPath)) { Stop-Transcript }
    Exit 1
}

$ErrorActionPreference = "Stop"
$Install = $true

try {
    $jreExeFile = Get-Item -Path $InstallPath
}
catch {
    Write-Error "Error accessing $($InstallPath)."
    Write-Error "$($Error[0])"
    if (-not [string]::IsNullOrWhiteSpace($LogPath)) { Stop-Transcript }
    Exit 1
}

$jreExeVersion = StringVersionToFloat $jreExeFile.VersionInfo.ProductVersion

Write-Host "JRE $($jreExeVersion) selected for installation."

if ($x64) {
    $jreInstalledVersion = Get-InstalledApps | Where-Object { $_.DisplayName -like '*Java*(64-bit)' }
}
else {
    $jreInstalledVersion = Get-InstalledApps | Where-Object { $_.DisplayName -like '*Java*' -and $_.DisplayName -notlike '*Java*(64-bit)' }
}

if ($jreInstalledVersion) {

    foreach ($installation in $jreInstalledVersion) {

        $version = StringVersionToFloat $installation.DisplayVersion
        
        if (($version -lt $jreExeVersion)) {
            Write-Host "JRE $($version) detected."
            if (-not $NoUninstall) {
                Write-Host "Unnistalling JRE $($version)."
                if ($installation.UninstallString -like '*msiexec*') {
                    Start-Process -FilePath cmd.exe -ArgumentList '/c', $installation.UninstallString, '/qn /norestart' -Wait
                }
                else {
                    Start-Process -FilePath cmd.exe -ArgumentList '/c', $installation.UninstallString, '/verysilent' -Wait
                }
            }
        }
        else {
            Write-Host "JRE $($version) or greater already installed."
            $Install = $false
        }
    }
} 
else {
    Write-Host "No JRE version detected."
}

if ($Install) {
    Write-Host "Starting installation of JRE $($jreExeVersion) and exiting."
    Start-Process $InstallPath -ArgumentList "/s SPONSORS=0" -Wait -PassThru | Wait-Process -Timeout 200
}

if (-not [string]::IsNullOrWhiteSpace($LogPath)) { Stop-Transcript }