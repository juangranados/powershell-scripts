<#
.SYNOPSIS
    Run Windows disks optimization.
.DESCRIPTION
    Run Windows disks optimization in all or selected drives.
.PARAMETER disks
    Disks to run optimization.
    Default: all.
    Example: "C:","D:","F:"
.PARAMETER LogPath
    Path where save log file.
    Default: Temp folder
.EXAMPLE
    Optimize all drives.
    Invoke-DiskOptimize.ps1
.EXAMPLE
    Optimize only C and D drives.
    Invoke-DiskOptimize.ps1 -disks "C:","D:"
.NOTES 
	Author:	Juan Granados
#>

Param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string[]]$disks = "all",
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$LogPath = $env:temp
)
#Requires -RunAsAdministrator

$ErrorActionPreference = "SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Stop"

$LogPath = $LogPath.TrimEnd('\')
if (-not (Test-Path $LogPath)) {
    Write-Host "Log path $($LogPath) not found"
    Exit (1)
}
Start-Transcript -path "$($LogPath)\$(get-date -Format yyyy_MM_dd)_$($env:COMPUTERNAME).txt"

try {
    if ($disks -eq "all") {
        $drives = get-wmiobject win32_volume | Where-Object { $_.DriveType -eq 3 -and $_.DriveLetter -and (Get-WMIObject Win32_LogicalDiskToPartition | Select-Object Dependent) -match $_.DriveLetter }
    }
    else {
        foreach ($disk in $disks) {
            if (-not ($disk -match '[A-Za-z]:')) {
                Write-Output "UNKNOWN: Error $($drive) is not a valid disk unit. Expected N:, where N is drive unit. Example C: or D: or F:"
                Exit(3)
            }
        }
        $drives = get-wmiobject win32_volume | Where-Object { $_.DriveType -eq 3 -and $_.DriveLetter -in $disks }
    }
    if (-not ($drives)) {
        Write-Output "UNKNOWN: No drives found with get-wmiobject win32_volume command"
        Exit(3)
    }
    foreach ($drive in $drives) {
        Optimize-Volume -Driveletter $($drive.DriveLetter) -Verbose
    }
    
}
catch {
    Write-Output "CRITICAL: $($_.Exception.Message)"
    Exit(2)
}
Stop-Transcript