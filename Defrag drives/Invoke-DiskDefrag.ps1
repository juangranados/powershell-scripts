<#
.SYNOPSIS
    Run Windows disks defragmentation.
.DESCRIPTION
    Run Windows disks defragmentation in all or selected drives.
.PARAMETER disks
    Disks to run defragmentation.
    Default: all.
    Example: "C:","D:","F:"
.PARAMETER defragPercentage
    Percentage of fragmentation in order to run defragmentation.
    Default: 10
.PARAMETER forceDefrag
    Defrag disks if free space is low.
    Default: false
.EXAMPLE
    Defrag all drives if they are 10% fragmented.
    Invoke-DiskDefrag.ps1
.EXAMPLE
    Defrag only C and D drives if they are 20% fragmented.
    Invoke-DiskDefrag.ps1 -disks "C:","D:" -defragPercentage 20
.EXAMPLE
    Defrag C: drive if they are 10% fragmented. It runs disk defragmentation even C: disk free space is low.
    Invoke-DiskDefrag.ps1 -disks "C:" -forceDefrag
.NOTES 
	Author:	Juan Granados
#>

Param(
    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 100)]
    [int]$defragPercentage = 10,
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string[]]$disks = "all",
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$LogPath = $env:temp,
    [Parameter()]
    [switch]$forceDefrag
)
#Requires -RunAsAdministrator
$ErrorActionPreference = "Stop"
$global:output = ""

Function Invoke-DiskDefragmentation($diskToDefrag) {
    
    if ([Environment]::UserInteractive) {
        Start-Process "C:\Windows\System32\dfrgui.exe"
    }
    if ($forceDefrag) {
        Write-Host "Forcing $($diskToDefrag.DriveLetter) defragmentation"
        $result = $diskToDefrag.Defrag($true)
    }
    else {
        Write-Host "Performing $($diskToDefrag.DriveLetter) defragmentation"
        $result = $diskToDefrag.Defrag($false)    
    }
    if ($result.ReturnValue -eq 0) {
        Write-Host "Defragmentation successful"
        Write-Host "Current fragmentation is $($result.DefragAnalysis.FilePercentFragmentation)"
        $diskToDefrag.DefragResult = $result
    }
    else {
        Write-Output "CRITICAL: Error $($result.ReturnValue) defragmenting drive $($diskToDefrag.DriveLetter)"
        Write-Output "Check error codes: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/vdswmi/defrag-method-in-class-win32-volume"
        Exit(2)
    }
    $global:output += "Disk $($diskToDefrag.DriveLetter) fragmentation is $($result.DefragAnalysis.FilePercentFragmentation)."
}

$ErrorActionPreference = "SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Continue"

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
        Write-Host "Analizing drive $($drive.DriveLetter)"
        $result = $drive.DefragAnalysis()
        if ($result.ReturnValue -eq 0) {
            Write-Host "Current fragmentation is $($result.DefragAnalysis.FilePercentFragmentation)"
            $drive | Add-Member -NotePropertyName 'DefragResult' -NotePropertyValue $result
            if (($defragPercentage -gt 0) -and ($result.DefragAnalysis.FilePercentFragmentation -gt $defragPercentage)) {
                Invoke-DiskDefragmentation -diskToDefrag $drive
            }
            else {
                $global:output += "Disk $($drive.DriveLetter) fragmentation is $($result.DefragAnalysis.FilePercentFragmentation)."
            }
        }
        else {
            Write-Output "CRITICAL: Error $($result.ReturnValue) checking status of drive $($drive.DriveLetter)"
            Write-Output "Check error codes: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/vdswmi/defraganalysis-method-in-class-win32-volume#return-value"
            Exit(2)
        }
    }
    
}
catch {
    Write-Output "CRITICAL: $($_.Exception.Message)"
    Exit(2)
}
Write-Host "-------"
Write-Host "Summary"
Write-Host "-------"
Write-Host $global:output
Stop-Transcript