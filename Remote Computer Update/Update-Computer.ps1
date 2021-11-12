<#PSScriptInfo

.VERSION 1.0

.GUID 5dc48422-ae3a-4b09-8510-5075b994a40f

.AUTHOR Juan Granados

.COPYRIGHT 2021 Juan Granados

.TAGS WSUS Windows Update Remote Software

.LICENSEURI https://raw.githubusercontent.com/juangranados/powershell-scripts/main/LICENSE

.PROJECTURI https://github.com/juangranados/powershell-scripts/tree/main/Remote%20Computer%20Update

.RELEASENOTES Initial release

#>

<#
.SYNOPSIS
    Installs Windows Updates on local or remote computer.
.DESCRIPTION
    Installs Windows Updates on local or remote computer.
    In order to run in remote computer, it has to be run from psexec: https://docs.microsoft.com/en-us/sysinternals/downloads/psexec
    See examples below.
    It uses RZGet.exe to update computer software: https://github.com/rzander/ruckzuck/releases
.PARAMETER logPath
    Log file path.
    Default "Documents"
    Example: "\\ES-CPD-BCK02\scripts\ComputerUpdate\Log"
.PARAMETER scheduleReboot
    Reboot wil be scheduled if needed.
    Default: false
.PARAMETER rebootHours
    Number of hours after finish update to reboot computer.
    Default: 2
.PARAMETER rebootNow
    Reboots after finish update.
    Default: false
.PARAMETER rebootMessage
    Shows a message to user.
    Default: none
.PARAMETER RZGetPath
    RZGet.exe path.
    Example: \\SRV-FS05\RZGet\rzget.exe
    If path not found RZGet will not be called unless 
    you set -downloadRZGet switch and it will be downloaded to this path.
    Default: none
.PARAMETER RZGetArguments
    RZGet.exe Arguments.
    Default: update --all
 .PARAMETER downloadRZGet
    Download RZGet.exe latest version to RZGetPath
.EXAMPLE
    Run remotely (PsExec needed: https://docs.microsoft.com/en-us/sysinternals/downloads/psexec)
    Two options:
    1 . Harcode parameters on script and run PsExec. 
        psexec.exe -s \\ComputerName powershell.exe -ExecutionPolicy Bypass -file \\ES-CPD-BCK02\scripts\WindowsUpdate\Update-Computer.ps1
    2. Use an script to run with parameters.
        LaunchRemote.cmd: run LaunchRemote.ps1 with PsExec.
        LaunchRemote.ps1: run Update-Computer.ps1 with arguments.
    Scripts are on this GitHub folder.
.EXAMPLE
    Run locally
    Update-Computer.ps1 -$logPath "\\ES-CPD-BCK02\scripts\WindowsUpdate\Log" -scheduleReboot -rebootHours 2 -rebootMessage "Computer will reboot in two hours. You can reboot now or it will reboot later" -RZGetPath "\\ES-CPD-BCK02\scripts\WindowsUpdate\RZGet.exe" $RZGetArguments 'update "7-Zip" "Google Chrome" "Notepad++(x64)" "AdobeReader DC"'
.LINK
    https://github.com/juangranados/powershell-scripts/tree/main/Remote%20Computer%20Update
.NOTES
    Author: Juan Granados 
#>

Param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$logPath = [Environment]::GetFolderPath("MyDocuments"),
    [Parameter()]
    [switch]$scheduleReboot,
    [Parameter(Mandatory = $false)] 
    [ValidateRange(0, 100)]
    [int]$rebootHours = 2,
    [Parameter()]
    [switch]$rebootNow,
    [Parameter(Mandatory = $false)] 
    [ValidateNotNullOrEmpty()]
    [string]$rebootMessage,
    [Parameter(Mandatory = $false)] 
    [ValidateNotNullOrEmpty()]
    [string]$RZGetPath,
    [Parameter(Mandatory = $false)] 
    [ValidateNotNullOrEmpty()]
    [string]$RZGetArguments = "update --all",
    [Parameter()]
    [switch]$downloadRZGet
)

#Requires -RunAsAdministrator

## ------------------------------------------------------------------
# function Get-PendingReboot
## ------------------------------------------------------------------
function Get-PendingReboot {

    #Check for Keys
    if ((Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") -eq $true) {
        return $true
    }

    if ((Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\PostRebootReporting") -eq $true) {
        return $true
    }

    if ((Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") -eq $true) {
        return $true
    }

    if ((Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending") -eq $true) {
        return $true
    }

    if ((Test-Path -Path "HKLM:\SOFTWARE\Microsoft\ServerManager\CurrentRebootAttempts") -eq $true) {
        return $true
    }

    #Check for Values
    if (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" | Select-Object -ExpandProperty "RebootInProgress" -ErrorAction SilentlyContinue) {
        return $true
    }

    if (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" | Select-Object -ExpandProperty "PackagesPending" -ErrorAction SilentlyContinue) {
        return $true
    }

    if (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce" | Select-Object -ExpandProperty "PendingFileRenameOperations" -ErrorAction SilentlyContinue) {
        return $true
    }

    if (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon" | Select-Object -ExpandProperty "PendingFileRenameOperations2" -ErrorAction SilentlyContinue) {
        return $true
    }

    if (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon" | Select-Object -ExpandProperty "DVDRebootSignal" -ErrorAction SilentlyContinue) {
        return $true
    }
    if (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon" | Select-Object -ExpandProperty "JoinDomain" -ErrorAction SilentlyContinue) {
        return $true
    }
    if (Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon" | Select-Object -ExpandProperty "AvoidSpnSet" -ErrorAction SilentlyContinue) {
        return $true
    }

    return $false
}
## ------------------------------------------------------------------
# function Invoke-Process
# https://stackoverflow.com/a/66700583
## ------------------------------------------------------------------
function Invoke-Process {
    param
    (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$FilePath,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$ArgumentList,

        [ValidateSet("Full", "StdOut", "StdErr", "ExitCode", "None")]
        [string]$DisplayLevel
    )

    $ErrorActionPreference = 'Stop'

    try {
        $pinfo = New-Object System.Diagnostics.ProcessStartInfo
        $pinfo.FileName = $FilePath
        $pinfo.RedirectStandardError = $true
        $pinfo.RedirectStandardOutput = $true
        $pinfo.UseShellExecute = $false
        $pinfo.WindowStyle = 'Hidden'
        $pinfo.CreateNoWindow = $true
        $pinfo.Arguments = $ArgumentList
        $p = New-Object System.Diagnostics.Process
        $p.StartInfo = $pinfo
        $p.Start() | Out-Null
        $result = [pscustomobject]@{
            Title     = ($MyInvocation.MyCommand).Name
            Command   = $FilePath
            Arguments = $ArgumentList
            StdOut    = $p.StandardOutput.ReadToEnd()
            StdErr    = $p.StandardError.ReadToEnd()
            ExitCode  = $p.ExitCode
        }
        $p.WaitForExit()

        if (-not([string]::IsNullOrEmpty($DisplayLevel))) {
            switch ($DisplayLevel) {
                "Full" { return $result; break }
                "StdOut" { return $result.StdOut; break }
                "StdErr" { return $result.StdErr; break }
                "ExitCode" { return $result.ExitCode; break }
            }
        }
    }
    catch {
        Write-Host "An error has ocurred"
    }
}
## ------------------------------------------------------------------
# Script start
## ------------------------------------------------------------------

if (-not (Test-Path $logPath)) {
    $logPath = [Environment]::GetFolderPath("MyDocuments")  
} 
$ReportFile = "$logPath\$(Get-Date -format yyyy-MM-dd-HH-mm)_$($env:COMPUTERNAME).txt";

Start-Transcript $ReportFile

$VerbosePreference = 'Continue'
$ErrorActionPreference = "SilentlyContinue"

if (Get-PendingReboot) {
    Write-Host "Computer has a pending reboot"
    Stop-Transcript
    Exit
}

$updateSearcher = New-Object -ComObject Microsoft.Update.Searcher
$updateSession = New-Object -ComObject Microsoft.Update.Session
Write-Host "----------------------------------------" -ForeGroundColor Cyan
Write-Host "Searching for applicable Windows updates" -ForeGroundColor Cyan
Write-Host "----------------------------------------" -ForeGroundColor Cyan
$updateSearcherResult = $updateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
if ($updateSearcherResult.Updates.Count -eq 0) {
    Write-Host "There are no applicable updates for this computer."
}
else {
    Write-Host "$($updateSearcherResult.Updates.Count) updates available"
    $updatesSelected = New-Object -com Microsoft.Update.UpdateColl
    foreach ($update in $updateSearcherResult.Updates) {
        if (!($update.EulaAccepted)) {
            $update.AcceptEula() | Out-Null
        }
        $updatesSelected.Add($update) | Out-Null
    }
   
    $updateDownloader = $updateSession.CreateUpdateDownloader()
    $updateDownloader.Updates = New-Object -com Microsoft.Update.UpdateColl
    $count = 1;
    Write-Host "---------------------------" -ForeGroundColor Cyan
    Write-Host "Downloading Windows updates" -ForeGroundColor Cyan
    Write-Host "---------------------------" -ForeGroundColor Cyan
    
    foreach ($update in $updatesSelected) {
        Write-Host "$($count) of $($updateSearcherResult.Updates.Count) - Downloading Update: $($update.Title)"
        $count++
        if (!$update.IsDownloaded) {
			
            $updateDownloader.Updates.Add($update) | Out-Null

            $downloadResult = $updateDownloader.Download()

            if (($downloadResult.HResult -eq 0) -and ($downloadResult.ResultCode -eq 2)) {
                Write-Host "Download succeeded" -ForegroundColor "Green"
            }
            else {
                Write-Warning  "Download error: $($downloadResult.HResult)"
            }

            $updateDownloader.Updates.RemoveAt(0) | Out-Null
        }
        else {
            Write-Host "Update already downloaded" -ForegroundColor "Green"
        }
    }
    $updateInstaller = $updateSession.CreateUpdateInstaller()
    if ($updateInstaller.ForceQuiet -eq $false) { 
        $updateInstaller.ForceQuiet = $true 
    }
    $updateInstaller.Updates = New-Object -com Microsoft.Update.UpdateColl
    $count = 1;
    Write-Host "--------------------------" -ForeGroundColor Cyan
    Write-Host "Installing Windows updates" -ForeGroundColor Cyan
    Write-Host "--------------------------" -ForeGroundColor Cyan
    foreach ($update in $updatesSelected) {
        if ($update.IsDownloaded) {
            Write-Host "$($count) of $($updateSearcherResult.Updates.Count) - Installing Update: $($update.Title)"
            $count++;

            $updateInstaller.Updates.Add($update) | Out-Null
            
            $installationResult = $updateInstaller.Install()

            if (($installationResult.ResultCode -eq 2) -and ($installationResult.HResult -eq 0)) {
                Write-Host "Installation succeeded" -ForegroundColor "Green"
            }
            elseif ($installationResult.ResultCode -eq 3) {
                Write-Warning " Installation succeeded with errors $($update.Title) ResultCode $($installationResult.ResultCode) HResult $($installationResult.HResult)"
            }
            elseif ($installationResult.ResultCode -eq 4) {
                Write-Warning "Installation failed $($update.Title) ResultCode $($installationResult.ResultCode) HResult $($installationResult.HResult)"
            }
            elseif ($installationResult.ResultCode -eq 5) {
                Write-Warning "Installation aborted $($update.Title) ResultCode $($installationResult.ResultCode) HResult $($installationResult.HResult)"
            }

            $updateInstaller.Updates.RemoveAt(0) | Out-Null
        }
        else {
            Write-Warning "$($count) of $($updateSearcherResult.Updates.Count) - Skipping Update: $($update.Title) because is not downloaded"
        }
        
    }
}
if ([string]::IsNullOrEmpty($RZGetPath)) {
    if ($downloadRZGet) {
        Write-Host "You have to set RZGetPath in order to download RZGet.exe" -ForegroundColor Yellow
    }
    else {
        write-Host "RZGetPath variable is empty. Skipping software update"
    }
}
else {
    if ($downloadRZGet) {
        try {
            Invoke-WebRequest "https://github.com/rzander/ruckzuck/releases/latest/download/RZGet.exe" -OutFile $RZGetPath -ErrorAction Stop
            Write-Host "Success downloading https://github.com/rzander/ruckzuck/releases/latest/download/RZGet.exe to $($RZGetPath)"
        }
        catch {
            Write-Host "Error downloading https://github.com/rzander/ruckzuck/releases/latest/download/RZGet.exe to $($RZGetPath): $($_.Exception.Message)"
            $_.Exception.Response 
        }
    }
    if (!(Test-Path $RZGetPath)) {
        Write-Host "RZGet.exe not found. Skipping software update"
    }
    else {
        Write-Host "------------------------------" -ForeGroundColor Cyan
        Write-Host "Searching for software updates" -ForeGroundColor Cyan
        Write-Host "------------------------------" -ForeGroundColor Cyan
        Write-Host "Available software updates"
        Invoke-Process -FilePath $RZGetPath -ArgumentList "update --list --all" -DisplayLevel StdOut
        Write-Host "Running $($RZGetPath) $($RZGetArguments)"
        Invoke-Process -FilePath $RZGetPath -ArgumentList $RZGetArguments -DisplayLevel StdOut
    }     
}

if (Get-PendingReboot) {
    Write-Host "A reboot is needed to finish installing updates" -ForegroundColor Yellow
    if ($scheduleReboot) {
        if ($rebootNow) {
            Write-Host "System will reboot now" -ForegroundColor Yellow
            Stop-Transcript
            Invoke-Process -FilePath "$env:SystemRoot\shutdown.exe" -ArgumentList "/r /t 10" -DisplayLevel Full
            Exit
        }
        $rebootTime = $(Get-Date).AddHours($rebootHours).toString('HH:mm')
        
        SCHTASKS /delete /tn ScheduledReboot /f 2> null
        SCHTASKS /create /sc once /tn "ScheduledReboot" /st "$rebootTime" /tr "shutdown.exe /r /t 0 /f" /RU "NT AUTHORITY\SYSTEM" /RL HIGHEST
        
        SCHTASKS /delete /tn DeleteScheduledReboot /f 2> null
        SCHTASKS /create /sc onstart /tn "DeleteScheduledReboot" /tr "SCHTASKS /delete /tn ScheduledReboot /f" /RU "NT AUTHORITY\SYSTEM" /RL HIGHEST
        if (-not ([string]::IsNullOrEmpty($rebootMessage))) {
            msg * /SERVER:localhost $rebootMessage
        }
        Write-Host "System will reboot at $($rebootTime)" -ForegroundColor Yellow
    }
}
Stop-Transcript
