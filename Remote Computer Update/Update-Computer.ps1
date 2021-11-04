#Requires -RunAsAdministrator
# It has to be executed from psexec: https://docs.microsoft.com/en-us/sysinternals/downloads/psexec
# Download it and copy to C:\Windows\System32
# psexec.exe -s \\ComputerName powershell.exe Set-ExecutionPolicy Unrestricted -Force
# psexec.exe -s \\ComputerName powershell.exe \\ES-CPD-BCK02\scripts\WindowsUpdate\Update-Computer.ps1
# It uses RZGet.exe to update computer software: https://github.com/rzander/ruckzuck/releases
## ------------------------------------------------------------------
# Variables
## ------------------------------------------------------------------
$logPath = "\\ES-CPD-BCK02\scripts\WindowsUpdate\Log" # Log file path.
$scheduleReboot = $true # if true, a reboot wil be scheduled if needed.
$rebootTime = 2 # Number of hours after finish update to reboot computer.
$rebootNow = $false # if true, reboots after finish update.
$showRebootMessage = $true # if true, shows a message to user
$rebootMessage = "Se va a reiniciar el equipo dentro de $($rebootTime) horas para terminar de instalar las actualizaciones de Windows. Por favor, cierra todo antes de esa hora o reinicia el equipo manualmente." # Reboot message
$RZGetPath = "\\ES-CPD-BCK02\scripts\WindowsUpdate\RZGet.exe" # Path to RZGet.exe
$RZGetArguments = 'update "7-Zip" "Google Chrome" "Notepad++" "Notepad++(x64)" "AdobeReader DC" "Putty" "WinSCP" "VLC" "JavaRuntime8" "JavaRuntime8x64" "KeePass" "Webex Meetings" "iTunes" "FileZilla" "Dell Command Update" "Dell Command Update W10"'
## ------------------------------------------------------------------
# Script
## ------------------------------------------------------------------
$timestamp = Get-Date -format yyyy-MM-dd-HH-mm
if (-not (Test-Path $logPath)) {
    if (-not (Test-Path C:\temp)) {
        mkdir C:\temp
    }
    $ReportFile = "C:\temp\$($timestamp)_$($env:COMPUTERNAME).txt";
} else {
    $ReportFile = "$logPath\$($timestamp)_$($env:COMPUTERNAME).txt";
}

Start-Transcript $ReportFile
$VerbosePreference='Continue'
$ErrorActionPreference = "SilentlyContinue"

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

        [ValidateSet("Full","StdOut","StdErr","ExitCode","None")]
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
        Title = ($MyInvocation.MyCommand).Name
        Command = $FilePath
        Arguments = $ArgumentList
        StdOut = $p.StandardOutput.ReadToEnd()
        StdErr = $p.StandardError.ReadToEnd()
        ExitCode = $p.ExitCode
        }
        $p.WaitForExit()

        if (-not([string]::IsNullOrEmpty($DisplayLevel))) {
            switch($DisplayLevel) {
                "Full" { return $result; break }
                "StdOut" { return $result.StdOut; break }
                "StdErr" { return $result.StdErr; break }
                "ExitCode" { return $result.ExitCode; break }
                }
            }
        } catch {
            Write-Host "An error has ocurred"
        }
}
## ------------------------------------------------------------------
# Script start
## ------------------------------------------------------------------

if (Get-PendingReboot) {
    Write-Host "Computer has a pending reboot"
    Stop-Transcript
    Exit
}
if ($Error) {
	$Error.Clear()
}
# http://blog.skadefro.dk/2014/05/windows-update-with-powershell.html
# https://ashbrook.io/2018-04-03-queue-up-or-pre-download-windows-updates-with-powershell/
$updateSearcher = New-Object -ComObject Microsoft.Update.Searcher
$updateSession = New-Object -ComObject Microsoft.Update.Session

Write-Host "Searching for applicable updates. Please wait..." -ForeGroundColor Yellow
$updateSearcherResult = $updateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
if ($updateSearcherResult.Updates.Count -eq 0) {
	Write-Host "There are no applicable updates for this computer."
} else {
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
    
    Write-Host "Downloading updates. Please wait..." -ForeGroundColor Yellow
    foreach($update in $updatesSelected) {
		Write-Host "$($count) of $($updateSearcherResult.Updates.Count) - Downloading Update: $($update.Title)"
        $count++
		if(!$update.IsDownloaded) {
			
			$updateDownloader.Updates.Add($update) | Out-Null

			$downloadResult = $updateDownloader.Download()

            if (($downloadResult.HResult -eq 0) -and ($downloadResult.ResultCode -eq 2)) {
                Write-Host "Download succeeded" -ForegroundColor "Green"
            } else {
                Write-Warning  "Download error: $($result.HResult)"
            }

            $updateDownloader.Updates.RemoveAt(0) | Out-Null
		}
	}
    $updateInstaller = $updateSession.CreateUpdateInstaller()
    if($updateInstaller.ForceQuiet -eq $false) { 
        $updateInstaller.ForceQuiet = $true 
    }
    $updateInstaller.Updates = New-Object -com Microsoft.Update.UpdateColl
    $count = 1;
    Write-Host "Installing updates. Please wait..." -ForeGroundColor Yellow
    foreach ($update in $updatesSelected) {
		Write-Host "$($count) of $($updateSearcherResult.Updates.Count) - Installing Update: $($update.Title)"
        $count++;

		$updateInstaller.Updates.Add($update) | Out-Null
		
		$installationResult = $updateInstaller.Install()

        if (($installationResult.ResultCode -eq 2) -and ($installationResult.HResult -eq 0)) {
			#$resultcode= @{0="Not Started"; 1="In Progress"; 2="Succeeded"; 3="Succeeded With Errors"; 4="Failed" ; 5="Aborted" }
			Write-Host "Installation succeeded" -ForegroundColor "Green"
		} elseif ($installationResult.ResultCode -eq 3) {
			Write-Warning " Installation succeeded with errors $($update.Title) ResultCode $($installationResult.ResultCode) HResult $($installationResult.HResult)"
		} elseif ($installationResult.ResultCode -eq 4) {
			Write-Warning "Installation failed $($update.Title) ResultCode $($installationResult.ResultCode) HResult $($installationResult.HResult)"
		} elseif ($installationResult.ResultCode -eq 5) {
			Write-Warning "Installation aborted $($update.Title) ResultCode $($installationResult.ResultCode) HResult $($installationResult.HResult)"
		}

        $updateInstaller.Updates.RemoveAt(0) | Out-Null
	}
}

if (Test-Path $RZGetPath) {
    Write-Host "Searching for app updates"
    Invoke-Process -FilePath $RZGetPath -ArgumentList "update --list --all" -DisplayLevel StdOut
    Write-Host "Running $($RZGetPath) $($RZGetArguments)"
    Invoke-Process -FilePath $RZGetPath -ArgumentList $RZGetArguments -DisplayLevel StdOut
} else {
    Write-Host "RZGet.exe not found"
}
if (Get-PendingReboot) {
    Write-Host "A reboot is needed to finish installing updates"
    if ($scheduleReboot) {
        if ($rebootNow) {
            $rebootTime = (Get-Date).AddSeconds(10)
        } else {
            $rebootTime = (Get-Date).AddHours($rebootTime)
        }
        $rebootTime = $rebootTime.toString('HH:mm:ss')
        
        SCHTASKS /delete /tn ScheduledReboot /f
        SCHTASKS /create /sc once /tn "ScheduledReboot" /st "$rebootTime" /tr "shutdown.exe /r /t 0 /f" /RU "NT AUTHORITY\SYSTEM" /RL HIGHEST
        
        SCHTASKS /delete /tn DeleteScheduledReboot /f
        SCHTASKS /create /sc onstart /tn "DeleteScheduledReboot" /tr "SCHTASKS /delete /tn ScheduledReboot /f" /RU "NT AUTHORITY\SYSTEM" /RL HIGHEST
        if ($showRebootMessage) {
            msg * /SERVER:localhost $rebootMessage
        }
    }
}
Stop-Transcript