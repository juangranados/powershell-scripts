<#
.SYNOPSIS
    Performs basic computer mainteinance.
.DESCRIPTION
    Performs basic computer mainteinance.
    Runs SFC, DISM, WMI heath check, Antivirus scan, drive drefrag, Windows Update.
.PARAMETER all
    Runs all.
    Default: false
.PARAMETER sfc
    Runs SFC /scannow.
    Default: false
.PARAMETER dism
    Runs DISM /Online /Cleanup-Image /RestoreHealth
    Default: false
.PARAMETER wmi
    Runs Winmgmt /salvagerepository
    Default: false
.PARAMETER defrag
    Defrag drives if required (Fragmentation > 10%)
    Default: false
.PARAMETER update
    Install Updates (except drivers)
    Default: false
.PARAMETER defender
    Runs Windows Defender Update and Quick Scan
    Default: false
.PARAMETER adaware
    Runs adaware update and quick/boot Scan
    Default: false Kaspersky Virus Removal Tool
.PARAMETER kas
    Runs Kaspersky Virus Removal Tool
    Default: false
PARAMETER LogPath
    Path where save log file.
    Default: My Documents
.EXAMPLE
    Runs all commands.
    Win-Mnt.ps1 -all
.EXAMPLE
    Runs sfc and defrag with custom log.
    Win-Mnt.ps1 -sfc -defrag -logPath "\\INFSRV001\Scripts$\Mainteinance\Logs"
.EXAMPLE
    Runs sfc, dism and adaware scan.
    Win-Mnt.ps1 -sfc -dism -adaware
.NOTES 
    Author: Juan Granados
#>

Param(
    [Parameter()] 
    [switch]$all,
    [Parameter()] 
    [switch]$sfc,
    [Parameter()] 
    [switch]$dism,
    [Parameter()] 
    [switch]$wmi,
    [Parameter()] 
    [switch]$defender,
    [Parameter()] 
    [switch]$adaware,
    [Parameter()] 
    [switch]$kas,
    [Parameter()] 
    [switch]$defrag,
    [Parameter()] 
    [switch]$update,
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
	[string]$logPath=[Environment]::GetFolderPath("MyDocuments")
)
#Requires -RunAsAdministrator
$timestamp = Get-Date -format yyyy-MM-dd-HH-mm
$log = "$logPath\$($timestamp)_$($env:COMPUTERNAME).txt"
$host.UI.RawUI.WindowTitle = "Performing mainteinance on computer $env:COMPUTERNAME"
Start-Transcript $log
if ($sfc -or $all) {
    Write-Host "--------------------"
    Write-Host "Running sfc /scannow"
    Write-Host "--------------------"
    $oldEncoding = [console]::OutputEncoding
    [console]::OutputEncoding = [Text.Encoding]::Unicode
    sfc /scannow
    [console]::OutputEncoding = $oldEncoding
    Write-Host "SFC returns exit code: $LASTEXITCODE"
}
if ($dism -or $all) {
    Write-Host "--------------------------------------------------"
    Write-Host "Running DISM /Online /Cleanup-Image /RestoreHealth"
    Write-Host "--------------------------------------------------"
    DISM /Online /Cleanup-Image /RestoreHealth
    Write-Host "DISM returns exit code: $LASTEXITCODE"
}
if ($wmi -or $all) {
    Write-Host "----------------------------------"
    Write-Host "Running Winmgmt /salvagerepository"
    Write-Host "----------------------------------"
    Winmgmt /salvagerepository

}
if ($defender -or $all) {
    Write-Host "------------------------------------------------"
    Write-Host "Updating and running Windows Defender Quick Scan"
    Write-Host "------------------------------------------------"
    try {
        Get-MpComputerStatus
        $startTime = Get-Date
        Update-MpSignature -UpdateSource MicrosoftUpdateServer -ErrorAction Stop
        Start-MpScan -ErrorAction Stop
        sleep 5
        Get-WinEvent -LogName 'Microsoft-Windows-Windows Defender/Operational' | Where-Object {$_.TimeCreated -ge $startTime} | Select-Object -ExpandProperty Message
    } catch {
        Write-Output "An error has ocurred: $($_.Exception.Message)"
    } 
}
if ($adaware -or $all) {
    Write-Host "-----------------------------------------"
    Write-Host "Downloading, updating and running adaware"
    Write-Host "-----------------------------------------"
    Write-Host "Downloading adawareCommandLineScanner"
    $(New-Object System.Net.WebClient).DownloadFile("https://www.adaware.com/sites/default/files/av/command-line-scanner/64/AdAwareCommandLineScanner.zip#inline-64bit","$env:LOCALAPPDATA\Temp\AdAwareCommandLineScanner.zip")
    if (Test-Path "$env:LOCALAPPDATA\Temp\AdAwareCommandLineScanner.zip") {
        Write-Host "Extracting AdAwareCommandLineScanner.zip"
        Expand-Archive -LiteralPath "$env:LOCALAPPDATA\Temp\AdAwareCommandLineScanner.zip" "$env:LOCALAPPDATA\Temp\AdAwareCommandLineScanner" -Force
        Remove-Item "$env:LOCALAPPDATA\Temp\AdAwareCommandLineScanner.zip" -Force
        if (Test-Path "$env:LOCALAPPDATA\Temp\AdAwareCommandLineScanner\AdAwareCommandLineScanner.exe") {
            Write-Host "Updating definitions"
            Start-Process "$env:LOCALAPPDATA\Temp\AdAwareCommandLineScanner\AdAwareCommandLineScanner.exe" "--updatedefs" -Wait
            Write-Host "Performing quick scan"
            Start-Process "$env:LOCALAPPDATA\Temp\AdAwareCommandLineScanner\AdAwareCommandLineScanner.exe" "--quick --disinfect --scan-result $($logPath)\$($timestamp)_$($env:COMPUTERNAME)_adware_quick_scan.xml" -Wait
            $xml = [xml](Get-Content "$($logPath)\$($timestamp)_$($env:COMPUTERNAME)_adware_quick_scan.xml")
            if ($xml.Summary.InfectedObjects.ChildNodes.Count -eq 0) {
                Write-Host "No infected items found"
            } else {
                Write-Host "$($xml.Summary.InfectedObjects.ChildNodes.Count) infected items found!"
                $xml.Summary.InfectedObjects.ChildNodes
            }
            Write-Host "Performing boot scan"
            Start-Process "$env:LOCALAPPDATA\Temp\AdAwareCommandLineScanner\AdAwareCommandLineScanner.exe" "--boot --disinfect --scan-result $($logPath)\$($timestamp)_$($env:COMPUTERNAME)_adware_boot_scan.xml" -Wait
            $xml = [xml](Get-Content "$($logPath)\$($timestamp)_$($env:COMPUTERNAME)_adware_boot_scan.xml")
            if ($myxml.Summary.InfectedObjects.ChildNodes.Count -eq 0) {
                Write-Host "No infected items found"
            } else {
                Write-Host "$($myxml.Summary.InfectedObjects.ChildNodes.Count) infected items found!"
                $myxml.Summary.InfectedObjects.ChildNodes
            }
        } else {
            Write-Host "Error extracting adawareCommandLineScanner"
        }
    } else {
        Write-Host "Error downloading adawareCommandLineScanner"
    }
}
if ($kas -or $all) {
    Write-Host "----------------------------"
    Write-Host "Downloading and running KVRT"
    Write-Host "----------------------------"
    Write-Host "Downloading KVRT"
    $(New-Object System.Net.WebClient).DownloadFile("https://devbuilds.s.kaspersky-labs.com/devbuilds/KVRT/latest/full/KVRT.exe","$env:LOCALAPPDATA\Temp\KVRT.exe")
    if (Test-Path "$env:LOCALAPPDATA\Temp\KVRT.exe") {
        Write-Host "Running KVRT"
        $kvrtProcess = Start-Process "$env:LOCALAPPDATA\Temp\KVRT.exe" "-accepteula -adinsilent -silent -processlevel 2 -fupdate" -Wait -NoNewWindow -RedirectStandardOutput "$($logPath)\$($timestamp)_$($env:COMPUTERNAME)_kvrt_scan.txt"
        Get-Content -Path "$($logPath)\$($timestamp)_$($env:COMPUTERNAME)_kvrt_scan.txt"
    } else {
        Write-Host "Error downloading KVRT"
    }
}
if ($defrag -or $all) {
    Write-Host "----------------------"
    Write-Host "Checking defrag status"
    Write-Host "----------------------"
    $drives = get-wmiobject win32_volume | ? { $_.DriveType -eq 3 -and $_.DriveLetter}
    if (-not ($drives)){
        Write-Output "No se han encontrado discos con el comando get-wmiobject win32_volume"
    } else {
        foreach ($drive in $drives) {
            $result = $drive.DefragAnalysis()
            if ($result.ReturnValue -eq 0) {
                if ($result.DefragRecommended -eq "True") {
                    Write-Host "Drive $($drive.DriveLetter) need defragmentation. Defragmenting"
                    $result = $drive.Defrag($true)
                    if ($result.ReturnValue -eq 0) {
                        Write-Host "Defragmentation performed on drive $($drive.DriveLetter). $($result.DefragAnalysis.FreeSpacePercent)% free"
                    } else {
                        Write-Host "An error with code $($result.ReturnValue) has ocurred while defragmenting drive $($drive.DriveLetter)"
                    }
                } else {
                    Write-Host "Drive $($drive.DriveLetter) does not need defragmentation"
                }
            } else {
                Write-Host "An error with code $($result.ReturnValue) has ocurred while checking status of drive $($drive.DriveLetter)"
            }
        }
    }
}
if ($update -or $all) {
    Write-Host "------------------"
    Write-Host "Installing Updates"
    Write-Host "------------------"
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    Install-PackageProvider -Name NuGet -Force -Confirm:$false
    Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
    Install-Module -Name PSWindowsUpdate -Force -Confirm:$false
    Install-WindowsUpdate -MicrosoftUpdate -NotCategory "Drivers" -AcceptAll -IgnoreUserInput -IgnoreReboot
}

Stop-Transcript