<#PSScriptInfo

.VERSION 1.0

.GUID ede5f6c5-50d3-42a9-b958-daff4b31972d

.AUTHOR Juan Granados

.COPYRIGHT 2021 Juan Granados

.TAGS Install Printer Drivers Remote Remotely PrintNightmare PrinterExport

.LICENSEURI https://raw.githubusercontent.com/juangranados/powershell-scripts/main/LICENSE

.PROJECTURI https://github.com/juangranados/powershell-scripts/tree/main/Install%20Print%20Drivers%20Remotely

.RELEASENOTES
    Initial release
#>

<#
.SYNOPSIS
	Install printer drivers export file (*.printerExport) in a list of computers.
.DESCRIPTION
	Install printer drivers export file (*.printerExport) in a list of computers (plain computer list, OU or CSV file) using PSExec.
    To generate a printer export file with all print drivers, run PrintbrmUI.exe from the computer or server you want to export them.
    Be carefully because this tool exports all printers and ports too. I recommend install all drivers in a test computer without printers and export them using PrintbrmUI.exe.
    Alternatively, you can export all from print server, import in a test computer, delete printers and ports and export again to obtain a printerExport file with only drivers.
    If PSExec is not found on computer, script asks to the user for download it and extract in system folder.
    I recommend use PSExec latest version because v2.2 does not launch printbrm.exe properly.
.PARAMETER printerExportFile
	Path to the printerExport file
    To generate a printer export file with all print drivers, run PrintbrmUI.exe from the computer or server you want to export them.
    Be carefully because this tool exports all printers and ports too. I recommend install all drivers in a test computer without printers and export them using PrintbrmUI.exe.
    Alternatively, you can export all from print server, import in a test computer, delete printers and ports and export again to obtain a printerExport file with only drivers.
.PARAMETER ComputerList
    List of computers in install printer drivers. You can only use one source of target computers: ComputerList, OU or CSV.
    Example: SRV-RDSH-001,SRV-RDSH-002,SRV-RDSH-003 (Without quotation marks)
.PARAMETER OU
    OU containing computers in which install printer drivers.
    RSAT for AD module for PowerShell must be installed in order to query AD.
     - Install on Windows 10: Get-WindowsCapability -Online |? {$_.Name -like "*RSAT.ActiveDirectory*" -and $_.State -eq "NotPresent"} | Add-WindowsCapability -Online
     - Install on server: Install-WindowsFeature RSAT-AD-PowerShell
    Restart console after installation.
    If you run script from a Domain Controller, AD module for PowerShell is already enabled.
    You can only use one source of target computers: ComputerList, OU or CSV.
    Example: 'OU=Test,OU=Computers,DC=CONTOSO,DC=COM'
.PARAMETER CSV
    CSV file containing computers in which install printer drivers. You can only use one source of target computers: ComputerList, OU or CSV.
    Example: 'C:\Scripts\Computers.csv'
    CSV Format:
        Name
        Computer001
        Computer002
        Computer003
.PARAMETER LogPath
    Path where save log file.
    Default: My Documents
    Example: C:\Logs
.PARAMETER Credential
    Script will ask for an account to perform remote installation.
.EXAMPLE
    Install-PrinterExportRemoteComputers.ps1 -printerExportFile "\\MV-SRV-PR01\Drivers\print_drivers.printerExport" -OU "OU=RDS,OU=Datacenter,DC=CONTOSO,DC=COM"
.EXAMPLE    
    Install-PrinterExportRemoteComputers.ps1 -printerExportFile "\\MV-SRV-PR01\Drivers\print_drivers.printerExport" -ComputerList SRVRSH-001,SRVRSH-002,SRVRSH-003 -Credential -LogPath C:\Temp\Logs
.EXAMPLE
    Install-PrinterExportRemoteComputers.ps1 -printerExportFile "\\MV-SRV-PR01\Drivers\print_drivers.printerExport" -CSV "C:\scripts\computers.csv"
.LINK
    https://github.com/juangranados/powershell-scripts/tree/main/Install%20Print%20Drivers%20Remotely
.NOTES 
	Author: Juan Granados
#>

Param(
		[Parameter(Mandatory=$true,Position=0)] 
		[ValidateNotNullOrEmpty()]
		[string]$printerExportFile,
		[Parameter(Mandatory=$false,Position=1)] 
		[ValidateNotNullOrEmpty()]
		[string]$LocalPath="C:\temp",
        [Parameter(Mandatory=$false,Position=2)] 
		[ValidateNotNullOrEmpty()]
		[string[]]$ComputerList,
        [Parameter(Mandatory=$false,Position=3)] 
		[ValidateNotNullOrEmpty()]
		[string]$OU,
        [Parameter(Mandatory=$false,Position=4)] 
		[ValidateNotNullOrEmpty()]
		[string]$CSV,
        [Parameter(Mandatory=$false,Position=5)] 
		[ValidateNotNullOrEmpty()]
		[string]$LogPath=[Environment]::GetFolderPath("MyDocuments"),
        [Parameter(Position=6)] 
		[switch]$Credential
	)

#Requires -RunAsAdministrator

#Functions

Add-Type -AssemblyName System.IO.Compression.FileSystem
Import-Module BitsTransfer

function Unzip
{
    param([string]$zipfile, [string]$outpath)
    
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
}

$ErrorActionPreference = "Stop"

#Initialice log

$LogPath += "\InstallPrinterExportRemoteComputers_" + $(get-date -Format "yyyy-mm-dd_hh-mm-ss") + ".txt"
Start-Transcript $LogPath
Write-Host "Start remote installation on $(get-date -Format "yyyy-mm-dd hh:mm:ss")"

#Initial validations.

If (!(Test-Path $printerExportFile)){
    Write-Host "Error accessing $($printerExportFile). Script can not continue"
    Stop-Transcript
    Exit 1
}
if (-not($printerExportFile -Like "*.printerExport")) {
    Write-Host "File $($printerExportFile) does not has .printerExport extension. Script can not continue"
    Stop-Transcript
    Exit 1
}

if (!(Get-Command "psexec.exe" -ErrorAction SilentlyContinue)){ 
    Write-Host "Error. Microsoft Psexec not found on system. Download it from https://download.sysinternals.com/files/PSTools.zip and extract all in C:\Windows\System32" -ForegroundColor Yellow
    $Answer=Read-Host "Do you want to download and install PSTools (y/n)?"
    if (($Answer -eq "y") -or ($Answer -eq "Y")){
        Write-Host "Downloading PSTools"
        If (Test-Path "$($env:temp)\PSTools.zip"){
            Remove-Item "$($env:temp)\PSTools.zip" -Force
        }
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        (New-Object System.Net.WebClient).DownloadFile("https://download.sysinternals.com/files/PSTools.zip", "$($env:temp)\PSTools.zip")
        if (Test-Path "$($env:temp)\PSTools.zip"){
            Write-Host "Unzipping PSTools"
            If (Test-Path "$($env:temp)\PSTools"){
                Remove-Item "$($env:temp)\PSTools" -Force -Recurse
            }
            Unzip "$($env:temp)\PSTools.zip" "$($env:temp)\PSTools" 
            Copy-Item "$($env:temp)\PSTools\*.exe" "$($env:SystemRoot)\System32" -Force
            if (Test-Path "$($env:SystemRoot)\System32\psexec.exe"){
                Write-Host "PSTools installed" -ForegroundColor Green 
            }
            else{
                Write-Host "Error unzipping PSTools" -ForegroundColor Red
                Remove-Item "$($env:temp)\PSTools.zip" -Force
                Stop-Transcript
                Exit 1
            }
        }
        else{
            Write-Host "Error downloading PSTools" -ForegroundColor Red
            Stop-Transcript
            Exit 1
        }
    }
    else{
        Stop-Transcript
        Exit 1
    }
}

If ($OU){
    if (!(Get-Command "Get-ADComputer" -ErrorAction SilentlyContinue)){ 
        Write-Host "Error. Get-ADComputer not found on system. You have to install the PowerShell Active Directory module order to query Active Directory." -ForegroundColor Red
        Write-Host 'Windows 10: Get-WindowsCapability -Online |? {$_.Name -like "*RSAT.ActiveDirectory*" -and $_.State -eq "NotPresent"} | Add-WindowsCapability -Online'
        Write-Host "Server: Install-WindowsFeature RSAT-AD-PowerShell"
        Write-Host "Restart console after installation"
        Stop-Transcript
        Exit 1
    }
    try{
        $ComputerList = Get-ADComputer -Filter * -SearchBase "$OU" | Select-Object -Expand name
    }catch{
        Write-Host "Error querying AD: $($_.Exception.Message)" -ForegroundColor Red
        Stop-Transcript
        Exit 1
    }
}
ElseIf ($CSV){
    try{
        $ComputerList = Get-Content $CSV | where {$_ -notmatch 'Name'} | Foreach-Object {$_ -replace '"', ''}
    }catch{
        Write-Host "Error getting CSV content: $($_.Exception.Message)" -ForegroundColor Red
        Stop-Transcript
        Exit 1
    }
}
ElseIf(!$ComputerList){
    Write-Host "You have to set a list of computers, OU or CSV." -ForegroundColor Red
    Stop-Transcript
    Exit 1
}
If ($Credential){
    $Cred = Get-Credential
}
$usingCredential = $false
If(!$Cred -or !$Credential){
    Write-Host "No credential specified. Using logon account"
}
Else{
    $usingCredential = $true;
    Write-Host "Using user $($Cred.UserName)"
    $UserName = $Cred.UserName
    $Password = $Cred.GetNetworkCredential().Password
}
ForEach ($Computer in $ComputerList) {
    Write-Host "Processing computer $Computer"
    $Destination = "\\$Computer\C$\Temp\printer_drivers.printerExport"
    try {
    if (-not(Test-Path "\\$Computer\C$\Temp\")) {
        mkdir "\\$Computer\C$\Temp\"
    }
    Start-BitsTransfer -Source $printerExportFile -Destination $Destination -Description "Copy $driverFile to $Computer" -DisplayName "Copying"
    } catch {
        Write-Host "Error copying file: $($_.Exception.Message)"
        continue
    }
    Write-Host "Launching installation using PSExec in $Computer. This may take a while, please be patient..."
    try {
        if ($usingCredential) {
            psexec.exe /accepteula -h -i "\\$Computer" -u $UserName -p $Password C:\Windows\System32\spool\tools\Printbrm.exe -F C:\Temp\printer_drivers.printerExport -R
        }
        else {
            psexec /accepteula -h "\\$Computer" C:\Windows\System32\spool\tools\Printbrm.exe -F C:\Temp\printer_drivers.printerExport -R
        }
    } catch {
        Write-Host "PSExec return an error. Check console output above"
    }

    try {
        Write-Host "Removing remote file"
        Remove-Item $Destination -Force;
    } catch {
        Write-Host "Error removing remote file: $($_.Exception.Message)" -ForegroundColor Red
    }
}     
Stop-Transcript