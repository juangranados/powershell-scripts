<#PSScriptInfo

.VERSION 1.0.1

.GUID 9d73a2b5-1329-42c0-b0ec-328198e3392d

.AUTHOR Juan Granados

.COPYRIGHT 2021 Juan Granados

.TAGS Install Software Remote Remotely

.LICENSEURI https://raw.githubusercontent.com/juangranados/powershell-scripts/main/LICENSE

.PROJECTURI https://github.com/juangranados/powershell-scripts/tree/main/Install%20Software%20Remotely

.RELEASENOTES
    Update examples
#>

<#
.SYNOPSIS
	Install software remotely in a group of computers and retry the installation in case of error.
.DESCRIPTION
	Install software remotely in a group of computers and retry the installation in case of error.
    It uses PowerShell to perform the installation. Target computer must allow Windows PowerShell Remoting.
    Script can try to enable Windows PowerShell Remoting using Microsoft Sysinternals Psexec with the paramenter -EnablePSRemoting. 
    If PSExec is not found on computer, script asks to the user for download it and extract in system folder.
.PARAMETER AppPath
	Path to the application executable, It can be a network or local path because entire folder will be copied to remote computer before installing and deleted after installation. 
    Example: 'C:\Software\TeamViewer\TeamvieverHost.msi' (Folder TeamViewer will be copied to remote computer before run ejecutable)
.PARAMETER AppArgs
    Application arguments to perform silent installation.
    Example: '/S /R settings.reg'
.PARAMETER LocalPath
    Local path of the remote computer where copy application directory.
    Default: 'C:\temp'
.PARAMETER Retries
    Number of times to retry failed installations.
    Default: 5
.PARAMETER TimeBetweenRetries
    Seconds to wait before retrying failed installations.
    Default: 60
.PARAMETER ComputerList
    List of computers in install software. You can only use one source of target computers: ComputerList, OU or CSV.
    Example: Computer001,Computer002,Computer003 (Without quotation marks)
.PARAMETER OU
    OU containing computers in which install software.
    RSAT for AD module for PowerShell must be installed in order to query AD.
    If you run script from a Domain Controller, AD module for PowerShell is already enabled.
    You can only use one source of target computers: ComputerList, OU or CSV.
    Example: 'OU=Test,OU=Computers,DC=CONTOSO,DC=COM'
.PARAMETER CSV
    CSV file containing computers in which install software. You can only use one source of target computers: ComputerList, OU or CSV.
    Example: 'C:\Scripts\Computers.csv'
    CSV Format:
                Name
                Computer001
                Computer002
                Computer003
.PARAMETER LogPath
    Path where save log file.
    Default: My Documents
.PARAMETER Credential
    Script will ask for an account to perform remote installation.
.PARAMETER EnablePSRemoting
    Try to enable PSRemoting on failed computers using Psexec. Psexec has to be on system path.
    If PSExec is not found. Script ask to download automatically PSTools and copy them to C:\Windows\System32.
.PARAMETER AppName
    App name as shown in registry to check if app is installed on remote computer and not reinstall it.
    You can check app name on a computer with it installed looking at:
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\'
    'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\'
    Example: 'TightVNC'
    Default: None
.PARAMETER AppVersion
    App name as shown in registry to check if app is installed on remote computer and not reinstall it.
    If not specified and AppName has a value, version will be ignored.
    You can check app version on a computer with it installed looking at:
    'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\'
    'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\'
    Example: '2.0.8.1'
    Default: all
.PARAMETER WMIQuery
    WMI Query to execute in remote computers. Software will be installed if query returns values.
    Example: 'select * from Win32_Processor where DeviceID="CPU0" and AddressWidth="64"' (64 bit computers)
    Example: 'select * from Win32_Processor where DeviceID="CPU0" and AddressWidth="32"' (32 bit computers)
    Default: None
.EXAMPLE
    TightVNC -> Install-SoftwareRemotely.ps1 -AppPath 'C:\Scripts\TightVNC\tightvnc-2.8.8-gpl-setup-64bit.msi' -AppArgs '/quiet /norestart ADDLOCAL="Server" SERVER_REGISTER_AS_SERVICE=1 SERVER_ADD_FIREWALL_EXCEPTION=1 SERVER_ALLOW_SAS=1 SET_USEVNCAUTHENTICATION=1 VALUE_OF_USEVNCAUTHENTICATION=1 SET_PASSWORD=1 VALUE_OF_PASSWORD=Password.01 SET_USECONTROLAUTHENTICATION=1 VALUE_OF_USECONTROLAUTHENTICATION=1 SET_CONTROLPASSWORD=1 VALUE_OF_CONTROLPASSWORD=3digits.01' -OU 'OU=Central,OU=Computers,DC=Contoso,DC=local' -Retries 2 -AppName 'TightVNC' -AppVersion '2.8.8.0' -EnablePSRemoting -WMIQuery 'select * from Win32_Processor where DeviceID="CPU0" and AddressWidth="64"'
.EXAMPLE
    TightVNC Mirage Driver -> Install-SoftwareRemotely.ps1 -AppPath 'C:\Scripts\TightVNC\dfmirage-setup-2.0.301.exe' -AppArgs '/verysilent /norestart' -OU 'OU=Central,OU=Computers,OU=MyBusiness,DC=Contoso,DC=local' -Retries 2 -AppName 'DemoForge Mirage Driver for TightVNC 2.0' -AppVersion '2.0' -EnablePSRemoting -WMIQuery 'select * from Win32_Processor where DeviceID="CPU0" and AddressWidth="64"'
.EXAMPLE
	Install-SoftwareRemotely.ps1 -AppPath "C:\Temp\Software\Miranda\miranda-im-v0.10.75-unicode.exe" -AppArgs "/S" -ComputerList Computer001,Computer002,Computer003 -AppName "Miranda IM 0.10.75" -AppVersion "0.10.75"
.EXAMPLE
	Install-SoftwareRemotely.ps1 -AppPath "C:\Temp\Software\Miranda\miranda-im-v0.10.75-unicode.exe" -AppArgs "/S" -CSV "C:\Computers.csv" -Credential -EnablePSRemoting
.EXAMPLE
	Install-SoftwareRemotely.ps1 -AppPath "\\Server01\Software\Miranda\miranda-im-v0.10.75-unicode.exe" -AppArgs "/S" -OU "OU=Test,OU=Computers,DC=CONTOSO,DC=COM"
.NOTES 
	Author: Juan Granados 
	Date:   July 2021
#>

Param(
		[Parameter(Mandatory=$true,Position=0)] 
		[ValidateNotNullOrEmpty()]
		[string]$AppPath,
		[Parameter(Mandatory=$false,Position=1)] 
		[ValidateNotNullOrEmpty()]
		[string]$AppArgs="None",
		[Parameter(Mandatory=$false,Position=2)] 
		[ValidateNotNullOrEmpty()]
		[string]$LocalPath="C:\temp",
        [Parameter(Mandatory=$false,Position=3)] 
		[ValidateNotNullOrEmpty()]
		[int]$Retries=5,
        [Parameter(Mandatory=$false,Position=4)] 
		[ValidateNotNullOrEmpty()]
		[int]$TimeBetweenRetries=60,
        [Parameter(Mandatory=$false,Position=5)] 
		[ValidateNotNullOrEmpty()]
		[string[]]$ComputerList,
        [Parameter(Mandatory=$false,Position=6)] 
		[ValidateNotNullOrEmpty()]
		[string]$OU,
        [Parameter(Mandatory=$false,Position=7)] 
		[ValidateNotNullOrEmpty()]
		[string]$CSV,
        [Parameter(Mandatory=$false,Position=7)] 
		[ValidateNotNullOrEmpty()]
		[string]$LogPath=[Environment]::GetFolderPath("MyDocuments"),
        [Parameter(Position=9)] 
		[switch]$EnablePSRemoting,
        [Parameter(Position=10)] 
		[switch]$Credential,
        [Parameter(Mandatory=$false,Position=11)] 
		[ValidateNotNullOrEmpty()]
		[string]$AppName="None",
        [Parameter(Mandatory=$false,Position=12)] 
		[ValidateNotNullOrEmpty()]
		[string]$AppVersion="all",
        [Parameter(Mandatory=$false,Position=13)] 
		[ValidateNotNullOrEmpty()]
		[string]$WMIQuery="None"
	)

#Requires -RunAsAdministrator

#Functions

Add-Type -AssemblyName System.IO.Compression.FileSystem

function Unzip
{
    param([string]$zipfile, [string]$outpath)
    
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
}

Function Copy-WithProgress
{
    Param([string]$Source,[string]$Destination)

    $Source=$Source.tolower()
    $Filelist=Get-Childitem $Source –Recurse
    $Total=$Filelist.count
    $Position=0
    If(!(Test-Path $Destination)){
        New-Item $Destination -Type Directory | Out-Null
    }
    foreach ($File in $Filelist){
        $Filename=$File.Fullname.tolower().replace($Source,'')
        $DestinationFile=($Destination+$Filename)
        try{
            Copy-Item $File.FullName -Destination $DestinationFile -Force
        }catch{throw $_.Exception}
        $Position++
        Write-Progress -Activity "Copying data from $source to $Destination" -Status "Copying File $Filename" -PercentComplete (($Position/$Total)*100)
    }
}

Function Set-Message([string]$Text,[string]$ForegroundColor="White",[int]$Append=$True){

    if ($Append){
        $Text | Out-File $LogPath -Append
    }
    else {
        $Text | Out-File $LogPath
    }
    Write-Host $Text -ForegroundColor $ForegroundColor
}

function Get-InstalledApps
{
    if ([IntPtr]::Size -eq 4) {
        $regpath = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    }
    else {
        $regpath = @(
            'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
            'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
        )
    }
    Get-ItemProperty $regpath | .{process{if($_.DisplayName -and $_.UninstallString) { $_ } }} | 
        Select DisplayName, Publisher, InstallDate, DisplayVersion, UninstallString | 
        Sort DisplayName
    #$result = Get-InstalledApps | where {$_.DisplayName -like $appToMatch}

}


Function CheckSoftwareInstalled([string]$Computer){
    If ($Cred){
        try{
            Return Invoke-Command -computername $Computer -ScriptBlock { 
                $AppName = $args[0]
                $AppVersion = $args[1]              
                if ([IntPtr]::Size -eq 4) {
                    $regpath = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
                }
                else {
                    $regpath = @(
                        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
                        'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
                    )
                }
                $InstalledApps = Get-ItemProperty $regpath | .{process{if($_.DisplayName -and $_.UninstallString) { $_ } }} | 
                    Select DisplayName, Publisher, InstallDate, DisplayVersion, UninstallString | 
                    Sort DisplayName
                If ($AppVersion -ne "all"){
                    Return $InstalledApps | where {$_.DisplayName -eq $AppName -and $_.DisplayVersion -eq  $AppVersion}
                }
                Else{
                    Return $InstalledApps | where {$_.DisplayName -eq $AppName}
                }
            } -ArgumentList $AppName, $AppVersion -Credential $Cred
        }catch{throw $_.Exception}
    }
    else{
        try{
            Return Invoke-Command -computername $Computer -ScriptBlock {
                $AppName = $args[0]
                $AppVersion = $args[1]              
                if ([IntPtr]::Size -eq 4) {
                    $regpath = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
                }
                else {
                    $regpath = @(
                        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
                        'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
                    )
                }
                $InstalledApps = Get-ItemProperty $regpath | .{process{if($_.DisplayName -and $_.UninstallString) { $_ } }} | 
                    Select DisplayName, Publisher, InstallDate, DisplayVersion, UninstallString | 
                    Sort DisplayName
                If ($AppVersion -ne "all"){
                    Return $InstalledApps | where {$_.DisplayName -eq $AppName -and $_.DisplayVersion -eq  $AppVersion}
                }
                Else{
                    Return $InstalledApps | where {$_.DisplayName -eq $AppName}
                }
            } -ArgumentList $AppName, $AppVersion
        }catch{throw $_.Exception}
    }
}

Function CheckWMIQuery([string]$Computer){
    If ($Cred){
        try{
            Return Invoke-Command -computername $Computer -ScriptBlock { 
                $WMIQuery = $args[0]
                Write-Host "Executing $($WMIQuery)"
                Return gwmi -Query $WMIQuery
            } -ArgumentList $WMIQuery -Credential $Cred
        }catch{throw $_.Exception}
    }
    else{
        try{
            Return Invoke-Command -computername $Computer -ScriptBlock { 
                $WMIQuery = $args[0]
                Write-Host "Executing $($WMIQuery)"
                Return gwmi -Query $WMIQuery
            } -ArgumentList $WMIQuery
        }catch{throw $_.Exception}   
    }
}

Function InstallRemoteSoftware([string]$Computer){
    If ($Cred){
        try{
            Return Invoke-Command -computername $Computer -ScriptBlock {
                $Application = $args[0]
                $AppArgs = $args[1]
                $ApplicationName = $Application.Substring($Application.LastIndexOf('\')+1)
                $ApplicationFolderPath = $Application.Substring(0,$Application.LastIndexOf('\'))
                $ApplicationExt = $Application.Substring($Application.LastIndexOf('.')+1)
                Write-Host "Installing $($ApplicationName) on $($env:COMPUTERNAME)"
                If($ApplicationExt -eq "msi"){
                    If ($AppArgs -ne "None"){
                         Write-Host "Installing as MSI: msiexec /i $($Application) $($AppArgs)"
                        $p = Start-Process "msiexec" -ArgumentList "/i $($Application) $($AppArgs)" -Wait -Passthru
                    }
                    else{
                        Write-Host "Installing as MSI: msiexec /i $($Application)"
                        $p = Start-Process "msiexec" -ArgumentList "/i $($Application) /quiet /norestart" -Wait -Passthru
                    }
                }
                ElseIf ($AppArgs -ne "None"){
                    Write-Host "Executing $Application $AppArgs"
                    $p = Start-Process $Application -ArgumentList $AppArgs -Wait -Passthru
                }
                Else{
                    Write-Host "Executing $Application"
                    $p = Start-Process $Application -Wait -Passthru
                }
                $p.WaitForExit()
                if ($p.ExitCode -ne 0) {
                    Write-Host "Failed installing with error code $($p.ExitCode)" -ForegroundColor Red
                    $Return = $($env:COMPUTERNAME)
                }
                else{
                    $Return = 0
                }
                Write-Host "Deleting $($ApplicationFolderPath)"
                Remove-Item $($ApplicationFolderPath) -Force -Recurse
                Return $Return
            } -ArgumentList "$($LocalPath)\$($ApplicationFolderName)\$($ApplicationName)", $AppArgs -Credential $Cred
        }catch{throw $_.Exception}
    }
    else{
        try{
            Return Invoke-Command -computername $Computer -ScriptBlock {
                $Application = $args[0]
                $AppArgs = $args[1]
                $ApplicationName = $Application.Substring($Application.LastIndexOf('\')+1)
                $ApplicationFolderPath = $Application.Substring(0,$Application.LastIndexOf('\'))
                $ApplicationExt = $Application.Substring($Application.LastIndexOf('.')+1)
                Write-Host "Installing $($ApplicationName) on $($env:COMPUTERNAME)"
                If($ApplicationExt -eq "msi"){
                    If ($AppArgs -ne "None"){
                         Write-Host "Installing as MSI: msiexec /i $($Application) $($AppArgs)"
                        $p = Start-Process "msiexec" -ArgumentList "/i $($Application) $($AppArgs)" -Wait -Passthru
                    }
                    else{
                        Write-Host "Installing as MSI: msiexec /i $($Application)"
                        $p = Start-Process "msiexec" -ArgumentList "/i $($Application) /quiet /norestart" -Wait -Passthru
                    }
                }
                ElseIf ($AppArgs -ne "None"){
                    Write-Host "Executing $Application $AppArgs"
                    $p = Start-Process $Application -ArgumentList $AppArgs -Wait -Passthru
                }
                Else{
                    Write-Host "Executing $Application"
                    $p = Start-Process $Application -Wait -Passthru
                }
                $p.WaitForExit()
                if ($p.ExitCode -ne 0) {
                    Write-Host "Failed installing with error code $($p.ExitCode)" -ForegroundColor Red
                    $Return = $($env:COMPUTERNAME)
                }
                else{
                    $Return = 0
                }
                Write-Host "Deleting $($ApplicationFolderPath)"
                Remove-Item $($ApplicationFolderPath) -Force -Recurse
                Return $Return
            } -ArgumentList "$($LocalPath)\$($ApplicationFolderName)\$($ApplicationName)", $AppArgs
        }catch{throw $_.Exception}
    }
}

Function CheckPSRemoting([string]$Computer){
    If ($EnablePSRemoting){
        Set-Message "Enabling PSRemoting on computer: psexec.exe /accepteula -h -d \\$($Computer) -s powershell Enable-PSRemoting"
        try{
            psexec.exe /accepteula -h -d "\\$($Computer)" -s powershell Enable-PSRemoting -Force 2>&1 | Out-Null
        }catch{
            Set-Message "PSExec running on background. Continue with next computer."
        }
    }
    Else{
        Set-Message "You can try to enable PowerShell Remoting on computer using parameter -EnablePSRemoting" -ForegroundColor DarkYellow
    }
}

$ErrorActionPreference = "Stop"

#Initialice log

$LogPath += "\InstallSoftwareRemotely_" + $(get-date -Format "yyyy-mm-dd_hh-mm-ss") + ".txt"
Set-Message "Start remote installation on $(get-date -Format "yyyy-mm-dd hh:mm:ss")" -Append $False

#Initial validations.

If (!(Test-Path $AppPath)){
    Set-Message "Error accessing $($AppPath). The script can not continue"
    Exit 1
}
If ($EnablePSRemoting){
    if (!(Get-Command "psexec.exe" -ErrorAction SilentlyContinue)){ 
        Set-Message "Error. Microsoft Psexec not found on system. Download it from https://download.sysinternals.com/files/PSTools.zip and extract all in C:\Windows\System32" -ForegroundColor Yellow
        $Answer=Read-Host "Do you want to download and install PSTools (y/n)?"
        if (($Answer -eq "y") -or ($Answer -eq "Y")){
            Set-Message "Downloading PSTools"
            If (Test-Path "$($env:temp)\PSTools.zip"){
                Remove-Item "$($env:temp)\PSTools.zip" -Force
            }
            (New-Object System.Net.WebClient).DownloadFile("https://download.sysinternals.com/files/PSTools.zip", "$($env:temp)\PSTools.zip")
            if (Test-Path "$($env:temp)\PSTools.zip"){
                Set-Message "Unzipping PSTools"
                If (Test-Path "$($env:temp)\PSTools"){
                    Remove-Item "$($env:temp)\PSTools" -Force -Recurse
                }
                Unzip "$($env:temp)\PSTools.zip" "$($env:temp)\PSTools" 
                Copy-Item "$($env:temp)\PSTools\*.exe" "$($env:SystemRoot)\System32" -Force
                if (Test-Path "$($env:SystemRoot)\System32\psexec.exe"){
                    Set-Message "PSTools installed" -ForegroundColor Green 
                }
                else{
                    Set-Message "Error unzipping PSTools" -ForegroundColor Red
                    Remove-Item "$($env:temp)\PSTools.zip" -Force
                    Exit 1
                }
            }
            else{
                Set-Message "Error downloading PSTools" -ForegroundColor Red
                Exit 1
            }
        }
        else{
            Exit 1
        }
    }
}
If ($OU){
    if (!(Get-Command "Get-ADComputer" -ErrorAction SilentlyContinue)){ 
        Set-Message "Error. Get-ADComputer not found on system. You have to install the PowerShell Active Directory module order to query Active Directory. https://4sysops.com/archives/how-to-install-the-powershell-active-directory-module/" -ForegroundColor Red
        Exit 1
    }
    try{
        $ComputerList = Get-ADComputer -Filter * -SearchBase "$OU" | Select-Object -Expand name
    }catch{
        Set-Message "Error querying AD: $($_.Exception.Message)" -ForegroundColor Red
        Exit 1
    }
}
ElseIf ($CSV){
    try{
        $ComputerList = Get-Content $CSV | where {$_ -notmatch 'Name'} | Foreach-Object {$_ -replace '"', ''}
    }catch{
        Set-Message "Error getting CSV content: $($_.Exception.Message)" -ForegroundColor Red
        Exit 1
    }
}
ElseIf(!$ComputerList){
    Set-Message "You have to set a list of computers, OU or CSV." -ForegroundColor Red
    Exit 1
}
If ($Credential){
    $Cred = Get-Credential
}
If(!$Cred -or !$Credential){
    Set-Message "No credential specified. Using logon account"
}
Else{
    Set-Message "Using user $($Cred.UserName)"
}

$ApplicationName = $AppPath.Substring($AppPath.LastIndexOf('\')+1)
$ApplicationFolderPath = $AppPath.Substring(0,$AppPath.LastIndexOf('\'))
$ApplicationFolderName = $ApplicationFolderPath.Substring($ApplicationFolderPath.LastIndexOf('\')+1)
$ComputerWithError = [System.Collections.ArrayList]@()
$ComputerWithSuccess = [System.Collections.ArrayList]@()
$ComputerSkipped = [System.Collections.ArrayList]@()
$TotalRetries = $Retries
$TotalComputers = $ComputerList.Count
Do{
    Set-Message "-----------------------------------------------------------------"
    Set-Message "Attempt $(($TotalRetries - $Retries) +1) of $($TotalRetries)" -ForegroundColor Cyan
    Set-Message "-----------------------------------------------------------------"
    $Count = 1
    ForEach ($Computer in $ComputerList){
        Set-Message "COMPUTER $($Computer.ToUpper()) ($($Count) of $($ComputerList.Count))" -ForegroundColor Yellow
        $Count++
        If($AppName -ne "None"){
            Set-Message "Checking if $($AppName) version $($AppVersion) is installed on remote computer."
            try{
                If(CheckSoftwareInstalled $Computer){
                    Set-Message "Software found on computer. Skipping installation." -ForegroundColor Green
                    $ComputerSkipped.Add($Computer) | Out-Null
                    Continue
                }
                Else{
                    Set-Message "Software not found on remote computer."
                }
            }catch{
                Set-Message "Error connecting: $($_.Exception.Message)" -ForegroundColor Red
                CheckPSRemoting $Computer
                $ComputerWithError.Add($Computer) | Out-Null
                Continue
            }
        }
        If($WMIQuery -ne "None"){
            Set-Message "Checking WMI Query on remote computer."
            try{
                If(!(CheckWMIQuery $Computer)){
                    Set-Message "WMI Query result is false. Skipping installation."
                    $ComputerSkipped.Add($Computer) | Out-Null
                    Continue
                }
                Else{
                    Set-Message "WMI Query result is true. Continue installation."
                }
            }catch{
                Set-Message "Error connecting: $($_.Exception.Message)" -ForegroundColor Red
                CheckPSRemoting $Computer
                $ComputerWithError.Add($Computer) | Out-Null
                Continue
            }
        }
        Set-Message "Coping $($ApplicationFolderPath) to \\$($Computer)\$($LocalPath -replace ':','$')"
        try{
            Copy-WithProgress "$ApplicationFolderPath" "\\$($Computer)\$("$($LocalPath)\$($ApplicationFolderName)" -replace ':','$')"
        }catch{
                Set-Message "Error copying folder: $($_.Exception.Message)" -ForegroundColor Red
                $ComputerWithError.Add($Computer) | Out-Null
                Continue;
            }
        try{
            $ExitCode = InstallRemoteSoftware $Computer
            If ($ExitCode){
                $ComputerWithError.Add($Computer) | Out-Null
                Set-Message "Error installing $($ApplicationName)." -ForegroundColor Red
            }
            else{
                Set-Message "$($ApplicationName) installed successfully." -ForegroundColor Green
                $ComputerWithSuccess.Add($Computer) | Out-Null
            }
            }catch{
                Set-Message "Error on remote execution: $($_.Exception.Message)" -ForegroundColor Red
                $ComputerWithError.Add($Computer) | Out-Null
                try{
                    Set-Message "Deleting \\$($Computer)\$($LocalPath -replace ':','$')\$($ApplicationFolderName)"
                }catch{
                    Set-Message "Error on remote deletion: $($_.Exception.Message)" -ForegroundColor Red
                }
                Remove-Item "\\$($Computer)\$($LocalPath -replace ':','$')\$($ApplicationFolderName)" -Force -Recurse
                CheckPSRemoting $Computer
            }
    }
    If ($ComputerWithError.Count -eq 0){
        break
    }
    $Retries--
    If ($Retries -gt 0){
        $ComputerList=$ComputerWithError
        $ComputerWithError = [System.Collections.ArrayList]@()
        If ($TimeBetweenRetries -gt 0){
            Set-Message "Waiting $($TimeBetweenRetries) seconds before next retry..."
            Sleep $TimeBetweenRetries
        }
    }
}While ($Retries -gt 0)

If($ComputerWithError.Count -gt 0){
    Set-Message "-----------------------------------------------------------------"
    Set-Message "Error installing $($ApplicationName) on $($ComputerWithError.Count) of $($TotalComputers) computers:"
    Set-Message $ComputerWithError
    $csvContents = @()
    ForEach($Computer in $ComputerWithError){
        $row = New-Object System.Object
        $row | Add-Member -MemberType NoteProperty -Name "Name" -Value $Computer
        $csvContents += $row
    }
    $CSV=(get-date).ToString('yyyyMMdd-HH_mm_ss') + "ComputerWithError.csv"
    $csvContents | Export-CSV -notype -Path "$([Environment]::GetFolderPath("MyDocuments"))\$($CSV)" -Encoding UTF8
    Set-Message "Computers with error exported to CSV file: $([Environment]::GetFolderPath("MyDocuments"))\$($CSV)" -ForegroundColor DarkYellow
    Set-Message "You can retry failed installation on this computers using parameter -CSV $([Environment]::GetFolderPath("MyDocuments"))\$($CSV)" -ForegroundColor DarkYellow
}
If ($ComputerWithSuccess.Count -gt 0){
    Set-Message "-----------------------------------------------------------------"
    Set-Message "$([math]::Round((($ComputerWithSuccess.Count * 100) / $TotalComputers), [System.MidpointRounding]::AwayFromZero) )% Success installing $($ApplicationName) on $($ComputerWithSuccess.Count) of $($TotalComputers) computers:"
    Set-Message $ComputerWithSuccess
}
Else{
    Set-Message "-----------------------------------------------------------------"
    Set-Message "Installation of $($ApplicationName) failed on all computers" -ForegroundColor Red
}
If ($ComputerSkipped.Count -gt 0){
    Set-Message "-----------------------------------------------------------------"
    Set-Message "$($ComputerSkipped.Count) skipped of $($TotalComputers) computers:"
    Set-Message $ComputerSkipped
}