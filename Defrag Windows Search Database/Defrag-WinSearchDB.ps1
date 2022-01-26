<#PSScriptInfo

.VERSION 1.0

.GUID c478611d-d1c2-4063-8f40-fa67874a3711

.AUTHOR Juan Granados

.COPYRIGHT 2021 Juan Granados

.TAGS WSUS Windows Update Remote Software

.LICENSEURI https://raw.githubusercontent.com/juangranados/powershell-scripts/main/LICENSE

.PROJECTURI https://github.com/juangranados/powershell-scripts/tree/main/Defrag%20Windows%20Search%20Database

.RELEASENOTES Initial release

#>

<#
.SYNOPSIS
    Defrag Windows Search Database.
.DESCRIPTION
    Defrag Windows Search Database and optionally deletes it after error.
.PARAMETER DataBase
    Windows Search Database path.
    Default C:\ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb
.PARAMETER TempPath
    Temporary folder to perform defrag. Using a different physical drive is recommended.
    Default: C:\ProgramData\Microsoft\Search\Data\Applications\Windows
.PARAMETER LogPath
    Log file path.
    Default "Documents"
.PARAMETER DeleteOnError
    Deletes Windows Search Database and modify registry to rebuild it at next Search Service startup.
    Default false. 
.EXAMPLE
    Defrag-WinSearchDB -LogPath "\\ES-CPD-BCK02\Log"
.LINK
    https://github.com/juangranados/powershell-scripts/tree/main/Defrag%20Windows%20Search%20Database
.NOTES
    Author: Juan Granados 
#>
Param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$DataBase = "$($env:ProgramData)\Microsoft\Search\Data\Applications\Windows\Windows.edb",
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$LogPath = [Environment]::GetFolderPath("MyDocuments"),
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$TempPath = "$($env:ProgramData)\Microsoft\Search\Data\Applications\Windows",
    [Parameter()]
    [switch]$DeleteOnError
)
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
$ErrorActionPreference = "SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Continue"

$LogPath = $LogPath.TrimEnd('\')
if (-not (Test-Path $LogPath)) {
    Write-Host "Log path $($LogPath) not found"
    Exit (1)
}

Start-Transcript -path "$($LogPath)\$(get-date -Format yyyy_MM_dd)_$($env:COMPUTERNAME).txt"

if (-not (Test-Path $DataBase) -or (($($DataBase.Substring($DataBase.Length - 4)) -ne ".edb"))) {
    Write-Host "Windows Search Database not found on $($DataBase)"
    Stop-Transcript
    Exit (1)
}

$TempPath = $TempPath.TrimEnd('\')
if (-not (Test-Path $TempPath)) {
    Write-Host "Temp path $($TempPath) not found"
    Exit (1)
}

$TempDataBase = $TempPath + "\tempdfrg_$(Get-Random).edb"

Write-Host "Disabling Windows Search"
Set-Service -Name 'wsearch' -StartupType 'Disabled'
Stop-Service wsearch -Force
Get-Service wsearch
Write-Host "Perform defrag on $($DataBase)"
$defragResult = Invoke-Process -FilePath "$([System.Environment]::SystemDirectory)\esentutl.exe" -ArgumentList "/d $($DataBase) /t $($TempDataBase)" -DisplayLevel Full
if ($defragResult.ExitCode -ne 0) {
    Write-Host "An error has ocurred: $($defragResult.ExitCode)"
    $defragResult.StdOut
    if ($DeleteOnError) {
        Write-Host "Deleting database $($DataBase)"
        Remove-Item $DataBase -Force
        New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows Search" -Name 'SetupCompletedSuccessfully' -Value 0 -PropertyType "DWord" -Force
    }
}
else {
    $defragResult.StdOut
}
Write-Host "Enabling Windows Search"
Invoke-Process -FilePath "$([System.Environment]::SystemDirectory)\sc.exe" -ArgumentLis "config wsearch start=delayed-auto" -DisplayLevel "StdOut"
Start-Service wsearch
Get-Service wsearch
Stop-Transcript