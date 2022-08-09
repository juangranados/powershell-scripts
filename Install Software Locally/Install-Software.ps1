<#
.SYNOPSIS
    Install software using RZGet from ruckzuck.tools
.DESCRIPTION
    Install software using RZGet from ruckzuck.tools
.PARAMETER folder
    Folder to download RZGet.exe
    Default: "C:\temp\InstallSoftware"
.PARAMETER RZGetArguments
    RZget Arguments. Check https://github.com/rzander/ruckzuck/wiki/RZGet and https://ruckzuck.tools/Home/Repository
    Default: "update --all"
    Example: 'install 7-Zip Notepad++(x64) Edge "3CXPhone for Windows" "Google Chrome"'
.PARAMETER logPath
    Log file path.
    Default: "C:\temp\InstallSoftware"
    Example: "\\ES-CPD-BCK02\scripts\InstallSoftware\Log"
.EXAMPLE
    .\Install-Software -folder C:\temp\InstallSoftware -RZGetArguments 'install 7-Zip Notepad++(x64) Edge "3CXPhone for Windows" "Google Chrome"' logPath "\\ES-CPD-BCK02\scripts\InstallSoftware\Log"
.LINK
    https://github.com/juangranados/powershell-scripts/tree/main/Install%20Software%20Locally
.NOTES
    Author: Juan Granados 
#>
Param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$folder = 'C:\temp\InstallSoftware',
    [Parameter(Mandatory = $false)]
    [string]$RZGetArguments = "update --all",
    [Parameter(Mandatory = $false)] 
    [string]$logPath = 'C:\temp\InstallSoftware'
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
        Write-Warning "An error has ocurred"
        Stop-Transcript
        Exit 1
    }
}
function Set-Folder([string]$folderPath) {
    if ($folderPath.Chars($folderPath.Length - 1) -eq '\') {
        $folderPath = ($folderPath.TrimEnd('\'))
    }
    if (!(Test-Path $folderPath)) {
        try {
            New-Item $folderPath -ItemType directory
        }
        catch {
            Write-Warning "Error creating $folderPath"
            Stop-Transcript
            Exit 1
        }
    }
}

$ErrorActionPreference = 'Stop'
Set-Folder $folder
Set-Folder $logPath

$transcriptFile = "$logPath\$(get-date -Format yyyy_MM_dd)_InstallSoftware.txt"
Start-Transcript $transcriptFile
Write-Host "Checking for elevated permissions"
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Insufficient permissions to run this script. Execute PowerShell script as an administrator."
    Stop-Transcript
    Exit 1
}
Write-Host "Script is elevated"
Write-Host "Downloading RZGet.exe"
Invoke-WebRequest "https://github.com/rzander/ruckzuck/releases/latest/download/RZGet.exe" -OutFile "$folder\RZGet.exe"
if (!(Test-Path "$folder\RZGet.exe")) {
    Write-Warning "Error downloading RZGet.exe"
    Stop-Transcript
    Exit 1
}
Write-Host "Running $($folder)\RZGet.exe $($RZGetArguments)"
Invoke-Process -FilePath "$folder\RZGet.exe" -ArgumentList $RZGetArguments -DisplayLevel StdOut
Stop-Transcript