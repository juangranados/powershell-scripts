<#
.SYNOPSIS
    Install/Update/Uninstall software using RZGet repository from https://ruckzuck.tools/
.DESCRIPTION
    Install/Update/Uninstall software using RZGet repository from https://ruckzuck.tools/
.PARAMETER tempFolder
    Folder to download installers
    Default: "C:\temp\InstallSoftware"
.PARAMETER software
    List of software to install. Check https://ruckzuck.tools/Home/Repository
    Default: None
    Example: "7-Zip","Notepad++","Edge","3CXPhone for Windows","Google Chrome","Teams","Postman"
.PARAMETER logFolder
    Log file path.
    Default: "C:\temp\InstallSoftware"
    Example: "\\ES-CPD-BCK02\scripts\InstallSoftware\Log"
.PARAMETER uninstall
    Uninstall software if it is already installed
.PARAMETER checkOnly
    Check if software list is already installed
.PARAMETER runAsAdmin
    Check if script is elevated
.EXAMPLE
    .\Install-Software -tempFolder C:\temp\InstallSoftware -software "7-Zip","Notepad++","Edge","3CXPhone for Windows","Google Chrome","Teams","Postman" -logFolder "\\ES-CPD-BCK02\scripts\InstallSoftware\Log"
.LINK
    https://github.com/juangranados/powershell-scripts/tree/main/Install%20Software%20Locally
.NOTES
    Thanks to Roger Zander for his amazing tool: https://ruckzuck.tools/
    Author: Juan Granados
#>
Param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$tempFolder = 'C:\temp\InstallSoftware',
    [Parameter(Mandatory = $true)]
    [string[]]$software,
    [Parameter(Mandatory = $false)] 
    [string]$logFolder = 'C:\temp\InstallSoftware',
    [Parameter()]
    [switch]$uninstall,
    [Parameter()]
    [switch]$checkOnly,
    [Parameter()]
    [switch]$runAsAdmin
)
function Set-Folder([string]$folderPath) {
    if ($folderPath.Chars($folderPath.Length - 1) -eq '\') {
        $folderPath = ($folderPath.TrimEnd('\'))
    }
    if (!(Test-Path $folderPath)) {
        try {
            New-Item $folderPath -ItemType directory
        }
        catch {
            Write-Error "Error creating $folderPath"
            try {
                Stop-Transcript
            }
            catch { Write-Warning $Error[0] }
            Exit 1
        }
    }
}
function Get-FileHashIsOk([string]$filePath, [string]$hashType, [string]$hash) {
    if ($hashType.ToUpper() -eq "X509") {
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($filePath)
        $certHash = $cert.GetCertHashString().ToLower().Replace(" ", "")
        if ($certHash -ne $hash) {
            return $false
        }
        else {
            return $true
        }
    }
    elseif (($hashType.ToUpper() -eq "MD5")) {
        $fileHash = Get-FileHash $filePath -Algorithm MD5
        if ($fileHash.Hash -ne $hash) {
            return $false
        }
        else {
            return $true
        }
    }
    elseif (($hashType.ToUpper() -eq "SHA1")) {
        $fileHash = Get-FileHash $filePath -Algorithm SHA1
        if ($fileHash.Hash -ne $hash) {
            return $false
        }
        else {
            return $true
        }
    }
    elseif (($hashType.ToUpper() -eq "SHA256")) {
        $fileHash = Get-FileHash $filePath -Algorithm SHA256
        if ($fileHash.Hash -ne $hash) {
            return $false
        }
        else {
            return $true
        }
    }
}
function Get-AppInstaller ($files) {
    foreach ($file in $files) {
        $filePath = "$tempFolder\$($file.FileName)"
        try {
            if (-not $file.URL.StartsWith("http")) {
                try {
                    $file.URL = Invoke-Expression -Command $file.URL
                }
                catch { Write-Warning $Error[0] }
                if (-not $file.URL) {
                    Write-Warning "Error getting file URL"
                    return $false
                }
            }
            Write-Host "Downloading $($file.URL)"
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest $file.URL -OutFile $filePath  -UseBasicParsing
        }
        catch {
            Write-Warning "Error downloading file"
            return $false
        }
        if (Test-Path $filePath) {
            Write-Host "File downloaded"
        }
        else {
            Write-Warning "File $filePath not found"
            return $false
        } 
        if ((Get-Item $filePath).Length -eq $file.FileSize) {
            Write-Host "File size is ok"
        }
        else {
            Write-Warning "File size mismatch"
        } 
        if (Get-FileHashIsOk $filePath $file.HashType $file.FileHash) {
            Write-Host "File hash is ok"
        }
        else {
            Write-Warning "File hash mismatch"
            return $false
        }
    }
    return $true
}
function Invoke-FilesDeletion ($files) {
    foreach ($file in $files) {
        $filePath = "$tempFolder\$($file.FileName)"
        Write-Host "Deleting file $filePath"
        Remove-Item -Path $filePath -Force -Confirm:$false
    }
}
#https://cdn.ruckzuck.tools/rest/v2/GetCatalog
$ErrorActionPreference = 'Stop'
Set-Folder $tempFolder
Set-Folder $logFolder
Set-Location $tempFolder
$transcriptFile = "$logFolder\$(get-date -Format yyyy_MM_dd)_$($env:COMPUTERNAME)_InstallSoftware.txt"
try {
    Start-Transcript $transcriptFile
}
catch { Write-Warning "Start-Transcript can not be started: $($Error[0])" }

if ($runAsAdmin) {
    Write-Host "Checking for elevated permissions"
    if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Error "Insufficient permissions to run this script. Execute PowerShell script as an administrator."
        try {
            Stop-Transcript
        }
        catch { Write-Warning $Error[0] }
        Exit 1
    }
    Write-Host "Script is elevated"
}
Write-Host "Checking for software in computer"
if ($software.Count -eq 1) { 
    $software = $software.Split(',') 
}
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$apiUrl = Invoke-RestMethod -Uri "https://ruckzuck.tools/rest/v2/geturl" -UseBasicParsing
if ([string]::IsNullOrEmpty($apiUrl)) {
    Write-Error "RuckZuck API can not be found"
}
foreach ($app in $software) {
    $ULine = '-' * $app.Length
    Write-Host -Object $ULine -ForegroundColor DarkCyan
    Write-Host $app -ForegroundColor DarkCyan
    Write-Host -Object $ULine -ForegroundColor DarkCyan
    $uriApp = [uri]::EscapeDataString($app)
    try {
        $ProgressPreference = 'SilentlyContinue'
        $appJson = Invoke-WebRequest "$apiUrl/rest/v2/getsoftwares?shortname=$uriApp"  -UseBasicParsing | ConvertFrom-Json
    }
    catch {
        Write-Warning $Error[0]
    }
    if ($appJson) {
        $isSoftwareInstalled = Invoke-Expression -Command $appJson.PSDetection
        if ($isSoftwareInstalled) {
            Write-Host "$app is installed on computer" -ForegroundColor Green
            if ($uninstall -and -not $checkOnly) {
                Write-Host "Running uninstall command"
                try {
                    Invoke-Expression -Command $appJson.PSUninstall
                    if ($ExitCode -eq 0) {
                        Write-Host "$app uninstallation sucessful" -ForegroundColor Green
                    }
                    else {
                        Write-Warning "$app uninstallation returned $ExitCode"
                    }
                }
                catch {
                    Write-Warning $Error[0]
                }
            }
        }
        else {
            Write-Host "$app not found or older version is installed." -ForegroundColor Yellow
            if (-not $checkOnly -and -not $uninstall) {
                if ($appJson.PSPreReq) {
                    If (Get-AppInstaller $appJson.Files) {
                        try {
                            if (-not [string]::IsNullOrEmpty($appJson.PSPreInstall)) {
                                Write-Host "Running pre install command"
                            
                                Invoke-Expression -Command $appJson.PSPreInstall
                            
                            }
            
                            Write-Host "Running install command"
                            Invoke-Expression -Command $appJson.PSInstall
                            if ($ExitCode -eq 0) {
                                Write-Host "$app installation sucessful" -ForegroundColor Green
                            }
                            else {
                                Write-Warning "$app installation returned $ExitCode"
                            }
                            if (-not [string]::IsNullOrEmpty($appJson.PSPostInstall)) {
                                Write-Host "Running post install command"
                                Invoke-Expression -Command $appJson.PSPostInstall
                            }
                            Invoke-FilesDeletion $appJson.Files
                        }
                        catch {
                            Write-Warning $Error[0]
                        }
                    }
                }
                else {
                    Write-Warning "$app can not be installed on computer because $($appJson.PSPreReq) is false"
                }
            }
        }
    }
    else {
        Write-Warning "$app not found in RuckZuck repository"
    }
}
Set-Location $PSScriptRoot
try {
    Stop-Transcript
}
catch { Write-Warning $Error[0] }