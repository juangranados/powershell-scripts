<#
.SYNOPSIS
	Download Java Runtime Environment (jre) latest version for Windows for later installation.
.DESCRIPTION
	This script Download Java Runtime Environment 32 and 64 bit (jre) latest version for Windows for later installation.
.PARAMETER DownloadPath
	Path to download both JRE installers: 32 and 64 bit.
    Example: D:\Software\JRE
.PARAMETER LogPath
    Log path (optional).
    Example: D:\Software\JRE\Log
.NOTES 
	Author: Juan Granados 
	Date:   April 2022
#>
Param(
    [Parameter(Mandatory = $true, Position = 0)] 
    [ValidateNotNullOrEmpty()]
    [string]$downloadPath,
    [Parameter(Mandatory = $false, Position = 1)] 
    [ValidateNotNullOrEmpty()]
    [string]$logPath
)
#Requires -RunAsAdministrator
if (-not [string]::IsNullOrWhiteSpace($logPath) -and $logPath.Chars($logPath.Length - 1) -eq '\') {
    $logPath = ($logPath.TrimEnd('\'))
}
if ($downloadPath.Chars($downloadPath.Length - 1) -eq '\') {
    $downloadPath = ($downloadPath.TrimEnd('\'))
}
if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
    Start-Transcript -Path "$($logPath)\$($env:COMPUTERNAME)_downloadJRE.log" | Out-Null
}
Write-Host "Surfing https://www.java.com/en/download/manual.jsp"

$URL = "https://www.java.com/en/download/manual.jsp"
$global:ie = New-Object -com "InternetExplorer.Application"
$global:ie.visible = $false
$global:ie.Navigate($URL)

do {
    Start-Sleep -s 1
} until(!($global:ie.Busy))
$global:doc = $global:ie.Document

$links = @($global:doc.links)
$ProgressPreference = 'SilentlyContinue'
$link = $links | Where-Object { $_.href -like "http*" } | Where-Object { $_.title -like "Download Java software for Windows (64-bit)" }
Write-Host "Downloading JRE x64"
Invoke-WebRequest $link[0].href -OutFile $downloadPath\jrex64.exe
$link = $links | Where-Object { $_.href -like "http*" } | Where-Object { $_.title -like "Download Java software for Windows Offline" }
Write-Host "Downloading JRE x32"
Invoke-WebRequest $link[0].href -OutFile $downloadPath\jrex32.exe

if (-not [string]::IsNullOrWhiteSpace($logPath)) {
    Stop-Transcript
}