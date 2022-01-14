#Requires -RunAsAdministrator
$folder = 'C:\Program Files\WindowsApps'
takeown.exe /f $($folder) /R /D S
icacls.exe $($folder) /grant "$([Security.Principal.WindowsIdentity]::GetCurrent().Name):(OI)(CI)F" /T
$items = Get-ChildItem $folder -ErrorAction SilentlyContinue | Select-Object FullName, Mode, Name
foreach ($item in $items) {
    Write-Host "Procesando $($item.FullName)"
    takeown.exe /f $($item.FullName) /R /D S >$null
    icacls.exe $($item.FullName) /grant "$([Security.Principal.WindowsIdentity]::GetCurrent().Name):(OI)(CI)F" /T >$null
}