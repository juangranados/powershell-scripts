Function Set-PrinterDriverAttributes([string]$RegistryPath){
    $Changes = 0
    Write-Host "Checking $($RegistryPath)"
    $Printers = Get-ChildItem -Path "$RegistryPath" -Recurse
    ForEach ($Printer in $Printers){
        $PrinterDriverAttributes = $Printer.GetValue("PrinterDriverAttributes")
        If($PrinterDriverAttributes % 2 -eq 0){
            Write-Host "Printer driver $($Printer) has PrinterDriverAttributes value of $($PrinterDriverAttributes)" -ForegroundColor Yellow
            Write-Host "Changing PrinterDriverAttributes to $($PrinterDriverAttributes + 1)" -ForegroundColor Yellow
            try{
            New-ItemProperty -Path $Printer.PSPath -Name PrinterDriverAttributes -PropertyType DWord -Value $($PrinterDriverAttributes + 1) -Force -ErrorAction Continue | Out-Null
            $Changes++
            } catch {
                Write-Host "Error changing registry key" -ForegroundColor Red
                Write-Host "Excepcion: $($_.Exception.GetType().FullName)" -ForegroundColor Red
                Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red 
            }
        }
        Else{
            Write-Host "Printer driver $($Printer) has PrinterDriverAttributes value of $($PrinterDriverAttributes)" -ForegroundColor DarkCyan
        }
    }
    Return $Changes
}

$ErrorActionPreference = "Stop"

If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Error "You need to run script as administrador"
    Exit(1)
}
$DriversChanged = Set-PrinterDriverAttributes "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows NT x86\Drivers\Version-3"
$DriversChanged += Set-PrinterDriverAttributes "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\Windows x64\Drivers\Version-3"
If ($DriversChanged){
    Write-Host "$($DriversChanged) registry keys changed." -ForegroundColor Green
    Write-Host "Restarting Spooler." -ForegroundColor Yellow
    Restart-Service Spooler
}
Else{
    Write-Host "All PrinterDriverAttributes registry keys OK." -ForegroundColor Green 
}