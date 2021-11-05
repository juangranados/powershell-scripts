$logPath = "\\ES-CPD-BCK02\scripts\WindowsUpdate\Log"
$rebootMessage = "Se va a reiniciar el equipo dentro de 2 horas para terminar de instalar las actualizaciones de Windows. Por favor, cierra todo antes de esa hora o reinicia el equipo manualmente"
$RZGetPath = "\\ES-CPD-BCK02\scripts\WindowsUpdate\RZGet.exe"
[int]$rebootHours = 2
\\SRVHTS-FS01\Scripts\RemoteComputerUpdate\Update-Computer.ps1 -logPath $logPath -scheduleReboot -rebootHours $rebootHours -rebootMessage $rebootMessage -RZGetPath $RZGetPath