Psexec.exe -s -i -accepteula powershell.exe -Command (Remove-Item -Path 'HKLM:\SYSTEM\CurrentControlSet\Enum\SWD\PRINTENUM\*' -Recurse -Force)
Psexec.exe -s -i -accepteula powershell.exe -Command (Remove-Item -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceClasses\{0ecef634-6ef0-472a-8085-5ad023ecbccd}\*' -Recurse -Force)
Psexec.exe -s -i -accepteula powershell.exe -Command (Remove-Item -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\Client Side Rendering Print Provider\*' -Recurse -Force)
Psexec.exe -s -i -accepteula powershell.exe -Command (Remove-Item -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Printers\*' -Recurse -Force)
Psexec.exe -s -i -accepteula powershell.exe -Command (Remove-Item -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\V4 Connections\*' -Recurse -Force)
@echo off
set /p r= Reboot computer? [y/n]
if %r% == y goto reboot
exit
:reboot
shutdown -r -f -t 0