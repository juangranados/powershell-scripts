@echo off
SETLOCAL
SET computer=192.168.0.26
SET scriptPath="\\SRVHTS-FS01\Scripts\RemoteComputerUpdate\Launch-Remote.ps1"
psexec.exe -s \\%computer% powershell.exe -ExecutionPolicy Bypass -file %scriptPath%
pause