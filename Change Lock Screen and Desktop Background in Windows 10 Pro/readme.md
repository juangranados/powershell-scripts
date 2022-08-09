# **Change Lock Screen and Desktop Background in Windows 10 Pro**

[Right click here and select "Save link as" to download](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Change%20Lock%20Screen%20and%20Desktop%20Background%20in%20Windows%2010%20Pro/Set-LockScreen.ps1)

This script allows you to change login screen and desktop background in Windows 10 Professional using GPO startup script.

By default, lock screen can not be changed by GPO in Windows 10  Professional, with this script you can change it to comply with  corporate image.

Create a GPO to run this PowerShell script.

Examples:

Set Lock Screen and Desktop Wallpaper with logs:

```powershell
.\Set-LockScreen.ps1 -LockScreenSource "\\SERVER-FS01\LockScreen.jpg" -BackgroundSource "\\SERVER-FS01\BackgroundScreen.jpg" -LogPath "\\SERVER-FS01\Logs"
```

Set Lock Screen and Desktop Wallpaper without logs:

```powershell
.\Set-LockScreen.ps1 -LockScreenSource "\\SERVER-FS01\LockScreen.jpg" -BackgroundSource "\\SERVER-FS01\BackgroundScreen.jpg"
```

Set Lock Screen only:

```powershell
.\Set-LockScreen.ps1 -LockScreenSource "\\SERVER-FS01\LockScreen.jpg" -LogPath "\\SERVER-FS01\Logs"
```

Set Desktop Wallpaper only:

```powershell
.\Set-LockScreen.ps1 -BackgroundSource "\\SERVER-FS01\BackgroundScreen.jpg" -LogPath "\\SERVER-FS01\Logs"
```

```powershell
<# 
.SYNOPSIS 
    Change Lock Screen and Desktop Background in Windows 10 Pro. 
.DESCRIPTION 
    This script allows you to change logon screen and desktop background in Windows 10 Professional using GPO startup script. 
.PARAMETER LockScreenSource (Optional) 
    Path to the Lock Screen image to copy locally in computer. 
    Example: "\\SERVER-FS01\LockScreen.jpg" 
.PARAMETER BackgroundSource (Optional) 
    Path to the Desktop Background image to copy locally in computer. 
    Example: "\\SERVER-FS01\BackgroundScreen.jpg" 
.PARAMETER LogPath (Optional) 
    Path where save log file. If it's not specified no log is recorded. 
.EXAMPLE 
    Set Lock Screen and Desktop Wallpaper with logs: 
    Set-Loca Screen -LockScreenSource "\\SERVER-FS01\LockScreen.jpg" -BackgroundSource "\\SERVER-FS01\BackgroundScreen.jpg" -LogPath "\\SERVER-FS01\Logs" 
.EXAMPLE 
    Set Lock Screen and Desktop Wallpaper without logs: 
    Set-LockScreen -LockScreenSource "\\SERVER-FS01\LockScreen.jpg" -BackgroundSource "\\SERVER-FS01\BackgroundScreen.jpg" 
.EXAMPLE 
    Set Lock Screen only: 
    .\Set-Screen.ps1 -LockScreenSource "\\SERVER-FS01\LockScreen.jpg" -LogPath "\\SERVER-FS01\Logs" 
.EXAMPLE 
    Set Desktop Wallpaper only: 
    .\Set-LockScreen.ps1 -BackgroundSource "\\SERVER-FS01\BackgroundScreen.jpg" -LogPath "\\SERVER-FS01\Logs" 
.NOTES  
    Author: Juan Granados  
#>
```
