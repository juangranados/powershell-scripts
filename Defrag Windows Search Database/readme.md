# Defrag Windows Search Database

[Right click here and select "Save link as" to download](https://github.com/juangranados/powershell-scripts/tree/main/Defrag%20Windows%20Search%20Database/DefragWinSearchDB.ps1)

Script to Defrag Windows Search Database and optionally deletes it after error. 

![Screenshot](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Defrag%20Windows%20Search%20Database/screenshot.png)

## Parameters

### DataBase

Windows Search Database path.

Default C:\ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb

### TempPath

Temporary folder to perform defrag. Using a different physical drive is recommended.

Default: C:\ProgramData\Microsoft\Search\Data\Applications\Windows

### LogPath

Log file path.

Default: Documents.

Example: "\\ES-CPD-BCK02\SearchDefrag\Log"

### DeleteOnError

Deletes Windows Search Database and modify registry to rebuild it at next Search Service startup.

Default: false

## Example
```powershell
Defrag-WinSearchDB -LogPath "\\ES-CPD-BCK02\DefragWindDB\Log" -TempPath "\\ES-CPD-BCK02\DefragWindDB\Temp"
```

```powershell
Defrag-WinSearchDB -TempPath "D:\Temp"
```