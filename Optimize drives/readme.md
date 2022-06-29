# **Optimize or defrag Windows Drives**

* [Invoke-DiskOptimize.ps1](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Optimize%20drives/Invoke-DiskOptimize.ps1): Runs Windows disks optimization in all or selected drives.
* [Invoke-DiskDefrag.ps1](https://raw.githubusercontent.com/juangranados/powershell-scripts/main/Optimize%20drives/Invoke-DiskDefrag.ps1): Runs Windows disks defragmentation in all or selected drives.

Examples:

Optimize all drives.

```powershell
Invoke-DiskOptimize.ps1 -LogPath "\\SERVER-FS01\Logs"
```

Optimize only C and D drives.

```powershell
Invoke-DiskOptimize.ps1 -disks "C:","D:"
```

Defrag all drives if they are 10% fragmented (default value).

```powershell
Invoke-DiskDefrag.ps1
```

Defrag only C and D drives if they are 20% fragmented.

```powershell
Invoke-DiskDefrag.ps1 -disks "C:","D:" -defragPercentage 20 -LogPath "\\SERVER-FS01\Logs"
```

Defrag C: drive if they are 10% fragmented. It runs disk defragmentation even C: disk free space is low.

```powershell
Invoke-DiskDefrag.ps1 -disks "C:" -forceDefrag
```

```powershell
<#
.SYNOPSIS
    Runs Windows disks optimization.
.DESCRIPTION
    Runs Windows disks optimization in all or selected drives.
.PARAMETER disks
    Disks to run optimization.
    Default: all.
    Example: "C:","D:","F:"
.PARAMETER LogPath
    Path where save log file.
    Default: Temp folder
.EXAMPLE
    Optimize all drives.
    Invoke-DiskOptimize.ps1
.EXAMPLE
    Optimize only C and D drives.
    Invoke-DiskOptimize.ps1 -disks "C:","D:"
.NOTES 
    Author:Juan Granados
#>
```

```powershell
<#
.SYNOPSIS
    Runs Windows disks defragmentation.
.DESCRIPTION
    Runs Windows disks defragmentation in all or selected drives.
.PARAMETER disks
    Disks to run defragmentation.
    Default: all.
    Example: "C:","D:","F:"
.PARAMETER defragPercentage
    Percentage of fragmentation in order to run defragmentation.
    Default: 10
.PARAMETER forceDefrag
    Defrag disks if free space is low.
    Default: false
.PARAMETER LogPath
    Path where save log file.
    Default: Temp folder
.EXAMPLE
    Defrag all drives if they are 10% fragmented.
    Invoke-DiskDefrag.ps1
.EXAMPLE
    Defrag only C and D drives if they are 20% fragmented.
    Invoke-DiskDefrag.ps1 -disks "C:","D:" -defragPercentage 20
.EXAMPLE
    Defrag C: drive if they are 10% fragmented. It runs disk defragmentation even C: disk free space is low.
    Invoke-DiskDefrag.ps1 -disks "C:" -forceDefrag
.NOTES 
    Author:Juan Granados
#>
```
