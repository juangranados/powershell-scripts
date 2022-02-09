# Delete Windows 10 Preinstalled Apps 

Windows 10 includes a variety of universal apps like Candy Crush, Facebook, ect. All of these can be uninstalled with a single PowerShell Command.

This command uninstalls all built-in apps except Windows Store, Calculator and Photos.

## Deletes current user apps

```powershell
Get-AppxPackage | where-object {$_.name -notlike "*Microsoft.WindowsStore*"} | where-object {$_.name -notlike "*Microsoft.WindowsCalculator*"} | where-object {$_.name -notlike "*Microsoft.Windows.Photos*"} | Remove-AppxPackage
```
## Deletes all users apps (current and new).

Run PowerShell console as administrator.

```powershell
Get-AppxPackage -AllUsers | where-object {$_.name -notlike "*Microsoft.WindowsStore*"} | where-object {$_.name -notlike "*Microsoft.WindowsCalculator*"} | where-object {$_.name -notlike "*Microsoft.Windows.Photos*"} | Remove-AppxPackage

Get-AppxProvisionedPackage -online | where-object {$_.name -notlike "*Microsoft.WindowsStore*"} | where-object {$_.name -notlike "*Microsoft.WindowsCalculator*"} | where-object {$_.name -notlike "*Microsoft.Windows.Photos*"} | Remove-AppxProvisionedPackage -online
```