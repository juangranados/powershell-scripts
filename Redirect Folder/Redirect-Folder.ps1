<#
.SYNOPSIS
    Redirects user shell folders.
.DESCRIPTION
    Redirects user shell folders to a custom path.
.PARAMETER logPath
    Log file path.
    Default "Documents"
    Example: "\\ES-CPD-BCK02\scripts\FolderRedirection\Log"
.PARAMETER folders
    Array of folders to redirect.
    Check folder names in HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders
    Example: "Personal","My Pictures","Desktop"
.PARAMETER paths
    Array of paths to redirect, in the same order than folders array.
    Example: "D:\%USERNAME%\Documents","D:\%USERNAME%\Pictures","D:\%USERNAME%\Desktop"
.EXAMPLE
    .\Redirect-Folder -folders "Personal","My Pictures","Desktop" -paths "D:\%USERNAME%\Documents","D:\%USERNAME%\Pictures","D:\%USERNAME%\Desktop"
.EXAMPLE
    .\Redirect-Folder -folders "Personal" -paths "\\SRV-FS01\Users\%USERNAME%\Documents" -logPath "\\SRV-DC01\Scripts\FolderRedirection\Log"
.LINK
    https://github.com/juangranados/powershell-scripts/tree/main/Redirect%20Folder
.NOTES
    Author: Juan Granados 
#>
Param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]]$folders,
    [Parameter(Mandatory = $true)]
    [string[]]$paths,
    [Parameter(Mandatory = $false)] 
    [string]$logPath = [Environment]::GetFolderPath("MyDocuments")
)
Function Set-RegistryKey ([string]$path, [string]$name, $value) {

    $currentValue = (Get-Item -Path $path).GetValue($name, $null, 'DoNotExpandEnvironmentNames')
    if (-not $currentValue) {
        Write-Host "Key $($path)\$($name) do not exist - Changing by $($value)" -ForegroundColor Yellow
        Set-ItemProperty -Path $path -name $name -Value $value -Type 'ExpandString'
    }
    elseif ($currentValue -ne $value) {
        Write-Host "Key $($path)\$($name) has a value of $($currentValue) - Changing by $($value)" -ForegroundColor Yellow
        Set-ItemProperty -Path $path -name $name -Value $value -Type 'ExpandString'
    }
    else {
        Write-Host "Key $($path)\$($name) already has the value of $($currentValue)" -ForegroundColor Green
    }
}
Start-Transcript $logPath
$regKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
$i = 0
foreach ($folder in $folders) {
    Write-Host "Setting $folder as $($paths[$i])"
    Set-RegistryKey -path $regKey -name $folder -value $paths[$i]
    $i++
}
Stop-Transcript
