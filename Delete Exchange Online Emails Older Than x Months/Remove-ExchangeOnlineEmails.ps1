<#PSScriptInfo

.VERSION 1.0

.GUID 

.AUTHOR Juan Granados

.COPYRIGHT 2022 Juan Granados

.TAGS Exchange Online Remove Messages

.LICENSEURI https://raw.githubusercontent.com/juangranados/powershell-scripts/main/LICENSE

.PROJECTURI https://github.com/juangranados/powershell-scripts/tree/main/Delete%20Exchange%20Online%20Emails%20Older%20Than%20x%20Months

.RELEASENOTES Initial release

#>

<#
.SYNOPSIS
    Delete Exchange Online Emails Older Than x Months.
.DESCRIPTION
    Delete Exchange Online Emails Older Than x Months from a mailbox or distribution group.
    Running user has to be member of eDiscovery Manager group: https://docs.microsoft.com/en-us/microsoft-365/compliance/assign-ediscovery-permissions?view=o365-worldwide
.PARAMETER target
    Mailbox or distribution group to delete emails from.
.PARAMETER months
    Emails older than this number of months will be deleted
.PARAMETER logPath
    Path where save log file.
    Default: Temp folder
.EXAMPLE
    Defrag-WinSearchDB -LogPath "\\ES-CPD-BCK02\Log" -TempPath \\ES-CPD-BCK02\Temp
.EXAMPLE
    Defrag-WinSearchDB -TempPath "D:\Temp"
.LINK
    https://github.com/juangranados/powershell-scripts/tree/main/Delete%20Exchange%20Online%20Emails%20Older%20Than%20x%20Months
.NOTES
    Author: Juan Granados 
#>
Param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$target,
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [ValidateRange("Positive")]
    [int]$months,
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$logPath = $env:temp
)
Connect-ExchangeOnline
Connect-IPPSSession

$searchName = "target_messages_older_than_$($months)_months"
if (Get-ComplianceSearch -Identity $searchName -ErrorAction SilentlyContinue) {
    Write-Host "Compliance Search: $searchName exists. Changing properties"
    Set-ComplianceSearch -Name $searchName -ExchangeLocation $target -ContentMatchQuery "(Received <= $((get-date).AddMonths(-$months).ToString("MM/dd/yyy")))"
}
else {
    Write-Host "Creationg Compliance Search: $searchName"
    New-ComplianceSearch -Name $searchName -ExchangeLocation $target -ContentMatchQuery "(Received <= $((get-date).AddMonths(-$months).ToString("MM/dd/yyy")))"
}
Write-Host "Running Compliance Search: $searchName"
Start-ComplianceSearch -Identity $searchName

While ((Get-ComplianceSearch  -Identity $searchName).status -ne "Completed") {
    Write-Host "."
    Start-Sleep 5
}

$searchActionName = "$($searchName)_preview"
if (Get-ComplianceSearchAction -Identity $searchActionName -ErrorAction SilentlyContinue) {
    Write-Host "Compliance Search Action : $searchName exists. Deleting"
    Remove-ComplianceSearchAction -Identity $searchActionName -Confirm:$false
}
Write-Host "Creating Compliance Search Action : $searchActionName"
New-ComplianceSearchAction SearchName $searchName -Preview
Write-Host "Waiting for Compliance Search Action to finish"
While ((Get-ComplianceSearchAction -Identity $searchActionName).status -ne "Completed") {
    Write-Host "."
    Start-Sleep 5
}
Get-ComplianceSearchAction  -Identity $searchActionName | Format-List -Property Results
#Delete
#New-ComplianceSearchAction -SearchName $searchActionName -Purge -PurgeType SoftDelete
#Get-ComplianceSearchAction  -Identity $searchActionName | Format-List -Property Results