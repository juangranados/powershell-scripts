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
.PARAMETER preview
    Only shows matching emails. It will not delete anything.
.EXAMPLE
    Remove-ExchangeOnlineEmails.ps1 -LogPath "C:\temp\Log" -target esmith@contoso.com -months 12
.EXAMPLE
    Remove-ExchangeOnlineEmails.ps1 -target finance@contoso.com -months 24
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
    [ValidateRange(0, 100)]
    [int]$months,
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$logPath = $env:temp,
    [Parameter()]
    [switch]$preview
)

$ErrorActionPreference = "SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Stop"

$logPath = $logPath.TrimEnd('\')
if (-not (Test-Path $logPath)) {
    Write-Host "Log path $($logPath) not found"
    Exit (1)
}

Start-Transcript -path "$($logPath)\$(get-date -Format yyyy_MM_dd)_Remove-ExchangeOnlineEmails.txt"

$sessions = Get-PSSession | Select-Object -Property State, Name, ComputerName
$exchangeOnlineConnection = (@($sessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*; ComputerName=outlook.office365.com*').Count -gt 0
$complianceConnection = (@($sessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*; ComputerName=*compliance.protection.outlook.com*').Count -gt 0
if (!$exchangeOnlineConnection) {
    Connect-ExchangeOnline
}
if (-not $complianceConnection) {
    Connect-IPPSSession
}

$searchName = "$($target)_messages_older_than_$($months)_months"
try {
    if (Get-ComplianceSearch -Identity $searchName -ErrorAction SilentlyContinue) {
        Write-Host "Compliance Search: $searchName exists. Changing properties"
        Set-ComplianceSearch -Identity $searchName -ExchangeLocation $target -ContentMatchQuery "(Received <= $((get-date).AddMonths(-$months).ToString("MM/dd/yyy")))" -ErrorAction "Stop"
    }
    else {
        Write-Host "Creating Compliance Search: $searchName"
        New-ComplianceSearch -Name $searchName -ExchangeLocation $target -ContentMatchQuery "(Received <= $((get-date).AddMonths(-$months).ToString("MM/dd/yyy")))" -ErrorAction "Stop"
    }
    Write-Host "Running Compliance Search: $searchName"
    Start-ComplianceSearch -Identity $searchName -ErrorAction "Stop"

    While ((Get-ComplianceSearch  -Identity $searchName -ErrorAction "Stop").status -ne "Completed") {
        Write-Host "." -NoNewLine
        Start-Sleep 5
    }
    $complianceSearchResults = Get-ComplianceSearch  -Identity $searchName -ErrorAction "Stop"
    if ($complianceSearchResults.Items -le 0) {
        Write-Host "Compliance Search returned 0 items"
        Stop-Transcript
        Exit
    }
    $searchActionName = "$($searchName)_preview"
    if (Get-ComplianceSearchAction -Identity $searchActionName -ErrorAction SilentlyContinue) {
        Write-Host "Compliance Search Action : $searchName exists. Deleting"
        Remove-ComplianceSearchAction -Identity $searchActionName -Confirm:$false -ErrorAction "Stop"
    }
    if ($preview) {
        Write-Host "Creating Compliance Search Action for Preview: $searchActionName"
        New-ComplianceSearchAction SearchName $searchName -Preview -ErrorAction "Stop"
    }
    else {
        Write-Host "Creating Compliance Search Action for Deletion: $searchActionName"
        New-ComplianceSearchAction -SearchName $searchActionName -Purge -PurgeType SoftDelete -ErrorAction "Stop"
    }
    Write-Host "Waiting for Compliance Search Action to finish"
    While ((Get-ComplianceSearchAction -Identity $searchActionName -ErrorAction "Stop").status -ne "Completed") {
        Write-Host "."
        Start-Sleep 5
    }
    Get-ComplianceSearchAction  -Identity $searchActionName -ErrorAction "Stop" | Format-List -Property Results
}
catch {
    Write-Error "An error occurred: $_"
}
Stop-Transcript