<#PSScriptInfo

.VERSION 1.0

.GUID 5d639597-a35a-4a95-8010-042fb21c343f

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
    Delete Exchange Online Emails Older Than x Months from a mailbox.
    Running user has to be member of eDiscovery Manager group: https://docs.microsoft.com/en-us/microsoft-365/compliance/assign-ediscovery-permissions?view=o365-worldwide
    in order to run New-ComplianceSearchAction.
    Thanks to: https://answers.microsoft.com/en-us/msoffice/forum/all/delete-more-than-10-items-from-a-mailbox-using/f28efa60-3766-4f50-af2d-e1f9be588931
.PARAMETER target
    Mailboxto delete emails from.
.PARAMETER months
    Emails older than this number of months will be deleted
.PARAMETER logPath
    Path where save log file.
    Default: Temp folder
.PARAMETER preview
    Only shows matching emails. It will not delete anything.
.PARAMETER deleteComplianceSearch
    Deletes Compliance Search object after using it.
.EXAMPLE
    Remove-ExchangeOnlineEmails.ps1 -LogPath "C:\temp\Log" -target esmith@contoso.com -months 12 -preview
.EXAMPLE
    Remove-ExchangeOnlineEmails.ps1 -target finance@contoso.com -months 24 -deleteComplianceSearch
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
    [switch]$preview,
    [Parameter()]
    [switch]$deleteComplianceSearch
)

$ErrorActionPreference = "SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Stop"

function Get-StringBetweenTwoStrings([string]$firstString, [string]$secondString, [string]$string) {
    $pattern = "$firstString(.*?)$secondString"
    $result = [regex]::Match($string, $pattern).Groups[1].Value
    return $result
}
function Get-ParsedLog([string]$log) {
    $log = $log -replace '{' -replace '}', ','
    $table = New-Object system.Data.DataTable "DetailedMessageStats"
    $table.columns.add($(New-Object system.Data.DataColumn Location, ([string])))
    $table.columns.add($(New-Object system.Data.DataColumn Sender, ([string])))
    $table.columns.add($(New-Object system.Data.DataColumn Subject, ([string])))
    $table.columns.add($(New-Object system.Data.DataColumn Type, ([string])))
    $table.columns.add($(New-Object system.Data.DataColumn Size, ([int])))
    $table.columns.add($(New-Object system.Data.DataColumn ReceivedTime, ([Datetime])))
    $table.columns.add($(New-Object system.Data.DataColumn DataLink, ([string])))
    ForEach ($line in $($log -split "`r`n")) {
        $row = $table.NewRow()
        $row.Location = Get-StringBetweenTwoStrings "Location: " "; Sender:" $line
        $row.Sender = Get-StringBetweenTwoStrings "Sender: " "; Subject:" $line
        $row.Subject = Get-StringBetweenTwoStrings "Subject: " "; Type:" $line
        $row.Type = Get-StringBetweenTwoStrings "Type: " "; Size:" $line
        $row.Size = Get-StringBetweenTwoStrings "Size: " "; Received Time:" $line
        $row.ReceivedTime = Get-StringBetweenTwoStrings "Received Time: " "; Data Link:" $line
        $row.DataLink = Get-StringBetweenTwoStrings "Data Link: " "," $line
        $table.Rows.Add($row)
    }
    return $table
}

$logPath = $logPath.TrimEnd('\')
if (-not (Test-Path $logPath)) {
    Write-Host "Log path $($logPath) not found"
    Exit (1)
}

Start-Transcript -path "$($logPath)\$(get-date -Format yyyy_MM_dd)_Remove-ExchangeOnlineEmails.txt"
if (-not (Get-InstalledModule -Name  ExchangeOnlineManagement)) {
    Write-Host "ExchangeOnlineManagement module not found! Run: Install-Module -Name ExchangeOnlineManagement"
    Exit
}

$sessions = Get-PSSession | Select-Object -Property State, Name, ComputerName
$exchangeOnlineConnection = (@($sessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*; ComputerName=outlook.office365.com*').Count -gt 0
$complianceConnection = (@($sessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*; ComputerName=*compliance.protection.outlook.com*').Count -gt 0

if (!$exchangeOnlineConnection) {
    Connect-ExchangeOnline
}

if (-not $complianceConnection) {
    Connect-IPPSSession
}

$folderQueries = @()
$folderStatistics = Get-MailboxFolderStatistics $target | where-object { ($_.FolderPath -eq "/Recoverable Items") -or ($_.FolderPath -eq "/Purges") -or ($_.FolderPath -eq "/Versions") -or ($_.FolderPath -eq "/DiscoveryHolds") }
foreach ($folderStatistic in $folderStatistics) {
    $folderId = $folderStatistic.FolderId;
    $folderPath = $folderStatistic.FolderPath;
    $encoding = [System.Text.Encoding]::GetEncoding("us-ascii")
    $nibbler = $encoding.GetBytes("0123456789ABCDEF");
    $folderIdBytes = [Convert]::FromBase64String($folderId);
    $indexIdBytes = New-Object byte[] 48;
    $indexIdIdx = 0;
    $folderIdBytes | Select-Object -skip 23 -First 24 | ForEach-Object { $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -shr 4]; $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -band 0xF] }
    $folderQuery = "folderid:$($encoding.GetString($indexIdBytes))";
    $folderStat = New-Object PSObject
    Add-Member -InputObject $folderStat -MemberType NoteProperty -Name FolderPath -Value $folderPath
    Add-Member -InputObject $folderStat -MemberType NoteProperty -Name FolderQuery -Value $folderQuery
    $folderQueries += $folderStat
}
       
$RecoverableItemsFolder = $folderQueries.folderquery[0]
$PurgesFolder = $folderQueries.folderquery[1]
$VersionsFolder = $folderQueries.folderquery[2]
$DiscoveryHoldsFolder = $folderQueries.folderquery[3]
$searchName = "$($target)_emails_older_than_$($months)_months"
try {
    if (Get-ComplianceSearch -Identity $searchName -ErrorAction SilentlyContinue) {
        Write-Host "Compliance Search $searchName exists. Changing properties"
        Set-ComplianceSearch -Identity $searchName -ExchangeLocation $target -ContentMatchQuery "(Received <= $((get-date).AddMonths(-$months).ToString("MM/dd/yyy"))) AND (kind:email) AND (NOT (($RecoverableItemsFolder) OR ($PurgesFolder) OR ($VersionsFolder) OR ($DiscoveryHoldsFolder)))" -ErrorAction "Stop"
    }
    else {
        Write-Host "Creating Compliance Search $searchName"
        New-ComplianceSearch -Name $searchName -ExchangeLocation $target -ContentMatchQuery "(Received <= $((get-date).AddMonths(-$months).ToString("MM/dd/yyy"))) AND (kind:email) AND (NOT (($RecoverableItemsFolder) OR ($PurgesFolder) OR ($VersionsFolder) OR ($DiscoveryHoldsFolder)))" -ErrorAction "Stop"
    }
    Write-Host "Running Compliance Search $searchName"
    Start-ComplianceSearch -Identity $searchName -ErrorAction "Stop"

    While ((Get-ComplianceSearch  -Identity $searchName -ErrorAction "Stop").status -ne "Completed") {
        Write-Host "." -NoNewLine
        Start-Sleep 5
    }
    Write-Host "."
    $complianceSearchResults = Get-ComplianceSearch  -Identity $searchName -ErrorAction "Stop"
    if ($complianceSearchResults.Items -le 0) {
        Write-Host "Compliance Search returned 0 items"
    }
    else {
        if ($complianceSearchResults.ExchangeLocation.count -ne 1) {
            Write-Host "You have selected a Compliance Search scoped for more than 1 mailbox, please restart and select a search scoped for a single mailbox."
            if ($deleteComplianceSearch) {
                Write-Host "Deleting object $searchName"
                Remove-ComplianceSearch -Identity $searchName -Confirm:$false -ErrorAction "Stop"
            }
            Exit
        }
        Write-Host "Compliance Search returned $($complianceSearchResults.Items) items"
        if ($preview) {
            $searchActionName = "$($searchName)_preview"
            if (Get-ComplianceSearchAction -Identity $searchActionName -ErrorAction SilentlyContinue) {
                Write-Host "Compliance Search Action $searchActionName exists. Deleting"
                Remove-ComplianceSearchAction -Identity $searchActionName -Confirm:$false -ErrorAction "Stop"
            }
            Write-Host "Creating Compliance Search Action for Preview $searchActionName"
            New-ComplianceSearchAction -SearchName $searchName -Preview -ErrorAction "Stop" | Out-Null
            
            Write-Host "Waiting for Compliance Search Action to finish"
            While ((Get-ComplianceSearchAction -Identity $searchActionName -ErrorAction "Stop").status -ne "Completed") {
                Write-Host "." -NoNewline
                Start-Sleep 5
            }
            Write-Host "."
            $complianceSearchActionResult = Get-ParsedLog (Get-ComplianceSearchAction $searchActionName -Details).Results
            $complianceSearchActionResult | Out-GridView -Title "Compliance Search Preview"
        }
        else {
            $searchActionName = "$($searchName)_purge"
            [int]$batches = [math]::floor($complianceSearchResults.Items / 10)
            if (Get-ComplianceSearchAction -Identity $searchActionName -ErrorAction SilentlyContinue) {
                Write-Host "Compliance Search Action $searchActionName exists. Deleting"
                Remove-ComplianceSearchAction -Identity $searchActionName -Confirm:$false -ErrorAction "Stop" | Out-Null
            }
            for ($batch = 1; $batch -le $batches; $batch++) {
                Write-Host "Batch $batch of $batches" -ForegroundColor Cyan
                Write-Host "Creating Compliance Search Action for Deletion $searchActionName"
                $repeat = $true
                $i = 1
                while ($repeat) {
                    try {
                        New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType HardDelete -Confirm:$false -ErrorAction "Stop" | Out-Null
                        $repeat = $false
                    }
                    catch {
                        Write-Host "Error trying to create Compliance Search Action. Waiting 5 seconds until next try. Try $i of 5"
                        Start-Sleep -Seconds 5
                        Remove-ComplianceSearchAction -Identity $searchActionName -Confirm:$false -ErrorAction "SilentlyContinue" | Out-Null
                        if ($i -lt 6) {
                            $i++
                        }
                        else {
                            Write-Host "Cannot create new Compliance Search Action" -ForegroundColor Red
                            if ($deleteComplianceSearch) {
                                Write-Host "Deleting object $searchName"
                                Remove-ComplianceSearch -Identity $searchName -Confirm:$false -ErrorAction "Stop"
                            }
                            Exit
                        }
                    }
                }
                $complianceSearchActionStatus = (Get-ComplianceSearchAction -Identity $searchActionName).status
                Write-Host "Waiting for Compliance Search Action to finish"
                do {
                    Write-Host "." -NoNewline
                    Start-Sleep 5
                    $complianceSearchActionStatus = (Get-ComplianceSearchAction -Identity $searchActionName).status
                } while ($complianceSearchActionStatus -ne "Completed")
                Write-Host "."
                Write-Host "10 items deleted. $($complianceSearchResults.Items - (10 * $batch)) remaining" -ForegroundColor Green
                Write-Host "Deleting Compliance Search Action $searchActionName"
                Remove-ComplianceSearchAction -Identity $searchActionName -Confirm:$false | Out-Null
            }
        }
    }
    
    if ($deleteComplianceSearch) {
        Write-Host "Deleting object $searchName"
        Remove-ComplianceSearch -Identity $searchName -Confirm:$false -ErrorAction "Stop"
    }
}
catch {
    Write-Error "An error occurred: $_"
}
Stop-Transcript
