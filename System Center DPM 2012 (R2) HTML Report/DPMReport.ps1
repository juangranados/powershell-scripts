<#
.SYNOPSIS
    Create and send a protection group backup report of System Center Data Protection Manager (DPM) 2012 (R2) Servers
.DESCRIPTION
    This script Create and send a protection group backup report of System Center Data Protection Manager (DPM) 2012 (R2) Servers based on the recovery points and protection groups
    Usage: DPMProtectionGroupReport.ps1 [-DPMServers <string[]>] [-ProtectionGroups <string[]>] [-MinLastBackupHours <int>] [-MinRecoveryPoints <int>] [ReportLocation <string>] [SMTPServer <string>] [Recipient <string[]>] [Sender <string>] [Username <string>] [Password <string>] [-SSL <True | False>] [Port <int>]
.PARAMETER DPMPServers
    List of DPM servers to get report.
    Default: localhost
.PARAMETER ProtectionGroups
    List of protection groups to get report.
    Default: All (Report will contain data of all protection groups)
.PARAMETER MinLastBackupHours
    If last recovery point is older than MinLastBackupHours it will show with red background.
    Default: 24
.PARAMETER MinRecoveryPoints
    If number of recovery points is less than MinRecoveryPoint it will show with red background.
    Default: 1
.PARAMETER ReportLocation
    Path where report will be saved
    Default: None
.PARAMETER SMTPServer
    Sets smtp server in order to sent an email with backup result.
    Default: None
.PARAMETER Recipient
    Array of emails addresses which will receive the backup result.
    Default: None
.PARAMETER Sender
    Email address which will send the backup result.
    Default: None
.PARAMETER Username
    Username in case of smtp server requires authentication.
    Default: None
.PARAMETER Password
    Password in case of smtp server requires authentication.
    Default: None
.PARAMETER SSL
    Use of SSL in case of smtp server requires SSL.
    Default: False
.PARAMETER Port
    Port to connect to smtp server.
    Default: 25
.EXAMPLE
    DPMReport.ps1 -DPMServers DPM01,DPM02,DPM03 -HtmlReport \\SERVER1\Reports\ -SMTPServer mail.contoso.com -Sender soporte@contoso.com -Recipient jgranados@contoso.com,administrador@contoso.com -Username contoso\jgranados -Password P@ssw0rd
.NOTES 
    Author: Juan Granados 
    Date:   July 2017
#>

Param(
        [Parameter(Mandatory=$false,Position=0)] 
        [ValidateNotNullOrEmpty()]
        [string[]]$DPMServers=$env:computername,
        [Parameter(Mandatory=$false,Position=1)] 
        [ValidateNotNullOrEmpty()]
        [string[]]$ProtectionGroups="All",
        [Parameter(Mandatory=$false,Position=2)] 
        [ValidateNotNullOrEmpty()]
        [int]$MinLastBackupHours=24,
        [Parameter(Mandatory=$false,Position=2)] 
        [ValidateNotNullOrEmpty()]
        [int]$MinRecoveryPoints=1,
        [Parameter(Mandatory=$false,Position=3)]
        [ValidateNotNullOrEmpty()]
        [string]$ReportLocation="None",
        [Parameter(Mandatory=$false,Position=4)] 
        [ValidateNotNullOrEmpty()]
        [string]$SMTPServer="None",
        [Parameter(Mandatory=$false,Position=5)] 
        [ValidateNotNullOrEmpty()]
        [string[]]$Recipient,
        [Parameter(Mandatory=$false,Position=6)] 
        [ValidateNotNullOrEmpty()]
        [string]$Sender,
        [Parameter(Mandatory=$false,Position=7)] 
        [ValidateNotNullOrEmpty()]
        [string]$Username="None",
        [Parameter(Mandatory=$false,Position=8)] 
        [ValidateNotNullOrEmpty()]
        [string]$Password="None",
        [Parameter(Mandatory=$false,Position=9)] 
        [ValidateSet("True","False")]
        [string[]]$SSL="False",
        [Parameter(Mandatory=$false,Position=10)] 
        [ValidateNotNullOrEmpty()]
        [int]$Port=25
    )

$ErrorActionPreference = "silentlycontinue"

if (Get-Module -ListAvailable -Name DataProtectionManager) {
    Import-Module DataProtectionManager
} else {
    Write-Host "Module DataProtectionManager does not exist. Script can not continue"
    Exit
}

$DomainName=((Get-WmiObject Win32_ComputerSystem).Domain).toupper()

 $HTMLFile = @"
<!DOCTYPE html PUBLIC -//W3C//DTD XHTML 1.0 Strict//EN  http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd>
<html xmlns=http://www.w3.org/1999/xhtml>
<head>
<meta http-equiv=content-type content=text/html;charset=utf-8/>
<title>DPM Backup Status</title>
<style>
    TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
    TH{border-width: 2px;padding: 5px;border-style: solid;border-color: black;background-color:#99CCFF}
    TD{border-width: 2px;padding: 5px;border-style: solid;border-color: black;background-color:#E0F5FF}
    td.green{background-color: LightGreen; color: black;}
    td.gray{background-color: gray; color: black;}
    td.silver{background-color: #B9E1F9; color: black;}
    td.fsdata{background-color: #87AFC7; color: black;}
    td.red{background-color: LightCoral; color: black;}
</style>
</head>
<body>
<H1 align="center">DPM Backup Status</H1>
"@

try{
foreach($DPMServer in $DPMServers){
$HTMLFile+=@"
    <table align= center>
    <font size=3 face=arial>
    <tr><th align=center colspan=7 class=gray>$DPMServer</th></tr>
    </font>  
    <font size=2 face=arial 
    <tr><td align=center class=silver>Protection group</td><td align=center class=silver>Computer</td><td align=center class=silver>Object</td><td align=center class=silver>Type</td><td align=center class=silver>Oldest RP</td><td align=center class=silver>Last RP</td><td align=center class=silver>Number of RP</td></tr>
    </font>
    <font size=2 face=arial>
"@

    $PGs = get-protectiongroup $DPMServer | sort-object Name

    foreach($ProtectionGroup in $PGs){
        $ProtectionGroupname = $ProtectionGroup.friendlyname.toupper()
        if(($ProtectionGroups -contains "All") -OR ($ProtectionGroups -contains $ProtectionGroupname)){
            $DataSources = get-datasource $ProtectionGroup | sort-object Computer, Name
            foreach($DataSource in $DataSources){
                $Computer = $DataSource.productionservername.toupper().replace(".$DomainName","")
                for ($i=0;$i -le 5;$i++){
                        if ($Datasource.TotalRecoveryPoints -ne 0) {break}
                        start-sleep -s 1
                    }
                $TotalRP = $Datasource.TotalRecoveryPoints
                if ($Datasource.TotalRecoveryPoints -ne 0){
                    for ($i=0;$i -le 5;$i++){
                        if([datetime]$DataSource.LatestRecoveryPoint -ne "01/01/0001 0:00:00") {break}
                            start-sleep -s 1   
                    }
                    for ($i=0;$i -le 5;$i++){
                        if([datetime]$DataSource.OldestRecoveryPoint -ne "01/01/0001 0:00:00") {break}
                            start-sleep -s 1   
                    }
                }
                $LastRPTime = [datetime]$DataSource.latestrecoverypoint
                $OldestRPTime=[datetime]$Datasource.OldestRecoveryPoint
                $DateDif = $(get-date) - $LastRPTime
                $Type = $Datasource.ObjectType
                $DataSourceName = $DataSource.name
                Write-Host "Computer: $Computer"
                Write-Host "Type: $Type"
                Write-Host "Name: $DataSourceName"
                Write-Host "Total recovery points: $TotalRP"
                Write-Host "Last recovery point: $LastRPTime"
                Write-Host "Oldest recovery point: $OldestRPTime"
                Write-Host "------------------------------------------------------------"

                $HTMLFile +="<tr align=center><td>$ProtectionGroupname</td><td>$Computer</td><td>$DataSourceName</td><td>$Type</td>"

                if($OldestRPTime -eq "01/01/0001 0:00:00"){ 
                    $HTMLFile +="<td class=red>$OldestRPTime</td>"
                }
                else{
                    $HTMLFile +="<td>$OldestRPTime</td>"
                }
                
                if($DateDif.TotalHours -gt $MinLastBackupHours){ 
                    $HTMLFile +="<td class=red>$LastRPTime</td>"
                }
                else{
                    $HTMLFile +="</td><td class=green>$LastRPTime</td>"
                }
                if($TotalRP -lt $MinRecoveryPoints){ 
                    $HTMLFile +="<td class=red>$TotalRP</td></tr>"
                }
                else{
                    $HTMLFile +="<td class=green>$TotalRP</td></tr>"
                }
        }
        }
    }
    $HTMLFile+="</font></table><br><br>"
 } 
  $HTMLFile+="</body>"

  disconnect-dpmserver
 }catch
    {
     write-Host "Error collecting data" -ForegroundColor Red
     write-host "Caught an exception:" -ForegroundColor Red
     write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
     write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
     $HTMLFile="Error collecting data: $($_.Exception.Message)"
    }
If ($ReportLocation -ne "None")
{
 if(!(Test-Path -Path $ReportLocation))
    {
     Write-Output "$ReportLocation does not exists and can not be created."
    }
 else
    {
     $timestamp = Get-Date -format yyyy-MM-dd-HH-mm-ss
     if ($ReportLocation.Substring($ReportLocation.Length-1) -eq "\")
        {$ReportLocation += "$timestamp" + "_DPMBackupReport.html"}
    else {$ReportLocation += "\" + "$timestamp" + "_DPMBackupReport.html"}
    try{
          ConvertTo-HTML -Body $HTMLFile -Title "DPM Backup Status" | Out-File $ReportLocation
    }catch
         {
          Write-Host "Error storing htlm report" -ForegroundColor Red
          write-host "Caught an exception:" -ForegroundColor Red
          write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
          write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red 
         }
    }
}


if($SMTPServer -ne "None")
    {
        $msg = new-object Net.Mail.MailMessage
        $smtp = new-object Net.Mail.SmtpClient($SMTPServer,$Port)
        $msg.From = $Sender
        $msg.ReplyTo = $Sender
        ForEach($mail in $Recipient)
        {
            $msg.To.Add($mail)
        }
        if ($Username -ne "None" -and $Password -ne "None")
            {
                $smtp.Credentials = new-object System.Net.NetworkCredential($Username, $Password)
            }
        if ($SSL -ne "False")
            {
                $smtp.EnableSsl = $true 
            }
        $msg.subject = "DPM Backup Status"
        $msg.body = $HTMLFile
        $msg.IsBodyHtml = $true
        try{
            Write-Output "Sending email"
            $smtp.Send($msg)
           }catch
                {
                    Write-Host "Error sending email" -ForegroundColor Red
                    write-host "Caught an exception:" -ForegroundColor Red
                    write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
                    write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red 
                }
     }