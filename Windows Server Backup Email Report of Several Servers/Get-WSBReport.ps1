<#
.SYNOPSIS
    This script collect information about Windows Server Backup on a list of servers.
    Requires Windows Server Backup Command Line Tools installed on remote servers (Add-WindowsFeature Backup-Tools)
.DESCRIPTION
    This script collect information about Windows Server Backup on a list of servers and show results on console. It has the posibility of generate an send an html report with backup results.
    Usage: Get-WSBReport.ps1 [-Servers <string>] [-HtmlReport <string>] [SMTPServer <string>] [Recipient <string[]>] [Sender <string>] [Username <string>] [Password <string>] [-SSL <True | False>] [Port <int>]
.PARAMETER Servers
   Full path of a file containing the list of servers to check Windows Server Backup Status   
   Default "C:\Scripts\Servers.txt"
.PARAMETER HtmlReport
   Folder to store html report file.
   Default "C:\Scripts\"
.PARAMETER SMTPServer
    Sets smtp server in order to sent an email with backup result.
    Default: None
.PARAMETER Recipient
    List of emails addresses which will receive the backup result separated by commas.
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
    .\Get-WSBReport.ps1 -ServerList C:\Scripts\servers_contoso.txt -HtmlReport \\SERVER1\Reports\ -SMTPServer mail.contoso.com -Sender soporte@contoso.com -Recipient jgranados@contoso.com,administrador@contoso.com -Username contoso\jgranados -Password P@ssw0rd
.NOTES
    Author: Juan Granados
#>
Param(
        [Parameter(Mandatory=$false,Position=0)] 
        [ValidateNotNullOrEmpty()]
        [string]$ServerList="C:\Scripts\Servers.txt",
        [Parameter(Mandatory=$false,Position=1)] 
        [ValidateNotNullOrEmpty()]
        [string]$HtmlReport="C:\Scripts\",
        [Parameter(Mandatory=$false,Position=2)] 
        [ValidateNotNullOrEmpty()]
        [string]$SMTPServer="None",
        [Parameter(Mandatory=$false,Position=3)] 
        [ValidateNotNullOrEmpty()]
        [string[]]$Recipient,
        [Parameter(Mandatory=$false,Position=4)] 
        [ValidateNotNullOrEmpty()]
        [string]$Sender,
        [Parameter(Mandatory=$false,Position=5)] 
        [ValidateNotNullOrEmpty()]
        [string]$Username="None",
        [Parameter(Mandatory=$false,Position=6)] 
        [ValidateNotNullOrEmpty()]
        [string]$Password="None",
        [Parameter(Mandatory=$false,Position=7)] 
        [ValidateSet("True","False")]
        [string[]]$SSL="False",
        [Parameter(Mandatory=$false,Position=8)] 
        [ValidateNotNullOrEmpty()]
        [int]$Port=25
    )

$timestamp = Get-Date -format yyyy-MM-dd-HH

#Check if server list exists
If (!(Test-Path $ServerList)){
    Write-Host "Can not get servers list. Script will not continue" -ForegroundColor Red;Exit}

$servers = @()

Get-Content $ServerList | Foreach-Object {$servers+=$_}

$results = (1..$servers.length)

for ($i=0; $i -lt $servers.length; $i++)
    {
        $ConnectionError=0
	    Write-Host "Getting result from server: " $servers[$i]
        try{
        $Session = New-PSSession -ComputerName $servers[$i]
        $WindowsVersion = Invoke-Command -session $session -ScriptBlock {(Get-WmiObject win32_operatingsystem).version}
        if ($WindowsVersion -match "6.1")
            {$WBSummary = Invoke-Command -session $session -ScriptBlock {add-pssnapin windows.serverbackup;Get-WBSummary}}
        else {$WBSummary = Invoke-Command -session $session -ScriptBlock {Get-WBSummary}}
        Remove-PSSession $Session
            }catch
                {
                    Write-Host "Error connecting remote server"
                    write-host "Caught an exception:" -ForegroundColor Red
                    write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
                    write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red 
                    $ConnectionError=1
                }
             
        #Storing results
        $results[$i] = New-Object Collections.Arraylist
        
        $results[$i].add("<td>" + $servers[$i] + "</td>") > $null

        if ($ConnectionError -eq 1)
            {
             $results[$i].add("<td><b><font color=red>Unknown</font></b></td>") > $null
             $results[$i].add("<td>" + "Unknown" + "</td>") > $null
             $results[$i].add("<td>" + "Error connecting remote server" + "</td>") > $null
             $results[$i].add("<td>" + "Unknown" + "</td>") > $null
             $results[$i].add("<td>" + "Unknown" + "</td>") > $null
            }
        else
            {
             if ($WBSummary.LastBackupResultHR -eq 0) {$results[$i].add("<td><b><font color=green>Success</font></b></td>") > $null;$result="Success"}
             else {$results[$i].add("<td><b><font color=red>Failure</font></b></td>") > $null;$result="Failure"}
        
             $results[$i].add("<td>" + $WBSummary.LastSuccessfulBackupTime + "</td>") > $null
   
             if ([string]::IsNullOrEmpty($WBSummary.DetailedMessage)){$results[$i].add("<td>Success</td>") > $null;$message="Success"}
             else{$results[$i].add("<td>" + $WBSummary.DetailedMessage + "</td>") > $null;$message= $WBSummary.DetailedMessage}

             $results[$i].add("<td>" + $WBSummary.NumberOfVersions + "</td>") > $null

             if ([string]::IsNullOrEmpty($WBSummary.LastBackupTarget)){$results[$i].add("<td>None</td>") > $null}
             else{$results[$i].add("<td>" + $WBSummary.LastBackupTarget + "</td>") > $null}
               
             Write-Host "Last Backup Result: $result"
             Write-Host "Last Successful Backup Time:" $WBSummary.LastSuccessfulBackupTime
             Write-Host "Detailed Message: $message"
             Write-Host "Number of Backups:" $WBSummary.NumberOfVersions
             Write-Host "-----------------------------------------------------------------" 
            }
    }

if ([boolean](get-variable "rows" -ErrorAction SilentlyContinue))
    {Clear-Variable -Name "rows" -Scope Global}

for ($i=0; $i -lt $servers.length; $i++)
    {
        $rows=$rows +"<tr>" + $results[$i] + "</tr>"
    }

$HTMLFile = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="content-type" content="text/html;charset=utf-8"/>
<style>TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
    TH{border-width: 2px;padding: 5px;border-style: solid;border-color: black;background-color:#99CCFF}
    TD{border-width: 2px;padding: 5px;border-style: solid;border-color: black;background-color:#E0F5FF}
</style>
</head>
<body>
<H1>Windows Server Backup Status</H1>
<table>
    <tr>
        <th>Server</th>
        <th>Last Backup Result</th>
        <th>Last Backup Result Time</th>
        <th>Message</th>
        <th>Number of backups</th>
        <th>Last Backup Target</th>
    </tr>
        $rows
</table>
</body>
</html>
"@
if ([boolean](get-variable "ReportPath" -ErrorAction SilentlyContinue))
    {Clear-Variable -Name "ReportPath" -Scope Global}
$ReportPath = $HtmlReport + "$timestamp" + "_WSBReport.html"
try{
ConvertTo-HTML -Body $HTMLFile -title "Windows Server Backup Report" | Out-File $ReportPath
    }catch
         {
          Write-Host "Error storing htlm report" -ForegroundColor Red
          write-host "Caught an exception:" -ForegroundColor Red
          write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
          write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red 
         }
# Mail sending  
if($SMTPServer -ne "None")
    {
        #Creating a Mail object
        $msg = new-object Net.Mail.MailMessage

        #Creating SMTP server object
        $smtp = new-object Net.Mail.SmtpClient($SMTPServer,$Port)

        #Email structure
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
        #Email subject
        $msg.subject = "Windows Server Backup Status"
        #Email body
        $msg.body = $HTMLFile
        $msg.IsBodyHtml = $true
        #Sending email
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
