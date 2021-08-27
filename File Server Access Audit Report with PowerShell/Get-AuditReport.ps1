<#PSScriptInfo

.VERSION 1.0

.GUID 03a6a60d-01b5-49a8-adb9-ca890ea6f2eb

.AUTHOR Juan Granados

.COPYRIGHT 2021 Juan Granados

.TAGS Audit Report HTML csv mail File Access

.LICENSEURI https://raw.githubusercontent.com/juangranados/powershell-scripts/main/LICENSE

.PROJECTURI https://github.com/juangranados/powershell-scripts/tree/main/File%20Server%20Access%20Audit%20Report%20with%20PowerShell

.RELEASENOTES
    Initial release

#>

<#
.DESCRIPTION
	This PowerShell script allows to audit several file servers and send a report in CSV and HTML by mail.
    CSV file can be import on Excel to generate a File Audit Report.
    HTML report can filter and sorting rows by server, time, user, file or operation (read, delete or write).
    Requirements:
        Enable Audit in Windows and target folders, instructions: https://github.com/juangranados/powershell-scripts/tree/main/File%20Server%20Access%20Audit%20Report%20with%20PowerShell
	Author: Juan Granados 
#>

#########################################################################################################################################
#Variables to modify
#########################################################################################################################################
# List of servers to check audit events
$server = "SRVJ-FS01","SVMEN01"
# List of file extensions to ignore
$ignoredExtensions = "tmp","rgt","mta","tlg","nd","ps","log","ldb","crdownload","DS_Store","cdf-ms","ini"
# List of users to ignore. If you do not want to ignore users, leave it empty.
# Example: $skippedUsers = "administrator","audit-test"
$skippedUsers = ""
# List of files to audit if you are interested only in a few files. If empty, all files (except those with ignored extensions) will be included in report.
# For example: $filesToAudit = "MontlyReport.xlsx","Internal Database.mdb" -> Only this two files will be included in report.
$filesToAudit = ""
# Number of hours back in time to check audit events
$hoursBackToCheck = "24"
# Reports path
$timestamp = Get-Date -format yyyy-MM-dd_HH-mm-ss # Timestamp to add to report name
$htmlReportPath = "$PSScriptRoot\" + "$timestamp" + "_AuditReport.html" # default: html report path in script path.
$csvReportPath = "$PSScriptRoot\" + "$timestamp" + "_AuditReport.csv" # default: csv report path in script path.
$transcriptPath = "$PSScriptRoot\" + "$timestamp" + "_AuditReport.log" # default: PowerShell transcript in script path.
# Mail settings
[string]$SMTPServer="mail.contoso.com" # If "None", no mail is sending. Example: [string]$SMTPServer="mail.contoso.com"
[string[]]$Recipient="jgranados@contoso.com" # List of recipients. Example: [string[]]$Recipient="jdoe@contoso.com","fsmith@contoso.com"
[string]$Sender = "audit-reports@contoso.com" # Sender. Example: [string]$Sender="reports@contoso.com"
[string]$Username="audit-reports@contoso.com" # User name to authenticate with mail server. If "None", no auth is performed. Example: [string]$Username="jdoe@gmail.com"
[string]$Password="P@ssw0rd" # Password to to authenticate with mail server. If "None", no auth is performed. Example: [string]$Password="P@ssw0rd"
[string]$SSL="True" # Using TLS/SSL to authenticate. Example: [string]$SSL="True" (Is required for Gmail or Office365)
[int]$Port=25 # Port of mail server. Example: [int]$Port=587
#########################################################################################################################################
#Internal variables. Do not modify.
#########################################################################################################################################
$ErrorActionPreference = "Stop"
$startDate = (get-date).AddHours(-$hoursBackToCheck)
$ns = @{e = "http://schemas.microsoft.com/win/2004/08/events/event"}
$htmlEvents = [System.Collections.ArrayList]@()
$csvContents = @()
$accessMasks = [ordered]@{
  '0x80' = 'Read'
  '0x2' = 'Write'
  '0x10000' = 'Delete'
}
$previousTimeCreated = ""
$previousSubjectUserName = ""
$previousObjectName = ""
$previousAccessMask = ""
$lastEventTimeCreated = ""
$evts = $null
#########################################################################################################################################
#Functions. Do not modify.
#########################################################################################################################################
Function getFileExtension($path) {
    $file = (Split-Path -Path $path -Leaf).Split(".")
    return $file[$file.Length-1]
}
# This function checks to see if the file should be ignored. 
Function isTempFile($path) {
    $fileName = (Split-Path -Path $path -Leaf)
    $fileExtension = getFileExtension $path
    
    If ($fileName.substring(0,1) -eq "~" -or $fileName -eq "thumbs.db" -or $fileName.Substring(0,1) -eq "$") {
        return $true
    }
    
    ForEach($extension in $ignoredExtensions) {
        if ($fileExtension -eq $extension) {
            return $true
        }
    }
    
    return $false
}
Function isUserToSkip($user) {
    foreach ($svr in $server) {
        $serverUser = $svr + '$'
        if ($user -eq $serverUser) {
            return $true
        }
    }
    foreach ($skippedUser in $skippedUsers) {
        if ($user -eq $skippedUser) {
            return $true
        }
    }
    return $false
}
Function isFileToAudit($path) {
    if ($filesToAudit -eq "") {
     return $true
    } else {
        $fileName = (Split-Path -Path $path -Leaf)
        foreach ($file in $filesToAudit) {
            if ($file -eq $fileName) {
                return $true
            }    
        }
        return $false
    }
}
Function isFile($path) {
    try {
        if ((Get-Item $path) -is [System.IO.FileInfo]) {
            return $true
        }
    } catch {
        if ((Split-Path -Path $path -Leaf) -like "*.*") {
            return $true
        }
    }
}
function checkIfAddEvent($serverName, $timeCreated, $userName, $objectName, [string]$accessMask) {
    if (-not [string]::IsNullOrEmpty($accessMask)){
        if ($script:previousTimeCreated) {
            if ($timeCreated -gt $script:previousTimeCreated.AddSeconds(1) -or $userName -ne $script:previousSubjectUserName -or $objectName -ne $script:previousObjectName) {
                Write-Host "Adding event: $serverName | $script:previousTimeCreated | $script:previousSubjectUserName | $script:previousObjectName | $script:previousAccessMask" -ForegroundColor Green
                # Add event to html
                $htmlEvents.add("<tr>`n") | Out-Null
                $htmlEvents.add("  <td>$($serverName)</td>`n") | Out-Null # Time of access           
                $htmlEvents.add("  <td>$($script:previousTimeCreated.ToString("yyyy-MM-ddTHH:mm:ss"))</td>`n") | Out-Null # Time of access           
                $htmlEvents.add("  <td>$($script:previousSubjectUserName)</td>`n") | Out-Null # User           
                $htmlEvents.add("  <td>$($script:previousObjectName)</td>`n") | Out-Null # File           
                $htmlEvents.add("  <td>$($script:previousAccessMask)</td>`n") | Out-Null # Action            
                $htmlEvents.add("</tr>`n") | Out-Null
                # Add event to csv
                $row = New-Object System.Object
                $row | Add-Member -MemberType NoteProperty -Name "Server" -Value $serverName
                $row | Add-Member -MemberType NoteProperty -Name "Time" -Value $script:previousTimeCreated.ToString("yyyy-MM-ddTHH:mm:ss")
                $row | Add-Member -MemberType NoteProperty -Name "User" -Value $script:previousSubjectUserName
                $row | Add-Member -MemberType NoteProperty -Name "File" -Value $script:previousObjectName
                $row | Add-Member -MemberType NoteProperty -Name "Action" -Value $script:previousAccessMask
                $script:csvContents += $row

                Write-Host "Analizing event: $timeCreated | $userName | $objectName | $accessMask" -ForegroundColor Cyan
                $script:lastEventTimeCreated = $script:previousTimeCreated
                $script:previousTimeCreated = $timeCreated
                $script:previousSubjectUserName = $userName
                $script:previousObjectName = $objectName
                $script:previousAccessMask = $accessMask
            } else {
                Write-Host "Analizing event: $timeCreated | $userName | $objectName | $accessMask" -ForegroundColor Cyan
                if ($script:previousAccessMask -ne "Write") {
                    $script:previousAccessMask = $accessMask
                }
            }
        } else {
            Write-Host "Analizing event: $timeCreated | $userName | $objectName | $accessMask" -ForegroundColor Cyan
            $script:previousTimeCreated = $timeCreated
            $script:previousSubjectUserName = $userName
            $script:previousObjectName = $objectName
            $script:previousAccessMask = $accessMask
        }
    } else {
        Write-Host "Audit event $accessMask not included in 'Read', 'Modified' or 'Deleted'" -ForegroundColor DarkRed
    }
    
}
#########################################################################################################################################
# Main. Do not modify.
#########################################################################################################################################
cls
Start-Transcript $transcriptPath
foreach ($svr in $server) {
    Write-Output "Getting events with ID 4663 of $svr since $startDate. This may take several minutes depending of log size..."
    $evts = Get-WinEvent -computer $svr -FilterHashtable @{LogName="security";ProviderName="Microsoft-Windows-Security-Auditing";ID="4663";StartTime=$startDate} -oldest
    foreach($evt in $evts) {
        $xml = [xml]$evt.ToXML()
        $SubjectUserName = Select-Xml -Xml $xml -Namespace $ns -XPath "//e:Data[@Name='SubjectUserName']/text()" | Select-Object -ExpandProperty Node | Select-Object -ExpandProperty Value
        $ObjectName = Select-Xml -Xml $xml -Namespace $ns -XPath "//e:Data[@Name='ObjectName']/text()" | Select-Object -ExpandProperty Node | Select-Object -ExpandProperty Value
        $AccessMask = Select-Xml -Xml $xml -Namespace $ns -XPath "//e:Data[@Name='AccessMask']/text()" | Select-Object -ExpandProperty Node | Select-Object -ExpandProperty Value

        if ($evt.TimeCreated -ge $startDate){
            if (isFile $ObjectName) {
                if (-not (isTempFile $ObjectName) -and -not (isUserToSkip $SubjectUserName) -and (isFileToAudit $ObjectName)) {
                    checkIfAddEvent $svr $evt.TimeCreated $SubjectUserName $ObjectName $accessMasks[$AccessMask.ToString()]
                }
            }
        }    
    }
    # Last event in case of has not been added
    if ($lastEventTimeCreated -ne $previousTimeCreated -and $previousTimeCreated -ne $null -and !([string]::IsNullOrWhiteSpace($previousTimeCreated))) {
        Write-Host "Adding event : $previousTimeCreated | $previousSubjectUserName | $previousObjectName | $previousAccessMask" -ForegroundColor Green
        $htmlEvents.add("<tr>`n") | Out-Null
        $htmlEvents.add("   <td>$($svr)</td>`n") | Out-Null
        $htmlEvents.add("   <td>$($previousTimeCreated.ToString("yyyy-MM-ddTHH:mm:ss"))</td>`n") | Out-Null # Time of access
        $htmlEvents.add("   <td>$($previousSubjectUserName)</td>`n") | Out-Null # User
        $htmlEvents.add("   <td>$($previousObjectName)</td>`n") | Out-Null # File or folder
        $htmlEvents.add("   <td>$($previousAccessMask)</td>`n") | Out-Null # Action
        $htmlEvents.add("</tr>`n") | Out-Null
        # Add event to csv
        $row = New-Object System.Object
        $row | Add-Member -MemberType NoteProperty -Name "Server" -Value $svr
        $row | Add-Member -MemberType NoteProperty -Name "Time" -Value $previousTimeCreated.ToString("yyyy-MM-ddTHH:mm:ss")
        $row | Add-Member -MemberType NoteProperty -Name "User" -Value $previousSubjectUserName
        $row | Add-Member -MemberType NoteProperty -Name "File" -Value $previousObjectName
        $row | Add-Member -MemberType NoteProperty -Name "Action" -Value $previousAccessMask
        $csvContents += $row
    }
    $previousTimeCreated = ""
    $previousSubjectUserName = ""
    $previousObjectName = ""
    $previousAccessMask = ""
}
$head = @"
<meta http-equiv="content-type" content="text/html;charset=utf-8"/>
<script>
function filterRows() {
    document.getElementById("spinner").className='loading';
    setTimeout(function() {
        var input0,input1, input2, input3, input4,
        filter0, filter1, filter2, filter3, filter4,
        table, tr, td0, td1, td2, td3, td4, i, 
        txtValue0, txtValue1, txtValue2, txtValue3, txtValue4;

        input0 = document.getElementById("myInput0");
        input1 = document.getElementById("myInput1");
        input2 = document.getElementById("myInput2");
        input3 = document.getElementById("myInput3");
        input4 = document.getElementById("myInput4");
        filter0 = input0.value.toUpperCase();
        filter1 = input1.value.toUpperCase();
        filter2 = input2.value.toUpperCase();
        filter3 = input3.value.toUpperCase();
        filter4 = input4.value.toUpperCase();
        table = document.getElementById("myTable");
        tr = table.getElementsByTagName("tr");

        for (i = 1; i < tr.length; i++) {
            td0 = tr[i].getElementsByTagName("td")[0];
            td1 = tr[i].getElementsByTagName("td")[1];
            td2 = tr[i].getElementsByTagName("td")[2];
            td3 = tr[i].getElementsByTagName("td")[3];
            td4 = tr[i].getElementsByTagName("td")[4];
            if (td0) {
                txtValue0 = td0.textContent || td0.innerText;
            } else {
		        txtValue0="";
	        }
            if (td1) {
                txtValue1 = td1.textContent || td1.innerText;
            } else {
		        txtValue1="";
	        }
	        if (td2) {
                txtValue2 = td2.textContent || td2.innerText;
            } else {
		        txtValue2="";
	        }
	        if (td3) {
                txtValue3 = td3.textContent || td3.innerText;
            } else {
		        txtValue3="";
	        }
	        if (td4) {
                txtValue4 = td4.textContent || td4.innerText;
            } else {
		        txtValue4="";
	        }
            if (txtValue0.toUpperCase().indexOf(filter0) > -1 &&
                txtValue1.toUpperCase().indexOf(filter1) > -1 &&
                txtValue2.toUpperCase().indexOf(filter2) > -1 && 
                txtValue3.toUpperCase().indexOf(filter3) > -1 && 
                txtValue4.toUpperCase().indexOf(filter4) > -1 
            ) {
                tr[i].style.display = "";
                } else {
                tr[i].style.display = "none";
                }
            }
            document.getElementById("spinner").className='hidden';
        }, 100);
}
const getCellValue = (tr, idx) => tr.children[idx].innerText || tr.children[idx].textContent;

const comparer = (idx, asc) => (a, b) => ((v1, v2) => 
    v1 !== '' && v2 !== '' && !isNaN(v1) && !isNaN(v2) ? v1 - v2 : v1.toString().localeCompare(v2)
    )(getCellValue(asc ? a : b, idx), getCellValue(asc ? b : a, idx));

window.onload=function(){
document.querySelectorAll('th').forEach(th => th.addEventListener('click', (() => {
    document.getElementById("spinner").className='loading';
    const table = th.closest('table');
    setTimeout(function() {
      Array.from(table.querySelectorAll('tr:nth-child(n+2)'))
        .sort(comparer(Array.from(th.parentNode.children).indexOf(th), this.asc = !this.asc))
        .forEach(tr => table.appendChild(tr) );
      document.getElementById("spinner").className='hidden';
	}, 100);  
})));
}
</script>
<style>
body{
	font-family:Verdana, Arial, sans-serif;
}
.myInput {
  width: 20%;
  font-size: 16px;
  padding: 12px 20px 12px 40px;
  border: 1px solid #ddd;
  margin-bottom: 12px;
}


#myTable {
  border-collapse: collapse;
  width: 100%;
  border: 1px solid #ddd;
  font-size: 18px;
}

#myTable th, #myTable td {
  text-align: left;
  padding: 12px;
}

#myTable tr {
  border-bottom: 1px solid #ddd;
}

#myTable tr.header, #myTable tr:hover {
  background-color: #f1f1f1;
  cursor: pointer;
}
th b.sort-by { 
	padding-right: 18px;
	position: relative;
}
b.sort-by:before,
b.sort-by:after {
	border: 4px solid transparent;
	content: "";
	display: block;
	height: 0;
	right: 5px;
	top: 50%;
	position: absolute;
	width: 0;
}
b.sort-by:before {
	border-bottom-color: #666;
	margin-top: -9px;
}
b.sort-by:after {
	border-top-color: #666;
	margin-top: 1px;
}

.loading {
  position: fixed;
  z-index: 999;
  height: 2em;
  width: 2em;
  overflow: visible;
  margin: auto;
  top: 0;
  left: 0;
  bottom: 0;
  right: 0;
}

.loading:before {
  content: '';
  display: block;
  position: fixed;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  background-color: rgba(0,0,0,0.3);
}

.loading:not(:required) {
  /* hide "loading..." text */
  font: 0/0 a;
  color: transparent;
  text-shadow: none;
  background-color: transparent;
  border: 0;
}

.loading:not(:required):after {
  content: '';
  display: block;
  font-size: 10px;
  width: 1em;
  height: 1em;
  margin-top: -0.5em;
  -webkit-animation: spinner 1500ms infinite linear;
  -moz-animation: spinner 1500ms infinite linear;
  -ms-animation: spinner 1500ms infinite linear;
  -o-animation: spinner 1500ms infinite linear;
  animation: spinner 1500ms infinite linear;
  border-radius: 0.5em;
  -webkit-box-shadow: rgba(0, 0, 0, 0.75) 1.5em 0 0 0, rgba(0, 0, 0, 0.75) 1.1em 1.1em 0 0, rgba(0, 0, 0, 0.75) 0 1.5em 0 0, rgba(0, 0, 0, 0.75) -1.1em 1.1em 0 0, rgba(0, 0, 0, 0.5) -1.5em 0 0 0, rgba(0, 0, 0, 0.5) -1.1em -1.1em 0 0, rgba(0, 0, 0, 0.75) 0 -1.5em 0 0, rgba(0, 0, 0, 0.75) 1.1em -1.1em 0 0;
  box-shadow: rgba(0, 0, 0, 0.75) 1.5em 0 0 0, rgba(0, 0, 0, 0.75) 1.1em 1.1em 0 0, rgba(0, 0, 0, 0.75) 0 1.5em 0 0, rgba(0, 0, 0, 0.75) -1.1em 1.1em 0 0, rgba(0, 0, 0, 0.75) -1.5em 0 0 0, rgba(0, 0, 0, 0.75) -1.1em -1.1em 0 0, rgba(0, 0, 0, 0.75) 0 -1.5em 0 0, rgba(0, 0, 0, 0.75) 1.1em -1.1em 0 0;
}


@-webkit-keyframes spinner {
  0% {
    -webkit-transform: rotate(0deg);
    -moz-transform: rotate(0deg);
    -ms-transform: rotate(0deg);
    -o-transform: rotate(0deg);
    transform: rotate(0deg);
  }
  100% {
    -webkit-transform: rotate(360deg);
    -moz-transform: rotate(360deg);
    -ms-transform: rotate(360deg);
    -o-transform: rotate(360deg);
    transform: rotate(360deg);
  }
}
@-moz-keyframes spinner {
  0% {
    -webkit-transform: rotate(0deg);
    -moz-transform: rotate(0deg);
    -ms-transform: rotate(0deg);
    -o-transform: rotate(0deg);
    transform: rotate(0deg);
  }
  100% {
    -webkit-transform: rotate(360deg);
    -moz-transform: rotate(360deg);
    -ms-transform: rotate(360deg);
    -o-transform: rotate(360deg);
    transform: rotate(360deg);
  }
}
@-o-keyframes spinner {
  0% {
    -webkit-transform: rotate(0deg);
    -moz-transform: rotate(0deg);
    -ms-transform: rotate(0deg);
    -o-transform: rotate(0deg);
    transform: rotate(0deg);
  }
  100% {
    -webkit-transform: rotate(360deg);
    -moz-transform: rotate(360deg);
    -ms-transform: rotate(360deg);
    -o-transform: rotate(360deg);
    transform: rotate(360deg);
  }
}
@keyframes spinner {
  0% {
    -webkit-transform: rotate(0deg);
    -moz-transform: rotate(0deg);
    -ms-transform: rotate(0deg);
    -o-transform: rotate(0deg);
    transform: rotate(0deg);
  }
  100% {
    -webkit-transform: rotate(360deg);
    -moz-transform: rotate(360deg);
    -ms-transform: rotate(360deg);
    -o-transform: rotate(360deg);
    transform: rotate(360deg);
  }
}

.hidden {display:none;}
.visible{display:block;}
</style>
"@
$body = @"
<H1>File Audit Report</H1>
<p>Click on a row to short by. Click again to reverse shorting.</p>
<p>Start typing on search box to filter by server, time, user, file or operation. You can filter by multiple fields.</p>
<input class="myInput" type="text" id="myInput0" onkeyup="filterRows()" placeholder="Search for server..">
<input class="myInput" type="text" id="myInput1" onkeyup="filterRows()" placeholder="Search for time..">
<input class="myInput" type="text" id="myInput2" onkeyup="filterRows()" placeholder="Search for user..">
<input class="myInput" type="text" id="myInput3" onkeyup="filterRows()" placeholder="Search for file..">
<input class="myInput" type="text" id="myInput4" onkeyup="filterRows()" placeholder="Search for operation..">
<div class="hidden" id="spinner">
  <div class="rect1"></div>
  <div class="rect2"></div>
  <div class="rect3"></div>
  <div class="rect4"></div>
  <div class="rect5"></div>
</div>
<table id="myTable">
    <tr class="header">
        <th style="width:20%;"><b class="sort-by">Server</th>
        <th style="width:20%;"><b class="sort-by">Time</th>
        <th style="width:20%;"><b class="sort-by">User</th>
        <th style="width:20%;"><b class="sort-by">File</th>
        <th style="width:20%;"><b class="sort-by">Operation</th>
    </tr>
        $htmlEvents
</table>
"@
if ($htmlEvents.Count -eq 0) {
    Write-Host "There is not any audit events to report. Check audit configuration." -ForegroundColor Red
    Stop-Transcript
    Exit(1)
}
$title = "File Audit Report"
$timestamp = Get-Date -format yyyy-MM-dd_HH-mm-ss
$htmlReportPath = "$PSScriptRoot\" + "$timestamp" + "_AuditReport.html"
$csvReportPath = "$PSScriptRoot\" + "$timestamp" + "_AuditReport.csv"
Write-Host "Creating CSV file: $csvReportPath" -ForegroundColor Yellow
$csvContents | Export-CSV -Path $csvReportPath -Encoding UTF8 -NoTypeInformation
try{
Write-Host "Creating HTML file: $htmlReportPath" -ForegroundColor Yellow
ConvertTo-HTML -Title $title -Head $Head -Body $Body | Out-File $htmlReportPath
    }catch
         {
          Write-Host "Error storing htlm report" -ForegroundColor Red
          write-host "Caught an exception:" -ForegroundColor Red
          write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
          write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red 
         }
# Mail sending  
if($SMTPServer -ne "None"){
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
    $msg.subject = "Audit Report"
    #Email body
    $msg.body = "File Audit Reports Attached."
    $msg.IsBodyHtml = $true
    $msg.Attachments.Add($csvReportPath)
    $msg.Attachments.Add($htmlReportPath)
    #Sending email
    try{
        Write-Host "Sending email" -ForegroundColor Yellow
        $smtp.Send($msg)
        Write-Host "Mail sending ok. End of script" -ForegroundColor Green
        }catch
            {
                Write-Host "Error sending email" -ForegroundColor Red
                write-host "Caught an exception:" -ForegroundColor Red
                write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
                write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red 
            }
}
Stop-Transcript