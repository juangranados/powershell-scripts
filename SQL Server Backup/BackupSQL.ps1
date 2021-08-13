<#
.SYNOPSIS
    Full and Log Backup of SQL Server instance databases with SMO 
    Tested on SQL Server 2008, 2008 R2 and 2012
.DESCRIPTION
    This script perform a instance databases backup of a SQL Server in to a zip file.
    Usage: SQLBackup.ps1 [-BackupDirectory <string>] [-DataBases <string[]>] [-Instance <string>] [-SimpleBackup <True | False>] [-RetainDays <int>] [TempDirectory <string>] [SMTPServer <string>] [Recipient <string[]>] [Sender <string>] [Username <string>] [Password <string>] [-SSL <True | False>] [Port <int>] [WriteEvent <True | False>] 

    Requires Powershell comunity extensions: https://pscx.codeplex.com/
    Install it and copy Pscx forder from C:\Program Files (x86)\PowerShell Community Extensions\Pscx3) 
    in C:\Windows\System32\WindowsPowerShell\v1.0\Modules
.PARAMETER BackupDirectory
    Directory where zip file will be saved. UNC paths are supported. 
    Default "...\My Documents\SQLBackup\"
.PARAMETER DataBases
    Array of databases to backup. 
    If empty all instance databases will be saved. 
.PARAMETER Instance
    Instance Name. 
    Default: default instance.
.PARAMETER SimpleBackup
    True (default): Copy only database files
    False: Copy database and logs
.PARAMETER RetainDays
    Days to keep backups in BackupDirectory. Backups prior to this number of days will be deleted.
    Default: keeps all backups  
.PARAMETER TempDirectory
    Temporary directory to save backups files in order to make a zip file. 
    Default "C:\temp\SQLBackup\"
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
.PARAMETER WriteEvent
    Writes an event with script result in Windows Application Event Log
    Default: False
.EXAMPLE
    Backup default instance databases to a network share
    C:\PS>.\BackupSQL.ps1 -BackupDirectory \\FS-SERVER01\BackupSQL
.EXAMPLE
    Backup default instance databases to a network share and send an email with result using gmail
    C:\PS>.\BackupSQL.ps1 -BackupDirectory \\FS-SERVER01\BackupSQL -SMTPServer smtp.gmail.com -Recipient jgranados@contoso.com,administrator@contoso.com -Sender backupSQL@gmail.com -Username backupSQL@gmail.com -Password Pa$$W0rd -SSL True -Port 587
.EXAMPLE
    Backup named instance databases to a network share
    C:\PS>.\BackupSQL.ps1 -BackupDirectory \\FS-SERVER01\BackupSQL -Instance SQLSVR01\BKUPEXEC
.EXAMPLE
    Backup default instance databases to a network share, delete from network share files older than a week and
    write result in Windows Application Event
    C:\PS>.\BackupSQL.ps1 -BackupDirectory \\FS-SERVER01\BackupSQL -RetainDays 7 -WriteEvent True
.EXAMPLE
    Backup only specified databases of a named instance
    C:\PS>.\BackupSQL.ps1 -Instance SQLSVR01\BKUPEXEC -DataBases BEDB,msdb,model

.NOTES
    Author: Juan Granados
    Date:   June 2015    
#>
    Param(
        [Parameter(Mandatory=$false,Position=0)] 
        [ValidateNotNullOrEmpty()]
        [string]$BackupDirectory=[environment]::getfolderpath("mydocuments") + "\SQLBackup\",
        [Parameter(Mandatory=$false,Position=1)] 
        [ValidateNotNullOrEmpty()]
        [string[]]$DataBases="all",
        [Parameter(Mandatory=$false,Position=2)] 
        [ValidateNotNullOrEmpty()]
        [string]$Instance=$env:computerName,
        [Parameter(Mandatory=$false,Position=3)] 
        [ValidateSet("True","False")] 
        [string]$SimpleBackup="True",
        [Parameter(Mandatory=$false,Position=4)] 
        [ValidateRange(0,36500)]
        [int]$RetainDays=0,
        [Parameter(Mandatory=$false,Position=5)] 
        [ValidateNotNullOrEmpty()]
        [string]$TempDirectory = "C:\temp\SQLBackup\",
        [Parameter(Mandatory=$false,Position=6)] 
        [ValidateNotNullOrEmpty()]
        [string]$SMTPServer="None",
        [Parameter(Mandatory=$false,Position=7)] 
        [ValidateNotNullOrEmpty()]
        [string[]]$Recipient,
        [Parameter(Mandatory=$false,Position=8)] 
        [ValidateNotNullOrEmpty()]
        [string]$Sender,
        [Parameter(Mandatory=$false,Position=9)] 
        [ValidateNotNullOrEmpty()]
        [string]$Username="None",
        [Parameter(Mandatory=$false,Position=10)] 
        [ValidateNotNullOrEmpty()]
        [string]$Password="None",
        [Parameter(Mandatory=$false,Position=11)] 
        [ValidateSet("True","False")]
        [string[]]$SSL="False",
        [Parameter(Mandatory=$false,Position=12)] 
        [ValidateNotNullOrEmpty()]
        [int]$Port=25,
        [Parameter(Mandatory=$false,Position=13)] 
        [ValidateSet("True","False")]
        [string]$WriteEvent="False"
    )

[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.ConnectionInfo')            
[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.Management.Sdk.Sfc')            
[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO')            

# Required if SQL Server 2008 (SMO 10.0).            
[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMOExtended')            

# Function to find dabatabase in $Databases[] variable
Function Get-Database
{
    Param([string]$Database)
    for($i = 0; $i -le $Databases.length -1; $i++)
    {
        if ($Databases[$i] -eq $Database)
            {$true}
    }
    $false
}

# Function to check for Module Dependencies
Function Get-MyModule
{
    Param([string]$name)
    if(-not(Get-Module -name $name))
    {
    if(Get-Module -ListAvailable |
    Where-Object { $_.name -eq $name })
    {
    Import-Module -Name $name
    $true
    } #end if module available then import
    else { $false } #module not available
    } # end if not module
    else { $true } #module already loaded
} #end function get-MyModule 

# Write-zip needs module https://pscx.codeplex.com/ install and copy Pscx from C:\Program Files (x86)\PowerShell Community Extensions\Pscx3) 
# to C:\Windows\System32\WindowsPowerShell\v1.0\Modules
If(! (Get-MyModule –name “Pscx”)) 
    {“Pscx module is not installed on this system. Please download and install it from https://pscx.codeplex.com/ and copy Pscx from C:\Program Files (x86)\PowerShell Community Extensions\Pscx3) to C:\Windows\System32\WindowsPowerShell\v1.0\Modules”;exit}

# Variable to measure script execution time
$startDTM = (Get-Date)

# Data sumatory
$DataSum = 0

# Timestamp
$timestamp = Get-Date -format yyyy-MM-dd-HH-mm-ss

# Create application event source 
if ($WriteEvent -eq "True")
{
    if (!(Get-Eventlog -LogName "Application" -Source "BackupSQL"))
        {New-Eventlog -LogName "Application" -Source "BackupSQL"}
}

# Check if backup directory exist and try to create if not
New-Item -ErrorAction Ignore -ItemType directory -Path $BackupDirectory

# Test if backup folder exists
if(!(Test-Path -Path $BackupDirectory))
    {
     Write-Output "$BackupDirectory does not exists and can not be created. Script can not continue"
     if ($WriteEvent -eq "True")
     {
        Write-EventLog –LogName Application –Source “BackupSQL” –EntryType Error –EventID 2 –Message "SQL Backup Failed. $BackupDirectory does not exists and can not be created."
     }
     Exit 
    }      

# Check if local directory exist and try to create if not
New-Item -ErrorAction Ignore -ItemType directory -Path $TempDirectory

# Test if local folder exists
if(!(Test-Path -Path $TempDirectory))
    {
     Write-Output "$TempDirectory does not exists and can not be created. Script can not continue"
     if ($WriteEvent -eq "True")
     {
         Write-EventLog –LogName Application –Source “BackupSQL” –EntryType Error –EventID 2 –Message "SQL Backup Failed.$TempDirectory does not exists and can not be created."
     }
     Exit 
    }

# Preparing mail sending  
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
     }
try
{
     Write-Output "Starting backup"
    # SQL Server
    $srv = New-Object Microsoft.SqlServer.Management.Smo.Server $Instance  
    
    #Disable timeouts monitorization
    $srv.ConnectionContext.StatementTimeout = 0

    # Delete temp folder archives         
    Get-ChildItem –Path  $TempDirectory | Remove-Item
    
    # Copy databases            
    foreach ($db in $srv.Databases)            
    {
        If( (($DataBases[0] -eq "all") -or (Get-Database $db.Name)) -and ($db.Name -ne "tempdb") )
            {            
                $timestamp = Get-Date -format yyyy-MM-dd-HH-mm-ss  
                $backup = New-Object ("Microsoft.SqlServer.Management.Smo.Backup")        
                $backup.Action = "Database"           
                $backup.Database = $db.Name
	            Write-Output "Copying $db"          
                $backup.Devices.AddDevice($TempDirectory + $db.Name + "_full_" + $timestamp + ".bak", "File")            
                $backup.BackupSetDescription = "Full backup of " + $db.Name + " " + $timestamp            
                $backup.Incremental = 0            
                # Full backup            
                $backup.SqlBackup($srv)     
                # Log backup if database recovery mode is not simple         
                If (($db.RecoveryModel -ne 3) -and ($SimpleBackup -eq "False"))            
                {                     
                    $backup = New-Object ("Microsoft.SqlServer.Management.Smo.Backup")            
                    $backup.Action = "Log"            
                    $backup.Database = $db.Name            
                    $backup.Devices.AddDevice($TempDirectory + $db.Name + "_log_" + $timestamp + ".trn", "File")            
                    $backup.BackupSetDescription = "Log backup of " + $db.Name + " " + $timestamp            
                    # Truncate log prior to backup            
                    $backup.LogTruncation = "Truncate"
                    # Log backup         
                    $backup.SqlBackup($srv)            
                }            
            }
    }

    # Creation time
    $timestamp = Get-Date -format yyyy-MM-dd-HH-mm-ss

    # Set zip file and delete '\'
    $InstanceName = $Instance -Replace "\\","-"
    $ZipFile = "$TempDirectory$timestamp" + "_" + $InstanceName + "_Backup.zip"
    $ZipFileRemote = "$BackupDirectory$timestamp" + "_" + $InstanceName + "_Backup.zip"

    # Zip compression
    Get-Childitem $TempDirectory -Recurse | Write-Zip -IncludeEmptyDirectories -OutputPath $ZipFile -EntryPathRoot $TempDirectory

    # Copy zip file to backup folder
    Copy-Item -Path $ZipFile -Destination $ZipFileRemote

    # Show information about the size of file copied
    $DataSum = "{0:N3}" -f (((Get-Item $ZipFileRemote).length) / 1MB)
    Write-Output = "Backup $ZipFileRemote stored. $DataSum MB copied"
    
    # Show information about the size of file copied on Windows registry
    if ($WriteEvent -eq "True")
    {
        Write-EventLog –LogName Application –Source “BackupSQL” –EntryType Information –EventID 1 –Message "Backup $ZipFileRemote stored. $DataSum MB copied"
    }

    # Sending email with backup result
    if ($SMTPServer -ne "None")
    {
        #Email subject
        $msg.subject = $Instance + " SQL Backup Success"
        #Email body
        $msg.body = "Backup $ZipFileRemote stored. $DataSum MB copied"
        #Sending email
        try{
            Write-Output "Sending email"
            $smtp.Send($msg)
            Write-EventLog –LogName Application –Source “BackupSQL” –EntryType Information –EventID 4 –Message "Success backup result sent to $Recipient"
           }catch
                {
                    # Write error on application event log
                    if ($WriteEvent -eq "True")
                    {
                        Write-EventLog –LogName Application –Source “BackupSQL” –EntryType Error –EventID 3 –Message "Error sending mail. " + "Exception Type: $($_.Exception.GetType().FullName)" + ". Exception Message: $($_.Exception.Message)"
                    }
                    write-host "Caught an exception:" -ForegroundColor Red
                    write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
                    write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red 
                }
    }
    # Delete temp directory         
    Remove-Item -Recurse -ErrorAction SilentlyContinue -Confirm:$false –Path $TempDirectory
    
    # Delete old files from backup directory
    if ($RetainDays -ne 0)
        {get-childitem $BackupDirectory -recurse | where {$_.lastwritetime -lt (get-date).adddays(-$RetainDays) -and -not $_.psiscontainer} |% {remove-item $_.fullname -force}}
    
    # Get execution time
    $endDTM = (Get-Date)
    
    Write-Output "Execution time: $(($endDTM-$startDTM).hours) hours $(($endDTM-$startDTM).Minutes) minutes $(($endDTM-$startDTM).Seconds) seconds"
}
catch
    {
        write-host "Caught an exception:" -ForegroundColor Red
        write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
        write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
        # Delete temp directory         
        Remove-Item -Recurse -ErrorAction SilentlyContinue -Confirm:$false –Path $TempDirectory
        Write-Host "Script can not continue."
        # Write error on application event log
        if ($WriteEvent -eq "True")
        {
            Write-EventLog –LogName Application –Source “BackupSQL” –EntryType Error –EventID 2 –Message "SQL Backup Failed" + "Exception Type: $($_.Exception.GetType().FullName)" + ". Exception Message: $($_.Exception.Message)"
        }
        # Sending email with backup result
        if ($SMTPServer -ne "None")
        {
            #Email subject
            $msg.subject = $Instance + " SQL Backup Error"
            #Email body
            $msg.body = "SQL Backup Failed" + "Exception Type: $($_.Exception.GetType().FullName)" + ". Exception Message: $($_.Exception.Message)"
            #Sending email
            Write-Output "Sending email"
            try{
            $smtp.Send($msg)
            Write-EventLog –LogName Application –Source “BackupSQL” –EntryType Information –EventID 4 –Message "Error backup result sent to $Recipient"
            }catch
                {
                    # Write error on application event log
                    if ($WriteEvent -eq "True")
                    {
                        Write-EventLog –LogName Application –Source “BackupSQL” –EntryType Error –EventID 3 –Message "Error sending mail. " + "Exception Type: $($_.Exception.GetType().FullName)" + ". Exception Message: $($_.Exception.Message)"
                    }
                    write-host "Caught an exception:" -ForegroundColor Red
                    write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
                    write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
                }
        }
    }