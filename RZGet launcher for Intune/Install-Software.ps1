## ------------------------------------------------------------------
# Arguments
## ------------------------------------------------------------------
[string]$folder = 'C:\temp\InstallSoftware' #Folder to download RZGet.exe and save log.
[string]$RZGetArguments = 'install 7-Zip Notepad++(x64) Edge' # RZget Arguments. Check https://github.com/rzander/ruckzuck/wiki/RZGet and https://ruckzuck.tools/Home/Repository
# Mail Settings. If you do not want to send and email, leave empty $SMTPServer variable: $SMTPServer = ''
[string]$SMTPServer = 'smtp.office365.com'
[string[]]$recipient = 'juangranados@contoso.com', 'support@contoso.com'
[String]$subject = "Installation on computer $env:COMPUTERNAME"
[string]$sender = 'support@contoso.com'
[string]$username = 'support@contoso.com'
[string]$password = 'P@ssw0rd'
[bool]$enableSsl = $true
[int]$port = 587
## ------------------------------------------------------------------
# function Invoke-Process
# https://stackoverflow.com/a/66700583
## ------------------------------------------------------------------
function Invoke-Process {
    param
    (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$FilePath,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$ArgumentList,

        [ValidateSet("Full", "StdOut", "StdErr", "ExitCode", "None")]
        [string]$DisplayLevel
    )

    $ErrorActionPreference = 'Stop'

    try {
        $pinfo = New-Object System.Diagnostics.ProcessStartInfo
        $pinfo.FileName = $FilePath
        $pinfo.RedirectStandardError = $true
        $pinfo.RedirectStandardOutput = $true
        $pinfo.UseShellExecute = $false
        $pinfo.WindowStyle = 'Hidden'
        $pinfo.CreateNoWindow = $true
        $pinfo.Arguments = $ArgumentList
        $p = New-Object System.Diagnostics.Process
        $p.StartInfo = $pinfo
        $p.Start() | Out-Null
        $result = [pscustomobject]@{
            Title     = ($MyInvocation.MyCommand).Name
            Command   = $FilePath
            Arguments = $ArgumentList
            StdOut    = $p.StandardOutput.ReadToEnd()
            StdErr    = $p.StandardError.ReadToEnd()
            ExitCode  = $p.ExitCode
        }
        $p.WaitForExit()

        if (-not([string]::IsNullOrEmpty($DisplayLevel))) {
            switch ($DisplayLevel) {
                "Full" { return $result; break }
                "StdOut" { return $result.StdOut; break }
                "StdErr" { return $result.StdErr; break }
                "ExitCode" { return $result.ExitCode; break }
            }
        }
    }
    catch {
        Write-Host "An error has ocurred"
    }
}
if ($folder.Chars($folder.Length - 1) -eq '\') {
    $folder = ($folder.TrimEnd('\'))
}
if (!(Test-Path $folder)) {
    mkdir $folder
}
$transcriptFile = "$folder\$(get-date -Format yyyy_MM_dd)_InstallSoftware.txt"
Start-Transcript $transcriptFile
Write-Host "Checking for elevated permissions"
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Insufficient permissions to run this script. Execute PowerShell script as an administrator."
    Stop-Transcript
    Exit
}
Write-Host "Script is elevated"
Write-Host "Downloading RZGet.exe"
Invoke-WebRequest "https://github.com/rzander/ruckzuck/releases/latest/download/RZGet.exe" -OutFile "$folder\RZGet.exe"
Write-Host "Running $($folder)\RZGet.exe $($RZGetArguments)"
Invoke-Process -FilePath "$folder\RZGet.exe" -ArgumentList $RZGetArguments -DisplayLevel StdOut
Stop-Transcript

if (!([string]::IsNullOrEmpty($SMTPServer))) {
    $msg = new-object Net.Mail.MailMessage
    $smtp = new-object Net.Mail.SmtpClient($SMTPServer, $port)
    $msg.From = $sender
    $msg.ReplyTo = $sender
    ForEach ($mail in $recipient) {
        $msg.To.Add($mail)
    }
    if ($username -ne "None" -and $password -ne "None") {
        $smtp.Credentials = new-object System.Net.NetworkCredential($username, $password)
    }
    $smtp.EnableSsl = $enableSsl
    $msg.subject = $subject
    $msg.body = Get-Content $transcriptFile -Raw
    try {
        Write-Output "Sending email"
        $smtp.Send($msg)
    }
    catch {
        Write-Host "Error sending email" -ForegroundColor Red
        write-host "Caught an exception:" -ForegroundColor Red
        write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
        write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red 
    }
}