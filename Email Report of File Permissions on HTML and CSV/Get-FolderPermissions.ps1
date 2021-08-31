<#PSScriptInfo

.VERSION 1.0

.GUID 1069276e-50b4-414a-ae8c-b8801445ae7e

.AUTHOR Juan Granados

.COPYRIGHT 2021 Juan Granados

.TAGS Folder Permission Report HTML CSV email mail

.LICENSEURI https://raw.githubusercontent.com/juangranados/powershell-scripts/main/LICENSE

.PROJECTURI https://github.com/juangranados/powershell-scripts/tree/main/Email%20Report%20of%20File%20Permissions%20on%20HTML%20and%20CSV

.RELEASENOTES
    Initial release
#>

<#
.SYNOPSIS
    Generate a folders permissions report.
.DESCRIPTION
    Starting with a root folder, it generates a folders permissions report. Number of subfolders examined depends on FolderDeep parameter.
    Report is generated in CSV format and can be send attached via mail with a html report in the body. 
.PARAMETER OutFile
    Path to store CSV file.
    Default .\Permissions.csv
.PARAMETER RootPath
    Folder to start checking permissions.
.PARAMETER FolderDeep
    Number of subfolders levels to check.
    Default 99.
.PARAMETER ObjectsIgnored
    Users or groups to ignore in report.
    Default NT AUTHORITY\SYSTEM,BUILTIN\Administrator
.PARAMETER InspectGroups
    List only users in report.
    Default $False
.PARAMETER SMTPServer
    Sets smtp server in order to sent an email with backup result. If leave blank, no email will be send.
.PARAMETER SMTPRecipient
    List of emails addresses which will receive the backup result separated by commas.
.PARAMETER SMTPSender
    Email address which will send the backup result.
.PARAMETER SMTPUser
    Username in case of smtp server requires authentication.
.PARAMETER SMTPPassword
    Password in case of smtp server requires authentication.
.PARAMETER SMTPSSL
    Use of SSL in case of smtp server requires SSL.
    Default: $False
.PARAMETER SMTPPort
    Port to connect to smtp server.
    Default: 25
.EXAMPLE
    Get-FoldersPermissions -RootPath "D:\Data\Departments" -FolderDeep 2 -SMTPServer "mail.server.com" -SMTPRecipient "megaboss@server.com","support@server.com" -SMTPSender "reports@server.com"
.LINK
    https://github.com/juangranados/powershell-scripts/tree/main/Email%20Report%20of%20File%20Permissions%20on%20HTML%20and%20CSV
.NOTES 
    Author: Juan Granados 
#>
[cmdletbinding()]

param(
    [parameter(Position=0,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Path to store CSV file')][string]$OutFile = ".\$(Get-Date -format "yyyyMMdd_hhmmss")-Permissions.csv",
    [parameter(Position=1,Mandatory=$true,ValueFromPipeline=$false,HelpMessage='Folder to start checking permissions')][string]$RootPath,
    [parameter(Position=2,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Number of subfolders levels to check')][string]$FolderDeep = 99,
    [parameter(Position=3,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Users or groups to ignore in report')][string[]]$ObjectsIgnored = @("NT AUTHORITY\SYSTEM","BUILTIN\Administrator"),
    [parameter(Position=4,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Inspect users in groups ($True/$False)')][bool]$InspectGroups=$false,
    [parameter(Position=5,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Mail From')][string]$SMTPSender,
    [parameter(Position=6,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Mail To')]$SMTPRecipient,
    [parameter(Position=7,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Mail Server')][string]$SMTPServer,
    [parameter(Position=8,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Mail User')][string]$SMTPUser,
    [parameter(Position=9,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Mail Password')][string]$SMTPPassword,
    [parameter(Position=10,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Mail Port')][string]$SMTPPort=25,
    [parameter(Position=11,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Use SSL in mail sending ($True/$False)')][bool]$SMTPSSL=$False,
    [parameter(Position=12,Mandatory=$false,ValueFromPipeline=$false,HelpMessage='Mail Subject')][string]$SMTPSubject="Permission report on server $($env:computername) on directory $($RootPath) with $($FolderDeep) level deep"
)
Function Get-FolderPermissions($Folder,[int]$Deep = 0){
    # Write current folder name 
    Write-Host "Examining folder $($Folder.FullName)"
    
    # Get folder ACLs
    $ACLs = get-acl $Folder.fullname | ForEach-Object {$_.Access}
	# Examining folder ACLs    
    Foreach ($ACL in $ACLs){
        # If current ACL contains one of object ignored list, skip ACL
        if (-not ($ObjectsIgnored.Contains($ACL.IdentityReference.value.toString()))){
            # Delete commas in folder name and ACL InheritanceFlags, FileSystemRights
            $FolderName = $($Folder.Fullname -replace ',','.')
            $ACLInheritanceFlags = $($ACL.InheritanceFlags -replace ',',' &')
            $ACLFileSystemRights = $($ACL.FileSystemRights -replace ',',' &')
            # If inspect groups is true, list group users.
            if ($InspectGroups){
                # Check if ACL identity reference is a group
                if (-not ($ACL.IdentityReference.value -like "BUILTIN\*") -and ($ACL.IdentityReference.Value.LastIndexOf('\') -ne -1)){
                    # Get group users
                    $ADGroup = Get-ADGroup -LDAPFilter "(SAMAccountName=$($ACL.IdentityReference.Value.Substring($ACL.IdentityReference.Value.LastIndexOf('\')+1)))" -ErrorAction SilentlyContinue
                    # If group has users
                    if ($ADGroup){
                        # Get group users
                        $users = Get-ADGroupMember -identity $ADGroup -Recursive | Get-ADUser -Property DisplayName
                        # Store users ACL information
                        ForEach($User in $users){
                            # Store user info in csv file
                            $OutInfo = $FolderName + "," + $User.UserPrincipalName  + "," + $ACL.AccessControlType + "," + $ACLFileSystemRights + "," + $ACL.IsInherited + "," + $ACLInheritanceFlags + "," + $ACL.PropagationFlags
                            Add-Content -Value $OutInfo -Path $OutFile
                            # Store user info in table
                            $Rows+="<tr><td>" + $Folder.Fullname + "</td>" + "<td>" + $User.UserPrincipalName + "</td>" + "<td>" + $ACL.AccessControlType + "</td>" + "<td>" + $ACL.FileSystemRights + "</td>" + "<td>" + $ACL.IsInherited + "</td>" + "<td>" + $ACL.InheritanceFlags + "</td>" + "<td>" + $ACL.PropagationFlags + "</td></tr>`r"
                        }
                        # Next loop
                        continue
                    }
                }
            }
        # Inspect groups is false or ACL identity reference is a group but users could not be retrieved
        # Store object info in csv file
        $OutInfo = $FolderName + "," + $ACL.IdentityReference  + "," + $ACL.AccessControlType + "," + $ACLFileSystemRights + "," + $ACL.IsInherited + "," + $ACLInheritanceFlags + "," + $ACL.PropagationFlags
	    Add-Content -Value $OutInfo -Path $OutFile
        # Store object info in table
        $Rows+="<tr><td>" + $Folder.Fullname + "</td>" + "<td>" + $ACL.IdentityReference + "</td>" + "<td>" + $ACL.AccessControlType + "</td>" + "<td>" + $ACL.FileSystemRights + "</td>" + "<td>" + $ACL.IsInherited + "</td>" + "<td>" + $ACL.InheritanceFlags + "</td>" + "<td>" + $ACL.PropagationFlags + "</td></tr>`r"

        }
	}
    # Get subfolders
    $Folders = Get-ChildItem $Folder.FullName -dir -ErrorAction SilentlyContinue
    # If there any folders and we want to inspect more folders
    If ($Folders -and $Deep){
        # Examining new folders
        ForEach ($F in $Folders){
            # Add results to table
            $Rows +=  $(Get-FolderPermissions $F $($Deep - 1) )
        }
    }
    # Return table
    return $Rows
}

# Variable initialization
# Set csv header
$Header = "Folder Path,Identity Reference,Access Control Type,File System Rights,Is Inherited,Inheritance Flags,Propagation Flags"
# Delete csv file
Del $OutFile -ErrorAction "SilentlyContinue"
# Add header to csv
Add-Content -Value $Header -Path $OutFile 
# Clear html table variable
if ([boolean](get-variable "Rows" -ErrorAction SilentlyContinue))
    {Clear-Variable -Name "Rows" -Scope Global}

#Generate table and csv
$Table = Get-FolderPermissions $(Get-Item $RootPath) $FolderDeep

#Create HTML body
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
<H1>$SMTPSubject</H1>
<table>
    <tr>
        <th>Folder Path</th>
        <th>Identity Reference</th>
        <th>Access Control Type</th>
         <th>File System Rights</th>
        <th>Is Inherited</th>
        <th>Inheritance Flags</th>
        <th>Propagation Flags</th>
    </tr>
        $Table
</table>
</body>
</html>
"@
#Send mail 
If ($SMTPServer)
{
    # Set smpt password
    $SecureSMTPPassword = ConvertTo-SecureString $SMTPPassword -AsPlainText -Force
    $SMTPCredential = New-Object System.Management.Automation.PSCredential($SMTPUser,$SecureSMTPPassword)
    Write-Host "Sending Email"
    if ($SMTPSSL){
        Send-MailMessage -To $SMTPRecipient -From $SMTPSender -Attachments $OutFile -SmtpServer $SMTPServer -Subject $SMTPSubject -UseSsl -Port $SMTPPort -Credential $SMTPCredential -BodyAsHtml -Body $HTMLFile -Encoding ([System.Text.Encoding]::utf8)
    }
    else{
        Send-MailMessage -To $SMTPRecipient -From $SMTPSender -Attachments $OutFile -SmtpServer $SMTPServer -Subject $SMTPSubject -Port $SMTPPort -Credential $SMTPCredential -BodyAsHtml -Body $HTMLFile -Encoding ([System.Text.Encoding]::utf8)
    }
}
