# Password Encryption for PowerShell scripts

[Original code from PsCustomObject](https://github.com/PsCustomObject/IT-ToolBox)

This two scripts generate a encrypted password and allow use it in PowerShell scripts.

**New-StringEncryption.ps1**: generate a an encrypted string that can only be decrypted on the same machine that performed the original encryption.

**New-StringDecryption.ps1**: takes a Base64 encoded string and will output the clear text version of it.

## Examples

Encrypt password

```powershell
New-StringEncryption.ps1 -StringToEncrypt 'Password'
```
Returns ```fzdxB8+jXgchfghU98mbOc5g==```

Decrypt password and use it

```powershell
$365Username="admin@contoso.onmicrosoft.com"
$365Password="fzdxB8+jXgchfghU98mbOc5g=="
$Secure365AdminPassword = ConvertTo-SecureString -String (C:\Scripts\New-StringDecryption.ps1 -EncryptedString $365Password) -AsPlainText -Force
$365Credentials  = New-Object System.Management.Automation.PSCredential $365Username, $Secure365AdminPassword
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $365Credentials -Authentication Basic -AllowRedirection
Import-PSSession $Session -AllowClobber | Out-Null
# .....
Remove-PSSession $Session
```
