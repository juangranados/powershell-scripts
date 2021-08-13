<#
.DESCRIPTION
	Powershell script for IT support team which allow to reset AD users password in a GUI interface.
    Requirements:
        - PowerShell 4 (Windows Management Framework 4) o highter.
        - Support users must be able to reset AD users password: 
            https://community.spiceworks.com/how_to/1464-how-to-delegate-password-reset-permissions-for-your-it-staff
        - Support computers must install RSAT (Remote Server Administration Tools):
            https://4sysops.com/archives/how-to-install-the-powershell-active-directory-module/
    .NOTES 
	Author:    Juan Granados 
	Date:      January 2018
    Thanks to: https://foxdeploy.com/2015/04/10/part-i-creating-powershell-guis-in-minutes-using-visual-studio-a-new-hope/
#>
$inputXML = @"
<Window x:Name="Password Reset Tool" x:Class="MainWindow"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Password Reset Tool" Height="283" Width="284" 
    ResizeMode="NoResize" WindowStartupLocation="CenterScreen" 
    FocusManager.FocusedElement="{Binding ElementName=TextBoxUser}">
    <Grid Margin="-2,-23,-6,-3" HorizontalAlignment="Left" Width="286">
        <Label x:Name="LabelTitle" Content="Password Reset Tool" HorizontalAlignment="Center" Height="33" Margin="31,33,0,0" VerticalAlignment="Top" Width="180" FontSize="15" Grid.ColumnSpan="2"/>
        <Label x:Name="LabelUser" Content="User" HorizontalAlignment="Left" Margin="15,79,0,0" VerticalAlignment="Top" Height="23" Width="83"/>
        <Label x:Name="LabelPass" Content="Password" HorizontalAlignment="Left" Margin="13,118,0,0" VerticalAlignment="Top" Height="23" Width="83"/>
        <TextBox x:Name="TextBoxUser" HorizontalAlignment="Left" Height="21" Margin="91,81,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="147" ToolTip="User account name" TabIndex="1" Grid.ColumnSpan="2"/>
        <CheckBox x:Name="CheckBoxChangePass" Content="Change password &#xD;&#xA;at next logon" HorizontalAlignment="Left" Height="33" Margin="91,159,0,0" VerticalAlignment="Top" Width="161" TabIndex="3"/>
        <PasswordBox x:Name="TextBoxPass" HorizontalAlignment="Left" Margin="91,117,0,0" VerticalAlignment="Top" Width="147" Height="21" ToolTip="User new password" TabIndex="2" Grid.ColumnSpan="2"/>
        <Button x:Name="ButtonOK" Content="OK" HorizontalAlignment="Center" Margin="47,0,167,18" Width="72" RenderTransformOrigin="0.5,0.5" Height="28" TabIndex="4" IsDefault="True" VerticalAlignment="Bottom">
            <Button.RenderTransform>
                <TransformGroup>
                    <ScaleTransform/>
                    <SkewTransform/>
                    <RotateTransform Angle="360.162"/>
                    <TranslateTransform/>
                </TransformGroup>
            </Button.RenderTransform>
        </Button>
        <Button x:Name="ButtonExit" Content="Exit" Margin="167,0,47,18" Width="72" RenderTransformOrigin="0.5,0.5" Height="28" TabIndex="5" IsCancel="True" HorizontalAlignment="Center" VerticalAlignment="Bottom">
            <Button.RenderTransform>
                <TransformGroup>
                    <ScaleTransform/>
                    <SkewTransform/>
                    <RotateTransform Angle="360.162"/>
                    <TranslateTransform/>
                </TransformGroup>
            </Button.RenderTransform>
        </Button>
        <CheckBox x:Name="CheckBoxEnable" Content="Unlock account" HorizontalAlignment="Left" Height="21" Margin="91,197,0,0" VerticalAlignment="Top" Width="150" TabIndex="3"/>
    </Grid>
</Window>
"@       
 
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'
 
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML

#Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $xaml) 
try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}

#===========================================================================
# Load XAML Objects In PowerShell
#===========================================================================
 
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}
 
Function Get-FormVariables{
if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
get-variable WPF*
}
 
#===========================================================================
# Funciones
#===========================================================================

#Get-FormVariables

$WPFButtonOK.Add_Click({
    try{
        if ([string]::IsNullOrEmpty($WPFTextBoxUser.Text)){
            $Message="Error. User name is empty."
        }
        elseif ([string]::IsNullOrEmpty($WPFTextBoxPass.Password)-and ($WPFCheckBoxEnable.IsChecked -eq $false) ){
            $Message="Error. Password is empty."
        }
        else{
            if (![string]::IsNullOrEmpty($WPFTextBoxPass.Password)){
                $AccountPassword = (ConvertTo-SecureString $WPFTextBoxPass.Password -AsPlainText -Force)
                if ($WPFCheckBoxChangePass.IsChecked){
                Set-ADAccountPassword $WPFTextBoxUser.Text -NewPassword $AccountPassword -Reset -PassThru | Set-ADuser -ChangePasswordAtLogon $True
                }
                else{
                    Set-ADAccountPassword $WPFTextBoxUser.Text -NewPassword $AccountPassword -Reset -PassThru
                }
                $Message="Password changed"
                if ($WPFCheckBoxEnable.IsChecked){
                    Enable-ADAccount -Identity $WPFTextBoxUser.Text
                    $Message+=". User account unlocked"
                }
            }
            elseif ($WPFCheckBoxEnable.IsChecked){
                Enable-ADAccount -Identity $WPFTextBoxUser.Text
                $Message+="User account unlocked"
            }
        }
        Write-Host "$Message" -ForegroundColor Cyan
    }catch{
        Write-Host "Error setting user properties" -ForegroundColor Red
        Write-Host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
        Write-Host "Error Message: $($_.Exception.Message)" -ForegroundColor Red
        $Message="An error has ocurred: "+$($_.Exception.Message)
    }
    [System.Windows.Forms.MessageBox]::Show("$Message")
})

$WPFButtonExit.Add_Click({
    $Form.Close()
})

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

#===========================================================================
# Shows the form
#===========================================================================
Write-Host "Cargando componentes necesarios..."
If(! (Get-MyModule –name "ActiveDirectory")){
    [System.Windows.Forms.MessageBox]::Show("Error loading Active Directory Module for Windows Powershell. Please install RSAT.")
    Exit
}
Import-Module ActiveDirectory
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | out-null
$Form.ShowDialog() | out-null