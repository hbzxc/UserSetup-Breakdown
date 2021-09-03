Write-Host "Setting up..."

$ScriptPath = $PSCommandPath
$dir = Split-Path $ScriptPath
$ScriptRoot = Split-Path $dir
Import-Module -Name "$dir\support_files\MSOLModule.psm1"
Import-Module -Name "$dir\support_files\NewUserModules.psm1"

### Fcn Calls ###
AdminCheck
FirstNamePrompt
LastNamePrompt
UserNamePrompt

### Default Settings ###
$UserUPN="$UserName@company.com"

Start-Transcript -Path "$ScriptRoot\Logs\DynamicsOnly setup\$UserName-DetailedLog.txt"

$DefaultPass = ConvertTo-SecureString "Password" -AsPlainText -Force

#Check if username exists
$UserNameExistCheck = Get-ADUser -Filter {sAMAccountName -eq $UserName}
if ($UserNameExistCheck -ne $null){
    Write-Host "User name $UserName already exists"
    exit
}

#creates new ad account
New-ADUser -Name "$FirstName $LastName" -GivenName $FirstName -Surname $LastName -DisplayName "$FirstName $LastName" -SamAccountName $UserName -UserPrincipalName "$UserUPN" -Path "OU=Dynamics Only Users,OU=_Other Users,OU=_Users,DC=pmp,DC=com" -AccountPassword($DefaultPass) -Enabled $true
Set-ADUser -Identity $UserName -Description "Dynamics Only"

Write-Host "--------------------------------------------------"
write-host "Set up for $FirstName $LastName in AD complete"
Write-host "Don't forget to add user in Dynamics"
Write-host "Returning to menu"
Write-Host "--------------------------------------------------"
Stop-Transcript
exit