Write-Host "Setting up..."
$ScriptPath = $PSCommandPath
$dir = Split-Path $ScriptPath
$ScriptRoot = Split-Path $dir
#checking folder structure
if(Test-Path -path "$ScriptRoot\Logs") {
    Write-Host "Logs folder found"
} else {
    Write-host "Creating Logs folder"
    new-item -Path $ScriptRoot -Name "Logs" -ItemType "Directory"
}

Start-Transcript -Path "$ScriptRoot\Logs\FirstRun\Log.txt"

Import-Module -Name "$dir\support_files\MSOLModule.psm1"
Import-Module -Name "$dir\support_files\FirstRunSetModule.psm1"

If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]'Administrator')) {
    Start-Process powershell.exe -ArgumentList ("-NoProfile -ExecutionPolicy Bypass -File `"{0}`"" -f $PSCommandPath) -Verb RunAs
    Exit
}

###Checking Logs folder structure###
#Checking top layer
$FileName = @('Add AD Groups Individual','Decom','DynamicsOnly setup','FirstRun','New User Setup','OneDrive Only')
for ($i = 0; $i -lt $FileName.count; $i++) {
    $FN = $FileName[$i]
    if(Test-Path -path "$ScriptRoot\Logs\$FN") {
        Write-Host "$FN logs folder found"
    } else {
        Write-host "Creating  $FN logs folder"
        new-item -Path "$ScriptRoot\Logs" -Name "$FN" -ItemType "Directory"
    }
}

#Checking sublayer 
$FNDSub = @('OneDrive Only','Users','Users-Detailed')
for ($i = 0; $i -lt $FNDSub.count; $i++) {
    $FN = $FNDSub[$i]
    if(Test-Path -path "$ScriptRoot\Logs\Decom\$FN") {
        Write-Host "$FN logs folder found"
    } else {
        Write-host "Creating  $FN logs folder"
        new-item -Path "$ScriptRoot\Logs\Decom" -Name "$FN" -ItemType "Directory"
    }
}

#User Settings
if(Test-Path -path "$dir\User Settings") {
    Write-Host "User Settings folder found"
} else {
    Write-host "Creating User Settings folder"
    new-item -Path $dir -Name "User Settings" -ItemType "Directory"
}

#Updating settings
$SettingLoc = "$dir\User Settings\settings.txt"
Write-Host $SettingLoc
UserCredSet

$FormLocation = "C:\someserver\New Hire Form.xlsx"
Add-Content $SettingLoc $FormLocation


OneDriveBackupLoc
AzureFirstRunModules

Write-Host "--------------------------------------------------"
Write-Host "User" $Credential.UserName "saved"
Write-Host "New User form path: $FormLocation"
Write-host "Returning to menu"
Write-Host "--------------------------------------------------"
Stop-Transcript
exit