Write-Host "Setting up..."
$ScriptPath = $PSCommandPath
$dir = Split-Path $ScriptPath
$ScriptRoot = Split-Path $dir

Import-Module -Name "$dir\support_files\MSOLModule.psm1"
Import-Module -Name "$dir\support_files\BlockModule.psm1"
Import-Module -Name "$dir\support_files\OneCopy.psm1"

#Emails to check more can be added
$EmailList = 'company.com','CoolWave.com','Pen14.com','Yeet.com'

#destination for OneDrive backup
$OfficeLoc = [string[]]$arrayFromFile = Get-Content -Path "$dir\User Settings\settings.txt"
if ($OfficeLoc[2] -eq "Honolulu") {
	$BackupLoc = "\\IP\Volume_1\User Backup"
}elseif ($OfficeLoc[2] -eq "Boulder") {
    $BackupLoc = "\\Server\Usershares\User Backup"
}else {
    $BackupLoc = "$ScriptRoot\UserBackup"
}

### Fcn calls ###
AdminCheck
AzureCredentals
ExchangeConnect
AccountDecomPrompt
UserNullCheck
Start-Transcript -Path "$ScriptRoot\Logs\Decom\OneDrive Only\$UserName.txt"

if ($DNEx -eq $NULL) {
    OneDriveBackup

    Write-Host "--------------------------------------------------"
    $UserObjectId = Get-AzureADUser -ObjectId $DecomUPN
    $OBJId = $UserObjectId.ObjectId
    Write-host "Returning to menu"
    Write-Host "--------------------------------------------------"
}

#stop Recording#
Stop-Transcript

exit