Write-Host "Setting up..."
$ScriptPath = $PSCommandPath
$dir = Split-Path $ScriptPath
$ScriptRoot = Split-Path $dir

Import-Module -Name "$dir\support_files\MSOLModule.psm1"
Import-Module -Name "$dir\support_files\BlockModule.psm1"
Import-Module -Name "$dir\support_files\OneCopy.psm1"

#Emails to check more can be added
$EmailList = 'coolcompany.com','othercompany.com','anothercompany.com','yetanothercompany.com'

#destination for OneDrive backup
$OfficeLoc = "$dir\User Settings\settings.txt"
if ([System.IO.File]::Exists($OfficeLoc) -eq $True) {
	$BackupLoc = "\\someip\Volume_1\User Backup"
}
else {
    $BackupLoc = "\\someserver\Usershares\User Backup"
}

### Fcn calls ###
AdminCheck
AzureCredentals
ExchangeConnect
AccountDecomPrompt
ADDecom
if ($DNEx -eq $NULL) {
    #Check if OneDrive is to be backed up
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
    $Button = "YesNo"
    $Ask = "Do you want to backup the users onedrive?"
    $epromptTitle = "OneDrive Backup"
    $LicCheckPrompt = [Microsoft.VisualBasic.Interaction]::MsgBox($Ask, $Button, $epromptTitle)
    Switch ($LicCheckPrompt) {
        Yes {
            OneDriveBackup
        }
        No {
            Write-Host "Skipping backup"
        }
    }
    
    CheckAll

    Write-Host "--------------------------------------------------"
    Write-Host "User $UserName blocked"
    Write-Host "Dont forget to disable External Sharing in OneDrive"
    $UserObjectId = Get-AzureADUser -ObjectId $DecomUPN
    $OBJId = $UserObjectId.ObjectId
    Start-Process "https://admin.microsoft.com/Adminportal/Home#/homepage/:/UserDetails/$OBJId/OneDrive"
    Write-host "Returning to menu"
    Write-Host "--------------------------------------------------"
}


### Will delete online accounts if uncommented (not tested probably needs to be inserted into one of the BlockModule Functions) ###
#Remove-AzureADUser -ObjectId $useraccount

#stop Recording#
Stop-Transcript

exit