Write-Host "Setting up..."

$ScriptPath = $PSCommandPath
$dir = Split-Path $ScriptPath
$ScriptRoot = Split-Path $dir
Import-Module -Name "$dir\support_files\MSOLModule.psm1"
Import-Module -Name "$dir\support_files\NewUserModules.psm1"

Start-Transcript -Path "$ScriptRoot\Logs\New User Setup\setupLog.txt"

### default password for new account(s)###
$DefaultPass = ConvertTo-SecureString "Password" -AsPlainText -Force

###AD groups lists###
$somecompanyADGroups = 'Hawaii Construction','SG somecompany PROJECTS','SG_DYNAMICS_SL','SG REFERENCES','SG_CP_Weekly_Logs_RO','SG NHCDC'

$otherCompanyADGroups = 'SG otherCompany CM','SG otherCompany','SG otherCompany GREENWAVE','SG otherCompany PROJECTS','SG otherCompany REFERENCES','SG otherCompany SCAN','SG otherCompany WORKSHARE','SG_DYNAMICS_SL'

$AccountingADGroups = 'SG ACCOUNTING','SG somecompany PROJECTS','SG FINANCIAL REPORTING','SG NHCDC','SG_DYNAMICS_SL'

### Where to make user folders ###
$UserFolderLocation = '\\SomeIp\qemydocuments','\\SomeIp\QueenEmmaScans'

$otherCompanyFolderLocation = '\\AnotherIP\otherCompany\Home','\\AnotherIP\otherCompany\Scans'

### Signature location for auto emails ###
$Signature = get-content -path "$dir\User Settings\signature.txt" -raw

### Main Function Calls ###
AdminCheck
AzureCredentals
ResponseIDPrompt
ExcelFormGet
CheckLicInventory
Stop-Transcript
UserNamePrompt
Start-Transcript -Path "$ScriptRoot\Logs\New User Setup\$UserName-DetailedLog.txt"
LocationPrompt
DivisionPrompt
CreateADUser
SecondEmailSetUp
AssignLaptop
AssignFOB
SetADGroups
CheckForSync
NewHireSheet

Write-Host "Assigning primary licence waiting for 90 full sync"
Start-Sleep -s 90

### Final checks ###
SetPrimaryLicense
MailConfirmation

$UserExistVal = get-msoluser -UserPrincipalName "$UserUPN"-ErrorAction SilentlyContinue

#double checking the the licese was assigned its really finickey
$count = 0
While (($False -eq $UserExistVal.isLicensed) -and ($count -le 5)){
    Write-Host "Primary License for $UserUPN not set"
    Write-host "Running Set License Redundancy"
    Write-Host "--------------------------------------------------"
    SetPrimaryLicense
    Start-Sleep -s 15
    $count++
    $UserExistVal = get-msoluser -UserPrincipalName "$UserUPN"-ErrorAction SilentlyContinue
}

Write-Host "--------------------------------------------------"
write-Host "Set up for $FirstName $LastName in AD complete"
if ($NoSecondaryEmail -eq 0) {
    Write-Host "Second Email setup:$SecondaryEmail"
}

if ($True -eq $UserExistVal.isLicensed){
Write-Host "Primary License set to: $365Skew"
Write-Host "Returning to menu"
Write-Host "--------------------------------------------------"
Stop-Transcript
exit
}
