Write-Host "Setting up..."

$ScriptPath = $PSCommandPath
$dir = Split-Path $ScriptPath
$ScriptRoot = Split-Path $dir
Import-Module -Name "$dir\support_files\NewUserModules.psm1"

Start-Transcript -Path "$ScriptRoot\Logs\Add AD Groups Individual\AddedGroups-DetailedLog.txt"

###AD groups lists###
$OtherCompanyADGroups = 'Hawaii Construction','SG OtherCompany PROJECTS','SG_NHCDC_RO','SG REFERENCES'

$SomeCompanyADGroups = 'SG SomeCompany CM','SG SomeCompany','SG SomeCompany GREENWAVE','SG SomeCompany PROJECTS','SG SomeCompany REFERENCES','SG SomeCompany SCAN','SG SomeCompany WORKSHARE'

$AccountingADGroups = 'SG ACCOUNTING','SG OtherCompany PROJECTS','SG FINANCIAL REPORTING','SG_NHCDC_RO'

### Fcn Call ###
UserNamePrompt
SetADGroups

Write-Host "--------------------------------------------------"
write-Host "AG groups Set"
Write-Host "--------------------------------------------------"
Stop-Transcript
exit