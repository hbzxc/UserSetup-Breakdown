Write-Host "Setting up..."

$ScriptPath = $PSCommandPath
$dir = Split-Path $ScriptPath
Import-Module -Name "$dir\support_files\MSOLModule.psm1"
Import-Module -Name "$dir\support_files\NewUserModules.psm1"

$365LicenseType = Get-Content "$dir\support_files\license.txt"
$UserUPN = Get-Content "$dir\support_files\UPN.txt"

AdminCheck
AzureCredentals
SetPrimaryLicense

$UserExistVal = get-msoluser -UserPrincipalName "$UserUPN"-ErrorAction SilentlyContinue

if ($UserExistVal-ne $Null){
Write-Host "Primary License set to: $365LicenseType"
Write-Host "Returning to menu"
Write-Host "--------------------------------------------------"
exit
}
elseif ($UserExistVal-eq $Null){
    Write-Host "Primary License for $UserUPN not set"
    Write-host "Run Set License Redundancy from Main menu"
    Write-Host "--------------------------------------------------"
    exit
}
exit