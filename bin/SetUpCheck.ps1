$ScriptPath = $PSCommandPath
$dir = Split-Path $ScriptPath

$UserExistVal = get-childitem -path "$dir\User Settings" -Recurse -filter "*$env:UserName*"   
Function InstalledCheck {
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
    $Button = "YesNo"
    $Ask = "It does not seem that $env:UserName has run this before.
    Need to run first time set up?"
    $epromptTitle = "First Time Set up"
    $FirstSetUpPrompt = [Microsoft.VisualBasic.Interaction]::MsgBox($Ask, $Button, $epromptTitle) 

    Switch ($FirstSetUpPrompt) {
        Yes {
            $Rsetup = $dir+"\FirstTimeSetup.ps1"
            &$Rsetup
        }
        No {
            exit
        }
    }
}

if ($UserExistVal-ne $Null){
    exit
}
elseif ($UserExistVal-eq $Null){
    InstalledCheck
    exit
}
else{
    Write-Host "Something went wrong"
}
exit