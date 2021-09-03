#This will self elevate the script so with a UAC prompt since this script needs to be run as an Administrator in order to function properly.
Function AdminCheck {
    If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]'Administrator')) {
        Start-Process powershell.exe -ArgumentList ("-NoProfile -ExecutionPolicy Bypass -File `"{0}`"" -f $PSCommandPath) -Verb RunAs
        Exit
    }
}

### credental set up for Azure and MSoline ###
Function AzureCredentals {
    $Settings = [string[]]$arrayFromFile = Get-Content -Path "$dir\User Settings\settings.txt"
    $global:AdminUser = $Settings[0]
    $global:Pword = Get-Content "$dir\User Settings\$AdminUser.cred" | ConvertTo-SecureString
    #connect to 365 online services
    $global:Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AdminUser, $PWord
    Connect-AzureAD -Credential $global:Credential 
    Connect-MsolService -Credential $global:Credential
}
#connection for Exchange email
Function ExchangeConnect {
    $global:Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection
}