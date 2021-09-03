
Function UserNullCheck {
    <# checks to see if user input corresponds to a user in ad
    checks is there was any input
        bounces back if empty #>
    If (($DecomSAM -eq $Null) -or ($UserName -eq "")) {
        Write-host "Username = $UserName"
        Write-host "Invalid or No Username Selected"
        pause
        $global:DNEx -eq 1
        exit
    }
    else {
        Write-host "User Name Selected:$DecomSAM"
        Write-host $DecomObj
    }
}
Function AccountDecomPrompt {
    $ScriptRoot = Split-Path $dir
    # Makes the popup windows that prompts for a user
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    
    $title = "Account Decommission"
    $msg   = "User Name of Target Account | No Email Tag

        Example: HBlazier"

    $global:UserName = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)

    #sets the user to global for use in backing up OneDrive
    $DecomObj = Get-ADUser $UserName
    $global:DecomUPN = $DecomObj.UserPrincipalName
    $global:DecomSAM = $DecomObj.SamAccountName
}

Function ADDecom {
    #Checks if the user exists
    UserNullCheck

    Start-Transcript -Path "$ScriptRoot\Logs\Decom\Users-Detailed\$UserName-DetailedLog.txt"

    $lPath = "$ScriptRoot\Logs\Decom\Users\$UserName-Decom_Log.txt"
    $Path = "$ScriptRoot\Logs\Decom\Users\$UserName-ADGroups.csv"

    Get-Date | out-File -FilePath $lPath
    Add-Content -Path $lPath -Value "Setting AD attributes for $UserName"

    # changes description to disabled
    try {
        Set-ADUser -Identity $UserName -Description "Disabled"
        Add-Content -Path $lPath -Value "Changed description to 'Disabled'"
    }
    catch {
        Add-Content -Path $lPath -Value "Failed to change description to 'Disabled'"
    }

    # hide from global Address list
    try {
        Get-ADUser $UserName | Set-ADObject -replace @{msExchHideFromAddressLists=$True}
        Get-ADUser $UserName | Set-ADObject -replace @{mailNickname=$UserName}
        Add-Content -Path $lPath -Value "Hidden from global address list"
    }
    catch {
        Add-Content -Path $lPath -Value "Failed to hide from global address list"
    }

    # moves the users OU 
    try {
        Get-ADUser $UserName| Move-ADObject -TargetPath "OU=Disabled Users, OU=_Other Users,OU=_Users,DC=pmp,DC=com"
        Add-Content -Path $lPath -Value "Moved to Disabled Users ou"
    }
    catch {
        Add-Content -Path $lPath -Value "Failed to move to Disabled Users ou"
    }
    
    #resets the default password
    try {
        Set-ADAccountPassword -Identity $UserName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "Password" -Force)
        Add-Content -Path $lPath -Value "Password reset to default"
    }
    catch {
        Add-Content -Path $lPath -Value "Did not reset password to default"
    }

    #Record groups This is for local ad 
    $ADmember = "$UserName"
    $ADuser = Get-ADUser -filter {(Name -like $ADMember -or SamAccountName -like $ADMember)}  
    Add-Content -Path $lPath -Value "--------------Exported AD groups to csv------------"
    $out = foreach($user in $ADuser)    { 
        $groups = Get-ADPrincipalGroupMembership $user 
        foreach ($group in $groups){ 
        $rec = New-Object PSObject 
            foreach($GP in $group.psobject.Properties) { 
                    foreach($UP in $user.psobject.Properties) { 
                    $rec | Add-Member -Type NoteProperty -Name ("U_" + $UP.Name) -Value $UP.value -Force 
                    $rec | Add-Member -Type NoteProperty -Name ("G_" + $GP.Name) -Value $GP.value -Force 
                    } 
            } 
            $rec|select U_Name, U_DistinguishedName,G_name,G_GroupCategory, G_GroupScope, G_distinguishedName
            Add-Content -Path $lPath -Value $rec.G_name
            Add-Content -Path $Path -Value $rec.G_name
        } 
    }
    $out |Export-Csv $Path -NoTypeInformation
    Add-Content -Path $lPath -Value "--------------Finished Exporting to csv------------"

    # Removes all groups
    Get-AdPrincipalGroupMembership -Identity $UserName | Where-Object -Property Name -Ne -Value 'Domain Users' | Remove-AdGroupMember -Members $UserName -Confirm:$false
    
    # disable the account
    try {
        Disable-ADAccount -Identity $UserName
        Add-Content -Path $lPath -Value "----------------Disabling user in AD---------------"
    }
    catch {
        Add-Content -Path $lPath -Value "Failed to disable user in AD"
    }
}

Function CheckAll {
    $ScriptRoot = Split-Path $dir
    $Path = "$ScriptRoot\Logs\Decom\Users\$UserName-Decom_Log.txt"
    #Connects to office 365
    Import-PSSession $Session
    Write-Host "Connected to Session"
    $ErrorActionPreference = "SilentlyContinue"
    # sets list of emails to check from main BlockUsers
    $EmailArray = @($EmailList)
    for ($i = 0; $i -lt $EmailArray.count; $i++) {
        $UserAccount = $UserName+'@'+$EmailArray[$i]
        # Checks to see what emails exists and skips them if they don't
        $AddressCheck = Get-MsolUser -UserPrincipalName $UserAccount
        If ($AddressCheck){
            Add-Content -Path $Path -Value "---Now updating $UserAccount 365 settings---"
            # disable active sync
            try {
                Set-CASMailbox -Identity $UserAccount -ActiveSyncEnabled $False
                Add-Content -Path $Path -Value "Disabled active sync"
            }
            catch {
                Add-Content -Path $Path -Value "Failed to disable active sync"
            }

            # disable OWA for devices
            try {
                Set-CASMailbox -Identity $UserAccount -OWAforDevicesEnabled $false
                Add-Content -Path $Path -Value "Disabled OWA for devices"
            }
            catch {
                Add-Content -Path $Path -Value "Failed to disable OWA for devices"
            }
            
            # disable owa
            try {
                Set-CASMailbox -Identity $UserAccount -OWAEnabled $false
                Add-Content -Path $Path -Value "Disabled OWA"
            }
            catch {
                Add-Content -Path $Path -Value "Failed to Disabled OWA"
            }

            # hide from Address List
            try {
                Set-Mailbox "$UserAccount" -HiddenFromAddressListsEnabled $True
                Add-Content -Path $Path -Value "Hidden from address list: msol"
            }
            catch {
                Add-Content -Path $Path -Value "Failed to hide from address list: msol"
            }
            
            <#Set message delivery restrictions
            
            try {
                Set-Mailbox -Identity "$UserAccount" -AcceptMessagesOnlyFrom "techteam@company.com"
                Add-Content -Path $Path -Value "Message delivery restricted to techteam@company.com"
            }
            catch {
                Add-Content -Path $Path -Value "Failed to restrict message delivery"
            }
            #>
            
            # set mailbox to shared
            try {
                Set-Mailbox "$UserAccount" -type Shared
                Add-Content -Path $Path -Value "Mailbox set to shared"
            }
            catch {
                Add-Content -Path $Path -Value "Failed to set mailbox to shared"
            }
            
            # disable online account
            try {
                Set-AzureADUser -ObjectID $UserAccount -AccountEnabled $false
                Add-Content -Path $Path -Value "Disabled online account"
            }
            catch {
                Add-Content -Path $Path -Value "Failed to disable online account"
            }
            
            #remove all remaining msol groups
            Add-Content -Path $Path -Value "----------------Removing 365 Groups----------------"
            $msolUser = Get-AzureADUser -ObjectId $UserAccount
            foreach ($GP in (Get-AzureADUserMembership  -ObjectId $msolUser.UserPrincipalName))
            {
                Remove-AzureADGroupMember -ObjectId $GP.ObjectId -MemberId $msolUser.ObjectId
                Add-Content -Path $Path  -Value $GP.Displayname
            }

            # remove licences associated with account
            Add-Content -Path $Path -Value "-----------------Removing licences-----------------"  
            (get-MsolUser -UserPrincipalName $UserAccount).licenses.AccountSkuId |
                foreach{
                Set-MsolUserLicense -UserPrincipalName $UserAccount -RemoveLicenses $_
                Add-Content -Path $Path  -Value $_
                }
                Add-Content -Path $Path -Value "---------------Licences not removed----------------"
                Add-Content -Path $Path -Value Get-AzureADUser -SearchString $UserAccount | Get-AzureADUserMembership
        }
        Else {
            Write-Host"$UserAccount does not exist"
        }
    }
    $date = Get-Date
    Add-Content -Path $Path -Value "Finished"
    Add-Content -Path $Path -Value $date
    Remove-PSSession $Session
    Write-Host "Ending Session"
}