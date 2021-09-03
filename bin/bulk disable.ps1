#change this line for the OU you want disabled
$DynamicsOnly = Get-ADUSER -Filter * -SearchBase "OU=Dynamics Only Users, OU=_Other Users,OU=_Users,DC=pmp,DC=com"

#this section is for logging
$ScriptPath = $PSCommandPath
$dir = Split-Path $ScriptPath
$ScriptRoot = Split-Path $dir

Foreach ($i in $DynamicsOnly) {
    $UName = $i.Name
    $UPN = $i.SamAccountName
    Start-Transcript -Path "$ScriptRoot\Logs\Decom\Bulk-Detailed\$UName-DetailedLog.txt"

    $lPath = "$ScriptRoot\Logs\Decom\Bulk-Users\$UName-Decom_Log.txt"
    Get-Date | out-File -FilePath $lPath
    Add-Content -Path $lPath -Value "Setting AD attributes for $UName"

    # changes description to disabled
    try {
        Set-ADUser -Identity $UPN -Description "Disabled"
        Add-Content -Path $lPath -Value "Changed description to 'Disabled'"
    }
    catch {
        Add-Content -Path $lPath -Value "Failed to change description to 'Disabled'"
    }

    # hide from global Address list
    try {
        Get-ADUser $UPN | Set-ADObject -replace @{msExchHideFromAddressLists=$True}
        Get-ADUser $UPN | Set-ADObject -replace @{mailNickname=$UPN}
        Add-Content -Path $lPath -Value "Hidden from global address list"
    }
    catch {
        Add-Content -Path $lPath -Value "Failed to hide from global address list"
    }

    # moves the users OU 
    try {
        Get-ADUser $UPN| Move-ADObject -TargetPath "OU=Disabled Users, OU=_Other Users,OU=_Users,DC=pmp,DC=com"
        Add-Content -Path $lPath -Value "Moved to Disabled Users ou"
    }
    catch {
        Add-Content -Path $lPath -Value "Failed to move to Disabled Users ou"
    }
    
    #resets the default password
    try {
        Set-ADAccountPassword -Identity $UPN -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "Welcome123" -Force)
        Add-Content -Path $lPath -Value "Password reset to default"
    }
    catch {
        Add-Content -Path $lPath -Value "Did not reset password to default"
    }

    #Record groups This is for local ad 
    $ADmember = "$UPN"
    $ADuser = Get-ADUser -filter {(Name -like $ADMember -or SamAccountName -like $ADMember) -and (Enabled -eq $true)}  
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
        } 
    }

    # Removes all groups
    Get-AdPrincipalGroupMembership -Identity $UPN | Where-Object -Property Name -Ne -Value 'Domain Users' | Remove-AdGroupMember -Members $UPN -Confirm:$false
    
    # disable the account
    try {
        Disable-ADAccount -Identity $UPN
        Add-Content -Path $lPath -Value "Disabling user in AD"
    }
    catch {
        Add-Content -Path $lPath -Value "Failed to disable user in AD"
    }

    Stop-Transcript
}

