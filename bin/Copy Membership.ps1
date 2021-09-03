Function Master {
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    
    $title = 'Master'
    $msg   = 'User to copy from:'
    
    $global:Master = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
    $UserNameExistCheck = Get-ADUser -Filter {sAMAccountName -eq $Master}
    if (-not $Master){
        Write-Host "Canceled no input"
        exit    
    }
    elseif ($UserNameExistCheck -eq $null){
        Write-Host "User name $Master doesn't exist"
        exit
    }
}

Function CopyTo {
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    
    $title = 'Copy to User'
    $msg   = 'Copy groups to this user:'
    
    $global:CopyTo = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
    $UserNameExistCheck = Get-ADUser -Filter {sAMAccountName -eq $CopyTo}
    if (-not $CopyTo){
        Write-Host "Canceled no input"
        exit    
    }
    elseif ($UserNameExistCheck -eq $null){
        Write-Host "User name $CopyTo doesn't exist"
        exit
    }
}
Master
CopyTo
$CopyFromUser = Get-ADUser $Master -prop MemberOf
$CopyToUser = Get-ADUser $CopyTo -prop MemberOf
$CopyFromUser.MemberOf | Where{$CopyToUser.MemberOf -notcontains $_} |  Add-ADGroupMember -Members $CopyToUser

