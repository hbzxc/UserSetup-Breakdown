Function OneDriveBackup {
    $ErrorActionPreference = "SilentlyContinue"
    #################### Parameters ###########################################
    $listUrl = "Documents"
    ###########################################################################
    ###Problems###
    #does not work if url is too long
    ###########################################################################
    #Credit to https://christopherclementen.wordpress.com/2017/08/14/download-all-files-from-library/

    Connect-MsolService -Credential $global:Credential
    $InitialDomain = Get-MsolDomain | Where-Object {$_.IsInitial -eq $true}
    $departingUserUnderscore = $DecomUPN -replace "[^a-zA-Z]", "_"
    #check for company user to ignore the '-'
    if (($departingUserUnderscore -match "company_hawaii") -eq $True){
        $departingUserUnderscore = $departingUserUnderscore -replace "company_hawaii", "company-hawaii"
    }
    $SharePointAdminURL = "https://$($InitialDomain.Name.Split(".")[0])-admin.sharepoint.com"
    $webUrl = "https://$($InitialDomain.Name.Split(".")[0])-my.sharepoint.com/personal/$departingUserUnderscore"
    
    Connect-SPOService -Url $SharePointAdminURL -Credential $Credential
    Set-SPOUser -Site $webUrl -LoginName $AdminUser -IsSiteCollectionAdmin $true 
    Set-SPOUser -Site $webUrl -LoginName $AdminUser -SharingCapability Disabled

    Connect-PnPOnline -Url $webUrl -Credential $Credential
    $web = Get-PnPWeb
    $list = Get-PNPList -Identity $listUrl

    function ExistingUserCheck {
        if ((Test-Path -path $destination) -eq $false) {
            Write-Host "No User backup exists making a new folder"
            New-Item -Path $destination -ItemType Directory 
        }
    }
    function ProcessFolder($folderUrl, $destinationFolder) {

        $folder = Get-PnPFolder -RelativeUrl $folderUrl
        $tempfiles = Get-PnPProperty -ClientObject $folder -Property Files
    
        if (!(Test-Path -path $destinationfolder)) {
            $dest = New-Item $destinationfolder -type directory 
        }

        $total = $folder.Files.Count
        For ($i = 0; $i -lt $total; $i++) {
            $file = $folder.Files[$i]
            Write-Host "Copying " $file.name
            Get-PnPFile -ServerRelativeUrl $file.ServerRelativeUrl -Path $destinationfolder -FileName $file.Name -AsFile
            Add-Content -Path $Path -Value $file.name
        }
    }

    function ProcessSubFolders($folders, $currentPath) {
        foreach ($folder in $folders) {
            $tempurls = Get-PnPProperty -ClientObject $folder -Property ServerRelativeUrl    
            #Avoid Forms folders
            if ($folder.Name -ne "Forms") {
                $targetFolder = $currentPath +"\"+ $folder.Name;
                ProcessFolder $folder.ServerRelativeUrl.Substring($web.ServerRelativeUrl.Length) $targetFolder 
                $tempfolders = Get-PnPProperty -ClientObject $folder -Property Folders
                ProcessSubFolders $tempfolders $targetFolder
                Write-Host "Creating folder" $folder.Name
                Add-Content -Path $Path -Value $folder.Name
            }
        }
    }

    $ScriptRoot = Split-Path $dir
    $Path = "$ScriptRoot\Logs\Decom\Users\$UserName-Decom_Log.txt"
    $destination = "$BackupLoc\$DecomSAM\OneDriveBackup"
    #Check if the target user exists
    ExistingUserCheck
    #Download root files
    #append file path
    Add-Content -Path $Path -Value "-----------------Backing up OneDrive---------------"
    ProcessFolder $listUrl $destination + "\" 
    #Download files in folders
    $tempfolders = Get-PnPProperty -ClientObject $list.RootFolder -Property Folders
    ProcessSubFolders $tempfolders $destination + "\"
    Compress-Archive -Path $destination -DestinationPath "$BackupLoc\$DecomSAM\$UserName"
    Remove-Item -Path $destination -Recurse
    Add-Content -Path $Path -Value "--------------Done Backing up OneDrive-------------"
    # Remove Global Admin from Site Collection Admin role
    Set-SPOUser -Site $webUrl -LoginName $AdminUser -IsSiteCollectionAdmin $false
}