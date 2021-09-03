### Core Functions ###
Function ExcelFormGet {
<# Notes: 
Pulls from the excel sheet created by ms forms "New Hire Form.xls"
After getting the row number the collum number is manually set under Collum List
Those values are then set to global variables to be called later when setting up the user acount

This Function also checks to see if a second email is needed by looking for 
    the value of email2. If email2 = None then that will set a global variable 
    to skip the secondemail setup function
A similar check is also applied to the prefred name
    if nothing is entered it will set the displayname variable to the first + last name
Additonaly licence values are converted to something msol will recognize 
#>

    #form location
    $Settings = [string[]]$arrayFromFile = Get-Content -Path "$dir\User Settings\settings.txt"
    $NewHireFormLocation = $Settings[1]
    <#starting row
    It determines the row number based on the response ID
    Because the sheet has a headder it is off by 1,
    so to compensate the entered value is increased by 1.

    Offset is now huge. dont want to bother making another form
    #>
    $global:startRow = $ResponseID - 19

    #counter to check for secondary email
    $global:NoSecondaryEmail = 0

    #Collum Lists
    $SubmitterCol = 4
    $FirstNameCol = 6
    $LastNameCol = 7
    $PreferedNameCol = 8 
    $ManagerCol = 9
    $JobTitleCol = 10
    $DivisionCol = 11
    $CompanyCol = 12
    $OfficeLocationCol = 13
    $OfficeLicenseCol = 17
    $PrimaryEmailCol = 18
    $SecondaryEmailCol = 19
    $PhoneNUmberCol = 22
    

    $excel = New-Object -Com Excel.Application
    $wb = $excel.Workbooks.Open("$NewHireFormLocation")
    $sh = $wb.Sheets.Item(1)
    $global:Submitter = $sh.Cells.Item($startRow, $SubmitterCol).Value2
    $global:FirstName = $sh.Cells.Item($startRow, $FirstNameCol).Value2
    $global:LastName = $sh.Cells.Item($startRow, $LastNameCol).Value2
    $global:PreferedName = $sh.Cells.Item($startRow, $PreferedNameCol).Value2
    $global:Manager = $sh.Cells.Item($startRow, $ManagerCol).Value2
    $global:JobTitle = $sh.Cells.Item($startRow, $JobTitleCol).Value2
    $global:Company = $sh.Cells.Item($startRow, $CompanyCol).Value2
    $global:LocationRaw = $sh.Cells.Item($startRow, $OfficeLocationCol).Value2
    $global:DivisionRaw = $sh.Cells.Item($startRow, $DivisionCol).Value2
    $OfficeLicenseRaw = $sh.Cells.Item($startRow, $OfficeLicenseCol).Value2
    $Email1 = $sh.Cells.Item($startRow, $PrimaryEmailCol).Value2
    $Email2 = $sh.Cells.Item($startRow, $SecondaryEmailCol).Value2
    $global:PhoneNumber = $sh.Cells.Item($startRow, $PhoneNumberCol).Value2

    #splitting Manager name
    $ManagerFirst = $Manager.Split(' ')[0]
    $ManagerLast = $Manager.Split(' ')[1]
    $ManagerOB = get-aduser -filter {(surname -eq $ManagerLast) -and (GivenName -eq $ManagerFirst)}
    $MCount = $ManagerOB.count
    if ($MCount -le 1) {
        $global:ManagerUPN = $ManagerOB.sAMAccountName
    }
    else {
        [Microsoft.VisualBasic.Interaction]::MsgBox("Manager not set $MCount accounts found", "OKOnly,SystemModal,Critical", "Exclamation")
        $global:ManagerUPN = $ManagerOB.sAMAccountName
    }
    #Check if row is populated
    if (($FirstName -eq $null) -and ($LastName -eq $null)) {
        Write-Host "$ResponseID does not exist double check response number"
        exit
    }

    #appending .com to emails
    if ($Email2 -eq 'None') {
        $global:NoSecondaryEmail = 1
    }
    $global:PrimaryEmail = "$Email1.com"
    $global:SecondaryEmail = "$Email2.com"


    #Reorganizing Licence Format
    If ($OfficeLicenseRaw -eq 'Basic Web Email ($4/mo)') {
        $global:365LicenseType = 'EXCHANGESTANDARD'
        $global:365Skew = 'E1'
    }
    If ($OfficeLicenseRaw -eq 'Email with Office Install ($16/mo)') {
        $global:365LicenseType = 'ENTERPRISEPREMIUM'
        $global:365LicenseType2 = 'OFFICESUBSCRIPTION'
        $global:365Skew = 'E1 & 365 Pro+ '
    }
    If ($OfficeLicenseRaw -eq 'Email with Office Install + Skype ($20/mo)') {
        $global:365LicenseType = 'ENTERPRISEPACK'
        $global:365Skew = 'E3'
    }
    If ($OfficeLicenseRaw -eq 'Email with Office Install + Skype + Toll Free Conferencing ($35/mo)') {
        $global:365LicenseType = 'ENTERPRISEPREMIUM'
        $global:365Skew = 'E5'
    }


    #Checking prefered Name
    if ($PreferedName -eq $Null) {
        $global:PreferedName = "$FirstName $LastName"
    }
    $excel.Workbooks.Close()
}
Function CheckLicInventory {
    <# Notes
Asks if you would like to check the license inventory
This script will not automatically set license's,
because if done wrong it could order any number of license's
    It can be done but I did not want to test it
#>
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
    $Button = "YesNo"
    $Ask = "User needs Office Lic $365Skew
    Second Email status: $SecondaryEmail  "
    $epromptTitle = "Check the License Inventory"
    $LicCheckPrompt = [Microsoft.VisualBasic.Interaction]::MsgBox($Ask, $Button, $epromptTitle) 

    Switch ($LicCheckPrompt) {
        Yes {
            $MAECObj = (Get-MSOLAccountSku | where AccountSkuID -match "$365LicenseType")
            $MainAccountEmailCheck = ($MAECObj | out-string)
            $SAECObj = (Get-MSOLAccountSku | where AccountSkuID -match "EXCHANGESTANDARD")
            $SecondaryAccountEmailCheck = ($SAECObj | Out-String)
            $LicenseCheck = $MainAccountEmailCheck
            if ($365LicenseType2 -ne $null -and $NoSecondaryEmail -eq 0) {
                $MainAccountCheck2 = (Get-MSOLAccountSku | where AccountSkuID -match "$365LicenseType2" | out-string)
                $LicenseCheck = "$MainAccountEmailCheck
                User needs an office License
                $MainAccountCheck2
                User also needs an E1 License for a secondary email
                $SecondaryAccountEmailCheck"
            } elseif ($NoSecondaryEmail -eq 0 -and $365LicenseType2 -eq $null) {
                $LicenseCheck = "$MainAccountEmailCheck
                User also needs an E1 License for a secondary email
                $SecondaryAccountEmailCheck"
            } elseif ($365LicenseType2 -ne $null -and $NoSecondaryEmail -eq 1) {
                $LicenseCheck = "$MainAccountEmailCheck
                User needs an office License
                $MainAccountCheck2"
            }
            [Microsoft.VisualBasic.Interaction]::MsgBox("
            User Needs a $365Skew
            $LicenseCheck","OKOnly","$365Skew")

            [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
            $Button = "YesNo"
            $Ask = "Need to add more Licences"
            $epromptTitle = "Adding more licences"
            $AddMorePrompt = [Microsoft.VisualBasic.Interaction]::MsgBox($Ask, $Button, $epromptTitle)
       
            Switch ($AddMorePrompt) {
                Yes {
                    if ($365LicenseType2 -ne $null -and $NoSecondaryEmail -eq 1) {
                        $Body = "Hello,

                        I need to add 1 $365LicenseType license to GSI service group.
                        
                        Regards,
                        $Signature"
                        if ($MAECObj.ActiveUnits -eq $MAECObj.ConsumedUnits) {
                            $LicMailSet = @{
                                Subject    = "Adding licenses for GSI Service Group"
                                Body       = "$Body"
                                To         = "$Submitter"
                                From       = "$AdminUser"
                                SmtpServer = "smtp.office365.com"
                                Port       = "587"
                                Credential = $Credential
                            }
                            Send-MailMessage @LicMailSet -UseSsl
                        }
                    } elseif ($NoSecondaryEmail -eq 0 -and $365LicenseType2 -eq $null) {
                        $Body = 0
                        if (($MAECObj.ActiveUnits -eq $MAECObj.ConsumedUnits) -and ($SAECObj.ActiveUnits -eq $SAECObj.ConsumedUnits)){
                            $Body = "Hello,

                            I need to add 1 $365LicenseType and 1 EXCHANGESTANDARD licence to GSI service group.
                            
                            Regards,
                            $Signature"
                        } elseif (($MAECObj.ActiveUnits -eq $MAECObj.ConsumedUnits) -and ($SAECObj.ActiveUnits -ne $SAECObj.ConsumedUnits)){
                            $Body = "Hello,

                            I need to add 1 $365LicenseType to GSI service group.
                            
                            Regards,
                            $Signature"
                        } elseif (($MAECObj.ActiveUnits -ne $MAECObj.ConsumedUnits) -and ($SAECObj.ActiveUnits -eq $SAECObj.ConsumedUnits)){
                            $Body = "Hello,

                            I need to add 1 EXCHANGESTANDARD licence to GSI service group.
                            
                            Regards,
                            $Signature"
                        }
                        if ($Body -ne 0) {
                            $LicMailSet = @{
                                Subject    = "Adding licenses for GSI Service Group"
                                Body       = "$Body"
                                To         = "$Submitter"
                                From       = "$AdminUser"
                                SmtpServer = "smtp.office365.com"
                                Port       = "587"  
                                Credential = $Credential
                            }
                            Write-Host "Sending an email to CSP requesting a licence"
                            Send-MailMessage @LicMailSet -UseSsl
                        } elseif ($Body -eq 0){
                            Write-Host "Licences are avaliable. Not sending a request"
                        } else {
                            [Microsoft.VisualBasic.Interaction]::MsgBox("Check your outbox something might have gone wrong", "OKOnly,SystemModal,Critical", "Exclamation")
                        }
                    }   

                    <# Old way to add licences
                    $IE = new-object -com internetexplorer.application
                    $IE.navigate2("https://portal.office.com/adminportal/home#/licenses")
                    $IE.visible = $true
                    #>
                }
                No {

                } 
            }

        }
        No {
            Write-host "Skipping Check"
        }
    }
}

#function prompts for AD attributes. Collects User Input for account creation
Function ResponseIDPrompt {
    $title = "ID"
    $msg = "Enter The Response List Number"
    Add-Type -AssemblyName Microsoft.VisualBasic
    $Response = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
    $global:ResponseID = $Response -as [Double]
    if (-not $Response){
        Write-Host "Canceled no input"
        exit
    }
}
Function FirstNamePrompt {
    #this if for Dynamics Only users
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    
    $title = 'First Name'
    $msg   = 'Enter First Name:'
    
    $global:FirstName = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
    if (-not $FirstName){
        Write-Host "Canceled no input"
        exit
    }
}
    Function LastNamePrompt {
    #This is for dynamics only users
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    
    $title = 'Last Name'
    $msg   = 'Enter Last Name:'
    
    $global:LastName = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
    if (-not $LastName){
        Write-Host "Canceled no input"
        exit
    }
}
Function UserNamePrompt {
    <# Notes:
Asks for a username Typically First letter of first name and lastname
This is left up to the account creators discretion
#>
    [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

    $title = 'UserName'
    $msg = "Enter UserName:

    First Name: $FirstName

    Last Name: $LastName
        "

    $global:UserName = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
    $UserNameExistCheck = Get-ADUser -Filter {sAMAccountName -eq $UserName}
    if (-not $UserName){
        Write-Host "Canceled no input"
        exit    
    }
    elseif ($UserNameExistCheck -ne $null){
        Write-Host "User name $UserName already exists"
        exit
    }
    
    #Sets Main Account UPN
    $global:UserUPN = "$UserName@$PrimaryEmail"
}

<#box lists made an independent function for editing ease
    List possible options in a local array #>
Function BoxList {
    for ($i = 0; $i -lt $BoxArray.count; $i++) {
        $Box = $BoxArray[$i]
        [void] $listBox.Items.Add("$Box")
        }
}
#arrays for list prompt windows
#AD folder structure.
Function LocationList {  
    $BoxArray = @('Boulder','Guam','Honolulu','Kamuela','Maryland','Remote','Washington')
    BoxList
}
Function DivisionList {

    if ($location -eq 'Boulder') {  
        $BoxArray = @('Administrative','Business Development','Construction','Environmentals','MEC (UXO)')
        BoxList
    }

    elseif ($Location -eq 'Guam') {
        $BoxArray = @('Construction')
        BoxList
    }

    elseif ($Location -eq 'Honolulu') {
        $BoxArray = @('Administrative','Construction','Engineering','Environmentals','MEC (UXO)','Water World')
        BoxList
    }

    elseif ($Location -eq 'Kamuela') {
        $BoxArray = @('MEC (UXO)')
        BoxList
    }

    elseif ($Location -eq 'Maryland') {
        $BoxArray = @('MEC (UXO)','Professional Services')
        BoxList
    }

    elseif ($Location -eq 'Remote') {
        $BoxArray = @('Administrative','Business Development','Construction','Environmentals','MEC (UXO)','Professional Services')
        BoxList
    }

    elseif ($Location -eq 'Washington') {
        $BoxArray = @('Construction')
        BoxList
    }

}
Function ADGroupList {  
    $BoxArray = @('enviro','Accounting','engin')
    BoxList   
}

### box prompts ###
#prompt window template
Function PromptWindow {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(300, 250)
    $form.StartPosition = 'CenterScreen'

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(75, 170)
    $OKButton.Size = New-Object System.Drawing.Size(75, 23)
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(150, 170)
    $CancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $CancelButton.Text = 'Cancel'
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(280, 50)
    $label.Text = $FormInfo
    $form.Controls.Add($label)

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(10, 70)
    $listBox.Size = New-Object System.Drawing.Size(260, 20)
    $listBox.Height = 80

    ListSelection

    $form.Controls.Add($listBox)

    $form.Topmost = $true

    $result = $form.ShowDialog()

    PromptSelection
}
#Sets the box list to show
Function ListSelection {
    if ($ListSetting -eq 'Location') {
        LocationList
    }
    elseif ($ListSetting -eq 'Division') {
        DivisionList
    }
    elseif ($ListSetting -eq 'ADGroup') {
        ADGroupList
    }
    else {
        Write-host "List Selection Error"
        Write-host "ListSetting = $ListSetting"
        pause
    }
}
#fills in window prompt settings based on prompt
Function LocationPrompt {
    $global:ListSetting = 'Location'
    $Title = 'Choose a Location'
    $FormInfo = "Please select a location:

    Listed Location: $LocationRaw"
    Function PromptSelection {
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $global:Location = $listBox.SelectedItem
        }
        if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
            exit
        }
    }
    PromptWindow
}
Function DivisionPrompt {
    $global:ListSetting = 'Division'
    $Title = 'Choose a Division'
    $FormInfo = "Please select a division:
    Manager: $Manager
    Division: $DivisionRaw"
    Function PromptSelection {
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $global:division = $listBox.SelectedItem
        }
        if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
            exit
        }
    }
    PromptWindow
}

#prompts for addming laptop names not nessary and can be cancled out
Function AssignLaptopPrompt {
    $title = "Assign Laptop"
    $msg = "Please Enter the Machine Name Below"
    Add-Type -AssemblyName Microsoft.VisualBasic
    $Response = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
    $global:ComputerName = $Response
    $global:LaptopName = $Response
    if (-not $Response){
        Write-Host "Will add to sheet later"
        break
    }
}

Function AssignDesktopPrompt {
    $title = "Assign Desktop"
    $msg = "Please Enter the Machine Name Below"
    Add-Type -AssemblyName Microsoft.VisualBasic
    $Response = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
    $global:ComputerName = $Response
    $global:DesktopName = $Response
    if (-not $Response){
        Write-Host "Will add to sheet later"
        break
    }
}

#creates new ad account
Function CreateADUser {
    New-ADUser -Name "$FirstName $LastName" -GivenName $FirstName -Surname $LastName -displayname "$PreferedName" -SamAccountName $UserName -UserPrincipalName "$UserUPN" -Department "$DivisionRaw" -title "$JobTitle" -office "$LocationRaw" -EmailAddress "$UserUPN" -company "$Company" -OfficePhone "$PhoneNumber" -Path "OU=$division,OU=$Location,OU=_Users,DC=pmp,DC=com" -AccountPassword($DefaultPass) -Enabled $true -ErrorVariable NewUserError;
    if ($NewUserError) {
        [Microsoft.VisualBasic.Interaction]::MsgBox("$UserName is already in AD", "OKOnly,SystemModal,Critical", "Error")
    }
    Set-ADUser -Identity $UserName -Manager $ManagerUPN
}
Function ADGroupSelectionPrompt {
    $global:ListSetting = 'ADGroup'
    $Title = 'Choose a Division'
    $FormInfo = "Please select an AD Group
    Location: $Location
    Manager: $Manager
    Division: $DivisionRaw"
    Function PromptSelection {
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            if ($listBox.SelectedItem -eq "enviro") {
                $global:ADCompany = "enviro" 
                enviroADGroups
            }
            elseif ($listBox.SelectedItem -eq "engin") {
                $global:ADCompany = "engin"
                CPEADGroups
            }
            elseif ($listBox.SelectedItem -eq "Accounting") {
                $global:ADCompany = "Accounting"
                AccountingADGroups
            }
        }
        if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
            Write-Host "Will set AD groups later"
            break
        }
    }
    PromptWindow
}
#Remote code execution on the ad server to force a sync
Function SyncAD {
    $ADSession = New-PSSession -ComputerName #name of target machine goes here
    Invoke-Command -Session $ADSession -ScriptBlock {Import-Module -Name 'ADSync'} -ErrorAction SilentlyContinue
    Invoke-Command -Session $ADSession -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta} -ErrorAction SilentlyContinue
    Remove-PSSession $ADSession
}
#Waiting for Sync
Function CheckForSync {
    <# Notes:
Waits for AD to sync to MSonline then moves to next section
#>
    #waiting feedback
    $Loading = 0..100
    $idx = 0
    #Checks if user info is avaliable in MSOnline
    $msolUserExist = Get-MsolUser -UserPrincipalName $UserUPN -ErrorAction SilentlyContinue

    while ($msolUserExist -eq $Null) {
	
        Write-Progress -Activity "Checking if $UserUPN is Synced to MsOnline" -Status "Waiting..." -PercentComplete $idx
        $idx++
        $msolUserExist = Get-MsolUser -UserPrincipalName $UserUPN -ErrorAction SilentlyContinue
        if ($idx -ge $Loading.Length) {
            $idx = 0
        }
        Start-Sleep 1
        SyncAD 
    }

}

#Set Primary License. Will fail if user is not synced to 365
Function SetPrimaryLicense {
    #sets license type
    $PlanName = "$365LicenseType"
    $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
    $License.SkuId = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $PlanName -EQ).SkuID
    $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    $LicensesToAssign.AddLicenses = $License
    
    #sets license region
    Set-MsolUser -UserPrincipalName "$UserUPN" -UsageLocation US
    
    #sets secondary license type (if just mail and office)
    if ($365LicenseType2 -ne $Null) {
        $PlanName2 = "$365LicenseType2"
        $License2 = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
        $License2.SkuId = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $PlanName2 -EQ).SkuID
        $LicensesToAssign2 = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
        $LicensesToAssign2.AddLicenses = $License2
        Set-AzureADUserLicense -ObjectId $UserUPN -AssignedLicenses $LicensesToAssign2 -ErrorVariable NoPrimaryLicense;
    }
    
    #assigns avaliable licenses
    Set-AzureADUserLicense -ObjectId $UserUPN -AssignedLicenses $LicensesToAssign -ErrorVariable NoPrimaryLicense;

    #output for License if first one fails 
    $365LicenseType > "$dir\support_files\license.txt"
    $UserUPN > "$dir\support_files\UPN.txt"
    if ($NoPrimaryLicense) {
        [Microsoft.VisualBasic.Interaction]::MsgBox("No $365Skew Licenses avaliable", "OKOnly,SystemModal,Critical", "Error")
    }
}

### Optional Functions ###

Function SecondEmailSetUp {
    <# Notes:
    Prompt If the user needs a second email and will create an account with the same username. 
    New email can be selected from a list. Will automatically assign the account an E1 license
    #>
    
    if ($NoSecondaryEmail -eq 0) {
        #creat an additonal 365 account 
        Write-host "Setting up Secondary Email"
        $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
        $PasswordProfile.ForceChangePasswordNextLogin = $false
        $PasswordProfile.Password = "Password"
        $SecondaryLicense = "EXCHANGESTANDARD"
        $PlanName = $SecondaryLicense
        $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
        $License.SkuId = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $PlanName -EQ).SkuID
        $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
        $LicensesToAssign.AddLicenses = $License
            
        New-AzureADUser -DisplayName "$FirstName $LastName" -GivenName "$FirstName" -SurName "$LastName" -JobTitle "$JobTitle"  -Department "$DivisionRaw"  -PhysicalDeliveryOfficeName "$LocationRaw" -UserPrincipalName "$UserName@$SecondaryEmail" -UsageLocation US -MailNickName "$UserName"  -PasswordProfile $PasswordProfile -PasswordPolicies DisablePasswordExpiration -AccountEnabled $true
        Set-AzureADUserLicense -ObjectId "$UserName@$SecondaryEmail" -AssignedLicenses $LicensesToAssign -ErrorVariable NoSecondaryLicense;
        if ($NoSecondaryLicense) {
            [Microsoft.VisualBasic.Interaction]::MsgBox("No E1 Licenses avaliable", "OKOnly,SystemModal,Critical", "Error")
        }
        start-sleep 5
    }
}

#AD group lists
Function AddADGroups {
    for ($i = 0; $i -lt $ADArray.count; $i++) {
        $ADList = $ADArray[$i]
        Add-ADGroupMember -Identity $ADList -Members $UserName
        }
    $global:FolderAccess = $NetFolders
}
Function enviroADGroups {
    $ADArray = @($enviroADGroups)
    ExchangeConnect
    Import-PSSession $Session -DisableNameChecking
    Add-UnifiedGroupLinks -Identity "Hawaii Operations Team Site" -LinkType Members -Links $UserUPN
    Remove-PSSession $Session
    #adds info to the NewHireSheet
    $NetFolders = "\\\location\qemydocuments\$UserName
            \\\location\QueenEmmaScans\$UserName 
            \\\location\EnvProjects 
            \\\location\NHCDC 
            \\\location\References 
            \\\location\Project Safety"
    AddADGroups
}
Function CPEADGroups {
    $ADArray = @($CPEADGroups)
    #adds info to the NewHireSheet
    $NetFolders = "\\location\cpe\home\$UserName
            \\location\cpe\marketing 
            \\location\cpe\greenwave
            \\location\cpe\projects 
            \\location\cpe\references 
            \\location\cpe\scans\OCEPlotter 
            \\location\cpe\scans\$UserName
            \\location\cpe\cm 
            \\location\CPE\Library"
    AddADGroups
}
Function AccountingADGroups {
    $ADArray = @($AccountingADGroups)
    #adds info to the NewHireSheet
    $NetFolders = "\\\location\qemydocuments\$UserName
            \\\location\QueenEmmaScans\$UserName
            \\\location\NHCDC
            \\\location\EnvProjects
            \\\location\qeaccounting
            \\\location\FinancialReporting"
    AddADGroups
}

    #for Qemydocuments folder and QEScans folder
Function SetupQEFolders {
    if ($Location -eq 'Honolulu') {
        if ($ADCompany -eq "engin") {
            $UFolderArray = @($CPEFolderLocation)
        }
        else {
        $UFolderArray = @($UserFolderLocation)
        }
        for ($i = 0; $i -lt $UFolderArray.count; $i++) {
            $UserFolder = $UFolderArray[$i]
            New-Item -Path "$UserFolder" -Name $UserName -ItemType "directory"
            $inheritanceFlag = [System.Security.AccessControl.InheritanceFlags]"ContainerInherit, ObjectInherit"
            $propagationFlag = [System.Security.AccessControl.PropagationFlags]::None
            $type = [System.Security.AccessControl.AccessControlType]::Allow
        
            $readWrite = [System.Security.AccessControl.FileSystemRights]"Modify"
            $Acl = Get-Acl "$UserFolder\$UserName"
            $Ar = New-Object System.Security.AccessControl.FileSystemAccessRule($UserName, $readWrite, $inheritanceFlag, $propagationFlag, $type)
            $Acl.SetAccessRule($Ar)
            Set-Acl -Path "$UserFolder\$UserName" -AclObject $Acl
        }
    }
}

#Asks if You want to set up AD groups using defaults baked in from AD Group lists
Function SetADGroups {
    #Set default ad group membership
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
    $Button = "YesNo"
    $Ask = "Company: $Company"
    $epromptTitle = "Want to set AD Group Membership now?"
    $LicCheckPrompt = [Microsoft.VisualBasic.Interaction]::MsgBox($Ask, $Button, $epromptTitle)
    Switch ($LicCheckPrompt) {
        Yes {
            ADGroupSelectionPrompt
            ##sets distribution lists more can be added with elseif
            if ($Location -eq 'Honolulu') {
                #no longer needed converted to sharepoint distro
                #Add-ADGroupMember -Identity "Queen Emma Distribution" -Members $UserName
                SetupQEFolders
            }
            elseif ($Location -eq 'Boulder') {
                #no longer needed converted to sharepoint distro
                #Add-DistributionGroupMember -Identity "Boulder Office" -Member $UserName
            }
        }
        No {
            Write-host "Will add to groups later"
        }
    }
}

Function AssignLaptop {
    #check if you want to add the machine name to the setup sheets
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
    $Button = "YesNo"
    $Ask = "Computer to Assign"
    $epromptTitle = "Ready To Assign a Computer"
    $LicCheckPrompt = [Microsoft.VisualBasic.Interaction]::MsgBox($Ask, $Button, $epromptTitle)
    Switch ($LicCheckPrompt) {
        Yes {
            #check if you want to add the machine name to the setup sheets
            [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
            $Button = "YesNo"
            $Ask = "Laptop = yes or Desktop = no"
            $epromptTitle = "Desktop or Laptop"
            $LicCheckPrompt = [Microsoft.VisualBasic.Interaction]::MsgBox($Ask, $Button, $epromptTitle)
            Switch ($LicCheckPrompt) {
                Yes {
                    AssignLaptopPrompt
                }
                No {
                    AssignDesktopPrompt
                }
            }
        }
        No {
            Write-host "Remember to add to sheet later"
        }
    }
}

Function AssignFOB {
    #Set default ad group membership
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
    $Button = "YesNo"
    $Ask = "Ready to Assign a fob?"
    $epromptTitle = "Assign FOB"
    $LicCheckPrompt = [Microsoft.VisualBasic.Interaction]::MsgBox($Ask, $Button, $epromptTitle)
    Switch ($LicCheckPrompt) {
        Yes {
            [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    
            $title = 'FOB Number'
            $msg   = 'Enter FOB Number:'
            
            $global:FOBNumber = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
        }
        No {
            Write-host "Remember to add to sheet later"
        }
    }
}
#Creates a NewHireSheet to add to the ticket outputs to a folder called NewHireSheets
Function NewHireSheet {
    $ScriptRoot = Split-Path $dir
    $date = Get-Date -UFormat "%D"
    New-Item -Path "$ScriptRoot\New Hire sheets" -Name "$PreferedName.txt" -ItemType "file" -Value "NEW STARTER KIT		
            
    Category	Tasks	Requirements
    Access	
            New Employee Name	$PreferedName
        Computer Account	$UserName
        Password	Password
        Password Reset	CTRL+ALT+DEL > Change Password
        Email Account	
            $UserUPN
            $UserName@$SecondaryEmail
        Password	Password
        Password Reset	Synced with Computer account
        Office 365 License	$365LicenseType
        Distribution Lists	N/A
            Folder Accesses:	
            $FolderAccess
    Hardware	
            Laptop	$DesktopName
        Desktop	$LaptopName
        Monitor	N/A
        Docking Station/Peripherals	N/A
        Phone	N/A
        Hotspot	N/A
        Door Key Fob 	$FOBNumber
    Software	
            Adobe Acrobat	Yes
        ArcGIS	N/A
        AutoDesk AutoCAD	N/A
        Dynamics SL Timesheet	Account created as BP Internal User/Project Employee
        Microsoft Visio	N/A
        Microsoft Project	N/A
        Primavera Contractor	N/A
        RMS 3.0	N/A
        VPN	N/A
        Others	N/A
        Supervisor	$Manager
        Completed By	$AdminUser
        Completion Date	$date
    "
}

Function EquipmentSheet {
    $ScriptRoot = Split-Path $dir
    $date = Get-Date -UFormat "%D"
    New-Item -Path "$ScriptRoot\Equipment sheets" -Name "$PreferedName.txt" -ItemType "file" -Value "EmpName    Date    Description serial numberRow1   Description  serial numberRow2
    $PreferedName   $date   $ComputerName   $FOBNumber
    "
}

Function MailConfirmation {
    $mailset = @{
        Subject    = "Set up for $FirstName $LastName Complete"
        Body       = "$FirstName $Lastname was set up in Active directory and office365 
    Primary Email: $UserUPN
    Secondary Email: $UserName@$SecondaryEmail
    Default Password: Password "
        To         = "$Submitter"
        From       = "$AdminUser"
        #Cc = "techteam@gsisg.com"
        SmtpServer = "smtp.office365.com"
        Port       = "587"
        Credential = $Credential
    }
    Send-MailMessage @mailset -UseSsl
}