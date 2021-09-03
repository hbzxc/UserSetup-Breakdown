Write-Host "Setting up..."

#credentals for Azure and MSoline
#connect to 365 online services
$ScriptPath = $MyInvocation.MyCommand.Path
$dir = Split-Path $ScriPtpath
$AdminUser = "hblazier@example.com"
$Pword = Get-Content "$dir\User Settings\$AdminUser.cred" | ConvertTo-SecureString

#connects using above stored credental. To create a new credental run the FirstTimeSetup script
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AdminUser, $PWord
Connect-AzureAD -Credential $Credential 
Connect-MsolService -Credential $Credential

#function prompts for AD attributes. Collects base info for account creation
Function FirstNamePrompt {
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

$title = 'First Name'
$msg   = 'Enter First Name:'

$global:FirstName = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
}
Function LastNamePrompt {
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

$title = 'Last Name'
$msg   = 'Enter Last Name:'

$global:LastName = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
}
Function UserNamePrompt {
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

$title = 'UserName'
$msg   = "Enter UserName:

    First Name: $FirstName

    Last Name: $LastName
        "

$global:UserName = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
}
Function JobTitlePrompt {
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

$title = 'Job Title'
$msg   = 'Enter Job Title:'

$global:JobTitle = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
}
Function CompanyPrompt {
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

$title = 'Company'
$msg   = 'Enter Company:'

$global:Company = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)
}

#box lists
Function EmailList {     
[void] $listBox.Items.Add('example.com')
[void] $listBox.Items.Add('company.com')
[void] $listBox.Items.Add('notreal.com')
[void] $listBox.Items.Add('bigorg.com')
[void] $listBox.Items.Add('biglobe.com')
}
Function LocationList {     
[void] $listBox.Items.Add('Boulder')
[void] $listBox.Items.Add('Connecticut')
[void] $listBox.Items.Add('Guam')
[void] $listBox.Items.Add('Honolulu')
[void] $listBox.Items.Add('Kamuela')
[void] $listBox.Items.Add('Maryland')
[void] $listBox.Items.Add('Ohio')
[void] $listBox.Items.Add('Virginia')
[void] $listBox.Items.Add('Washington')
}
Function DivisionList {     
[void] $listBox.Items.Add('Accounting')
[void] $listBox.Items.Add('Business Development')
[void] $listBox.Items.Add('Conference Rooms')
[void] $listBox.Items.Add('Construction')
[void] $listBox.Items.Add('Engin')
[void] $listBox.Items.Add('anotherCompany')
[void] $listBox.Items.Add('anotherCompany-Mec')
[void] $listBox.Items.Add('company-Americas')
[void] $listBox.Items.Add('company-Pacific')
[void] $listBox.Items.Add('Green')
[void] $listBox.Items.Add('HR')
[void] $listBox.Items.Add('MEC')
[void] $listBox.Items.Add('Professional Services')
[void] $listBox.Items.Add('International')
[void] $listBox.Items.Add('Water World')
}
Function 365LicenseList {     
[void] $listBox.Items.Add('EXCHANGESTANDARD')
[void] $listBox.Items.Add('ENTERPRISEPACK')
[void] $listBox.Items.Add('ENTERPRISEPREMIUM')
}

#box prompts
Function EmailPrompt {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Choose an Email'
    $form.Size = New-Object System.Drawing.Size(300,200)
    $form.StartPosition = 'CenterScreen'

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(75,120)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(150,120)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = 'Cancel'
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.Text = 'Please select a email:'
    $form.Controls.Add($label)

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(10,40)
    $listBox.Size = New-Object System.Drawing.Size(260,20)
    $listBox.Height = 80

    EmailList

    $form.Controls.Add($listBox)

    $form.Topmost = $true

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $global:Email = $listBox.SelectedItem
    }
    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
    {
        exit
    }
}
Function LocationPrompt {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Choose a location'
    $form.Size = New-Object System.Drawing.Size(300,200)
    $form.StartPosition = 'CenterScreen'

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(75,120)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(150,120)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = 'Cancel'
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.Text = 'Please select a location:'
    $form.Controls.Add($label)

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(10,40)
    $listBox.Size = New-Object System.Drawing.Size(260,20)
    $listBox.Height = 80

    LocationList

    $form.Controls.Add($listBox)

    $form.Topmost = $true

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $global:location = $listBox.SelectedItem
    }
    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
    {
        exit
    }
}
Function DivisionPrompt {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Choose a Division'
    $form.Size = New-Object System.Drawing.Size(300,200)
    $form.StartPosition = 'CenterScreen'

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(75,120)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(150,120)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = 'Cancel'
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(280,20)
    $label.Text = 'Please select a division:'
    $form.Controls.Add($label)

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(10,40)
    $listBox.Size = New-Object System.Drawing.Size(260,20)
    $listBox.Height = 80

    DivisionList

    $form.Controls.Add($listBox)

    $form.Topmost = $true

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $global:division = $listBox.SelectedItem
    }
    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
    {
        exit
    }
}
Function 365LicensePrompt {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Choose a License'
    $form.Size = New-Object System.Drawing.Size(300,250)
    $form.StartPosition = 'CenterScreen'

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(75,170)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(150,170)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = 'Cancel'
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,10)
    $label.Size = New-Object System.Drawing.Size(280,100)
    $label.Text = '
    Please choose a LIcense:
    
    EXCHANGESTANDARD = E1

    ENTERPRISEPACK = E3

    ENTERPRISEPREMIUM =E5
    '
    $form.Controls.Add($label)

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(10,115)
    $listBox.Size = New-Object System.Drawing.Size(260,20)
    $listBox.Height = 50

    365Licenselist

    $form.Controls.Add($listBox)

    $form.Topmost = $true

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $global:365LicenseType = $listBox.SelectedItem
    }
    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
    {
        exit
    }
}

#wait till synced from AD to 365
Function CheckForSync {
#waiting feedback
$Loading = 0..100
$idx = 0
#Checks if user info is avaliable in MSOnline
$msolUserExist = Get-MsolUser -UserPrincipalName $UserUPN -ErrorAction SilentlyContinue

while ($msolUserExist -eq $Null) {
	
	Write-Progress -Activity "Checking if $UserUPN is Synced to MsOnline" -Status "Waiting..." -PercentComplete $idx
	$idx++
    $msolUserExist = Get-MsolUser -UserPrincipalName $UserUPN -ErrorAction SilentlyContinue
	    if ($idx -ge $Loading.Length)
	    {
		    $idx = 0
	    }
	Start-Sleep 1
    }
}

#Prompt If the user needs a second email and will create an account with the same username. 
#New email can be selected from a list. Will automatically assign the account an E1 license
Function SecondEmailSetUp {
#Second Email Prompt
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
$Button = "YesNo"
$Ask = 'Does the user need an additonal Email?'
$epromptTitle = "Additonal Email or Not?"
$AdditonalEmailPrompt = [Microsoft.VisualBasic.Interaction]::MsgBox($Ask, $Button, $epromptTitle) 

    Switch ($AdditonalEmailPrompt) {
        Yes {
            #creat an additonal 365 account 
            EmailPrompt
            #creat an additonal 365 account 
            $PasswordProfile=New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
            $PasswordProfile.ForceChangePasswordNextLogin = $false
            $PasswordProfile.Password="$DefaultPass"
            $SecondaryLicense = "EXCHANGESTANDARD"
            $PlanName = $SecondaryLicense
            $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
            $License.SkuId = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $PlanName -EQ).SkuID
            $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
            $LicensesToAssign.AddLicenses = $License
        
            New-AzureADUser -DisplayName "$FirstName $LastName" -GivenName "$FirstName" -SurName "$LastName" -UserPrincipalName "$UserName@$Email" -UsageLocation US -MailNickName "$UserName"  -PasswordProfile $PasswordProfile -PasswordPolicies DisablePasswordExpiration -AccountEnabled $true
            Set-AzureADUserLicense -ObjectId "$UserName@$Email" -AssignedLicenses $LicensesToAssign 
            start-sleep 5
            }
        No {
            start-sleep 5
            }
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

#assigns avaliable licenses
Set-AzureADUserLicense -ObjectId $UserUPN -AssignedLicenses $LicensesToAssign
}

#user input
#SetCredentials
FirstNamePrompt
LastNamePrompt
UserNamePrompt
JobTitlePrompt
CompanyPrompt
EmailPrompt
LocationPrompt
DivisionPrompt

$UserUPN="$UserName@$email"
$DefaultPass = ConvertTo-SecureString "Password" -AsPlainText -Force

#creates new ad account
New-ADUser -Name $FirstName$LastName -GivenName $FirstName -Surname $LastName -SamAccountName $UserName -UserPrincipalName "$UserUPN" -title "$JobTitle" -company "$Company" -Path "OU=Users,OU=$division,OU=$location,OU=_Locations,DC=pmp,DC=com" -AccountPassword($DefaultPass) -Enabled $true

365LicensePrompt

SecondEmailSetUp

CheckForSync
start-sleep 5
SetPrimaryLicense

write-host "Set up for $FirstName $LastName complete"
Write-host "Closing in 30 seconds"
Start-Sleep 30
exit