<# install the the prerequisites for
    Active Directory (RSAT)
    MSOl connectivity
    Azure connections
    Sets admin credentals for login to MSOL services (saved as .cred) 
    Set new hire form location 
    #>
Function NewHireFilePath {
    #location of the new hire form will usually be in one constant place
    $FormLocation > "$dir\User Settings\NewHireFormLocation.txt"
}
Function Install-ADModule {
    #https://blogs.technet.microsoft.com/ashleymcglone/2016/02/26/install-the-active-directory-powershell-module-on-windows-10/
    [CmdletBinding()]
    Param(
        [switch]$Test = $false
    )

    If ((Get-CimInstance Win32_OperatingSystem).Caption -like "*Windows 10*") {
        Write-Host '---This system is running Windows 10'
    } Else {
        Write-Warning '---This system is not running Windows 10'
        break
    }

    If (Get-HotFix -Id KB2693643 -ErrorAction SilentlyContinue) {

        Write-Host '---RSAT for Windows 10 is already installed'

    } Else {

        Write-Host '---Downloading RSAT for Windows 10'

        If ((Get-CimInstance Win32_ComputerSystem).SystemType -like "x64*") {
            $dl = 'WindowsTH-KB2693643-x64.msu'
        } Else {
            $dl = 'WindowsTH-KB2693643-x86.msu'
        }
        Write-Host "---Hotfix file is $dl"

        Write-Host "---$(Get-Date)"
        $BaseURL = 'https://download.microsoft.com/download/1/D/8/1D8B5022-5477-4B9A-8104-6A71FF9D98AB/'
        $URL = $BaseURL + $dl
        $Destination = Join-Path -Path $HOME -ChildPath "Downloads\$dl"
        $WebClient = New-Object System.Net.WebClient
        $WebClient.DownloadFile($URL,$Destination)
        $WebClient.Dispose()

        Write-Host '---Installing RSAT for Windows 10'
        Write-Host "---$(Get-Date)"
        wusa.exe $Destination /quiet /norestart /log:$home\Documents\RSAT.log

        # wusa.exe returns immediately. Loop until install complete.
        do {
            Write-Host "." -NoNewline
            Start-Sleep -Seconds 3
        } until (Get-HotFix -Id KB2693643 -ErrorAction SilentlyContinue)
        Write-Host "."
        Write-Host "---$(Get-Date)"
    }

    #The latest versions of the RSAT automatically enable all RSAT features
    $RSATCheck = [string[]]$arrayFromFile = Get-Content -Path "$dir\support_files\RSAT.txt"
    for ($i = 0; $i -lt $RSATCheck.count; $i++) {
        $RSName = $RSATCheck[$i]
        If ((Get-WindowsCapability -Name $RSName -Online -ErrorAction SilentlyContinue).State -eq 'Installed') {

            Write-Host "---$RSName PowerShell already enabled"

        } Else {

            Write-Host "---Installing $RSName"
            Add-WindowsCapability -online -Name $RSName

        }
    }

    Write-Host '---ActiveDirectory PowerShell module install complete.'

    # Verify
    If ($Test) {
        Write-Host '---Validating AD PowerShell install'
        dir (Join-Path -Path $HOME -ChildPath Downloads\*msu)
        Get-HotFix -Id KB2693643
        Get-Help Get-ADDomain
        Get-ADDomain
    }
}
Function AzureFirstRunModules {
    $SystemVersion = [System.Environment]::OSVersion.Version | select-object -first 1 Major
    $OSVer = "$SystemVersion"

    $NuG = get-packageprovider -Name NuGet
    if ($NuG -ne $null) {
        Write-Host "NuGet already installed"
    } else {
        Write-Host "Installing NuGet"
        Install-PackageProvider NuGet -force
    }

    $ModName = [string[]]$arrayFromFile = Get-Content -Path "$dir\support_files\PSModules.txt"
    for ($i = 0; $i -lt $ModName.count; $i++) {
        $PSModName = $ModName[$i]
        $ModCheck = get-package -Name $PSModName -ProviderName PowerShellGet
        if ($ModCheck -ne $null) {
            Write-Host "$PSModName already installed"
        } else {
            Write-Host "Installing $PSModName"
            Install-Module $PSModName -force
        }
    }

    if ($OSVer -eq "@{Major=7}"){
        ### Have not tested this on win7 but should work ###
        Write-Host "Installing Local AD"
        Install-Module ActiveDirectory -force

        Write-Host "Installing RSAT-AD-PowerShell"
        Import-Module ServerManagerAdd-WindowsFeature RSAT-AD-PowerShell
        Write-Host "OS 7"
    }
    if ($OSVer -eq "@{Major=10}") {
        Add-WindowsCapability -Online -Name "Rsat.Dns.Tools~~~0.01.0"
        Install-ADModule
    }
}

Function UserCredSet {
    $global:Credential = Get-Credential -Message "Enter a user name and password"

    $Credential.Password | ConvertFrom-SecureString | Out-File "$dir\User Settings\$($Credential.Username).cred" -Force

    $Credential.UserName > $SettingLoc

    $global:SavedAdmin = $Credential.UserName
}

Function OneDriveBackupLoc {
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
    $Button = "YesNo"
    $Ask = "Are you in the Honolulu office?"
    $epromptTitle = "OneDrive Backup Location"
    $AddMorePrompt = [Microsoft.VisualBasic.Interaction]::MsgBox($Ask, $Button, $epromptTitle)
    
        Switch ($AddMorePrompt) {
            Yes {
                Add-Content $SettingLoc "Honolulu"
            }
            No {
                Add-Content $SettingLoc "Boulder"
            } 
        }
}