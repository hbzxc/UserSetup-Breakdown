$ScriptPath = $PSCommandPath
$dir = Split-Path $ScriptPath
$top = Split-Path $dir

Import-Module -Name "$dir\support_files\MSOLModule.psm1"
AdminCheck

Function ModuleCheck{
    get-package -ProviderName PowerShellGet
    Get-WindowsCapability -Name RSAT* -Online | Select-Object -Property DisplayName, State | where State -eq 'Installed'
}

Function ModuleUpdate {
    Write-Host "Updating NuGet"
    Install-PackageProvider NuGet -force

    $ModName = [string[]]$arrayFromFile = Get-Content -Path "$dir\support_files\PSModules.txt"
    for ($i = 0; $i -lt $ModName.count; $i++) {
        $PSModName = $ModName[$i]
        Write-Host "Updating $PSModName"
        Install-Module $PSModName -force
    }

    $RSATCheck = [string[]]$arrayFromFile = Get-Content -Path "$dir\support_files\RSAT.txt"
    for ($i = 0; $i -lt $RSATCheck.count; $i++) {
        $RSName = $RSATCheck[$i]
        Write-Host "---Installing $RSName"
        Add-WindowsCapability -online -Name $RSName
    }
}

Function ModuleUninstall {
    $ModName = [string[]]$arrayFromFile = Get-Content -Path "$dir\support_files\PSModules.txt"
    for ($i = 0; $i -lt $ModName.count; $i++) {
        $PSModName = $ModName[$i]
        Write-Host "Uninstalling $PSModName module"
        get-package -Name $PSModName -ProviderName PowerShellGet | Uninstall-Package -force
    }

    $RSATCheck = [string[]]$arrayFromFile = Get-Content -Path "$dir\support_files\RSAT.txt"
    for ($i = 0; $i -lt $RSATCheck.count; $i++) {
        $RSName = $RSATCheck[$i]
        Write-Host "---Installing $RSName"
        Remove-WindowsCapability -online -Name $RSName
    }
}

Function FileSync{
    xcopy "\\10.100.7.15\it\_Scripts\_UserCreateDepart_Tool\UserSetup-Breakdown" $top /s /d /e /f /y
}

Function LogUpdatesFS{
    xcopy "\\10.100.7.15\it\_Scripts\_UserCreateDepart_Tool\UserSetup-Breakdown\Logs" "$top\Logs" /s /d /e /f /y
}

Function LogUpdatesTS{
    xcopy "$top\Logs" "\\10.100.7.15\it\_Scripts\_UserCreateDepart_Tool\UserSetup-Breakdown\Logs" /s /d /e /f /y
}

Function AutoUpdate{
    $SerVerRaw = get-content "\\10.100.7.15\it\_Scripts\_UserCreateDepart_Tool\UserSetup-Breakdown\_RunAccount.bat"
    $SerVer = $SerVerRaw[0]
    $SerVer = $SerVer -replace "\.|\@",""

    $LocalVerRaw = get-content "$top\_RunAccount.bat"
    $LocalVer = $LocalVerRaw[0]
    $LocalVer = $LocalVer -replace "\.|\@",""

    if ($LocalVer -lt $SerVer){
        [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
        $Button = "YesNo"
        $Ask = "Update Available

        Local  Version  : $LocalVer
        Server Version  : $SerVer"
        $epromptTitle = "Update to Version $SerVer"
        $LicCheckPrompt = [Microsoft.VisualBasic.Interaction]::MsgBox($Ask, $Button, $epromptTitle)
        Switch ($LicCheckPrompt) {
            Yes {
                xcopy "\\10.100.7.15\it\_Scripts\_UserCreateDepart_Tool\UserSetup-Breakdown" $top /s /d /e /f /y
            }
            No {
                exit
            }
        }
    }
}