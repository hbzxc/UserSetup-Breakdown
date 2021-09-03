@0.7.7.4
mode con:cols=50 lines=15
@echo off
:AdminCheck
echo --------------------------------------------------
echo Checking for Administrative permissions...
echo --------------------------------------------------

net session >nul 2>&1
if %errorLevel% == 0 (
        echo Success: [92mAdministrative permissions confirmed.[0m
	echo --------------------------------------------------
	echo Checking for existing user
	PowerShell -executionpolicy bypass -File "%~dp0bin/SetUpCheck.ps1"
	echo updating logs
	PowerShell -executionpolicy bypass -command "&{. '%~dp0bin/ModuleUpdate.ps1'; LogUpdatesFS}"
	cls
	goto :run
) else (
	setlocal
        echo [93mPlease relaunch as admin[0m
	echo --------------------------------------------------
	timeout /t 20
	exit
)


:run
PowerShell -executionpolicy bypass -command "&{. '%~dp0bin/ModuleUpdate.ps1'; AutoUpdate}"
PowerShell -executionpolicy bypass -command "&{. '%~dp0bin/ModuleUpdate.ps1'; LogUpdatesTS}"
mode con:cols=50 lines=15
echo --------------------------------------------------
echo         [35m------[0m[36mSelection List v0.7.7.4[0m[35m------[0m
echo 1 [46m-[0m               [97mNew User Setup[0m               [46m-[0m 1
echo 2 [46m-[0m            [97mAccount Decommission[0m            [46m-[0m 2
echo 3 [46m-[0m          [97mSet Licence Redundancy[0m            [46m-[0m 3
echo 4 [46m-[0m           [97mDynamics Only Account[0m            [46m-[0m 4
echo 5 [46m-[0m             [97mCopy AD Membership[0m             [46m-[0m 5
echo 6 [46m-[0m             [97mFirst time Setup[0m               [46m-[0m 6
echo 7 [46m-[0m               [97mAdd AD Groups[0m                [46m-[0m 7
echo 8 [46m-[0m             [97mBack up One Drive[0m              [46m-[0m 8
echo 9 [46m-[0m                [97mModule Menu[0m                 [46m-[0m 9
echo 0 [46m-[0m                  [97mUpdate[0m                    [46m-[0m 0
echo ` [46m-[0m                   [97mExit[0m                     [46m-[0m `
echo --------------------------------------------------
set /P c=""
cls
if /I "%c%" equ "1" goto :NewUser
if /I "%c%" equ "New User Setup" goto :NewUser
if /I "%c%" equ "new user setup" goto :NewUser

if /I "%c%" equ "2" goto :AccountDecom
if /I "%c%" equ "Account Decomission" goto :AccountDecom
if /I "%c%" equ "account decomission" goto :AccountDecom

if /I "%c%" equ "3" goto :SetLic
if /I "%c%" equ "Set Licence Redundancy" goto :SetLic
if /I "%c%" equ "set licence redundancy" goto :SetLic

if /I "%c%" equ "4" goto :DynamicsOnly
if /I "%c%" equ "Dynamics Only Account" goto :DynamicsOnly
if /I "%c%" equ "dynamics onlt account" goto :DynamicsOnly

if /I "%c%" equ "5" goto :CopyADMemberships
if /I "%c%" equ "First time Setup" goto :CopyADMemberships
if /I "%c%" equ "first time setup" goto :CopyADMemberships

if /I "%c%" equ "6" goto :FirstSetUp
if /I "%c%" equ "First time Setup" goto :FirstSetUp
if /I "%c%" equ "first time setup" goto :FirstSetUp

if /I "%c%" equ "7" goto :AddAD
if /I "%c%" equ "Add AD Groups" goto :AddAD
if /I "%c%" equ "add ad groups" goto :AddAD

if /I "%c%" equ "8" goto :OneDriveBackUp
if /I "%c%" equ "OneDrive Backup" goto :OneDriveBackUp
if /I "%c%" equ "onedrive backup" goto :OneDriveBackUp
if /I "%c%" equ "OneDrive" goto :OneDriveBackUp
if /I "%c%" equ "onedrive" goto :OneDriveBackUp

if /I "%c%" equ "9" goto :UpdateMenu

if /I "%c%" equ "`" goto :Finished
if /I "%c%" equ "Exit" goto :Finished
if /I "%c%" equ "exit" goto :Finished

if /I "%c%" equ "update" goto:Update
if /I "%c%" equ "0" goto:Update
echo --------------------------------------------------
cls
mode con:cols=50 lines=15
echo [91mChoose 1-9[0m
@timeout /T 2 /nobreak>null
goto :run

:NewUser
mode con:cols=100 lines=30
echo --------------------------------------------------
echo Syncing Workbook
cd "C:\Users\%username%\_User Onboarding"
Start "" /b "C:\Users\%username%\New Hire Form.xlsx"
timeout /T 5 /nobreak >nul
taskkill /IM EXCEL.EXE /F

PowerShell -executionpolicy bypass -File "%~dp0bin/New AD user_form_grab.ps1"
pause
cls
goto:run
exit

:AccountDecom
mode con:cols=100 lines=30
echo --------------------------------------------------
PowerShell -executionpolicy bypass -File "%~dp0bin/Block account.ps1"
pause
cls
goto:run
exit

:SetLic
echo --------------------------------------------------
PowerShell -executionpolicy bypass -File "%~dp0bin/SetLicense.ps1"
pause
cls
goto:run
exit

:DynamicsOnly
mode con:cols=100 lines=30
echo --------------------------------------------------
PowerShell -executionpolicy bypass -File "%~dp0bin/DynamicsOnly.ps1"
pause
cls
goto:run
exit

:FirstSetUp
mode con:cols=100 lines=30
echo --------------------------------------------------
PowerShell -executionpolicy bypass -File "%~dp0bin/FirstTimeSetup.ps1"
pause
cls
goto:run
pause
exit

:AddAD
mode con:cols=100 lines=30
echo --------------------------------------------------
PowerShell -executionpolicy bypass -File "%~dp0bin/AddADGroups.ps1"
pause
cls
goto:run

:CopyADMemberships
mode con:cols=100 lines=30
echo --------------------------------------------------
PowerShell -executionpolicy bypass -File "%~dp0bin/Copy Membership.ps1"
pause
cls
goto:run

:OneDriveBackUp
mode con:cols=100 lines=30
echo --------------------------------------------------
PowerShell -executionpolicy bypass -File "%~dp0bin/OneDriveBackUp.ps1"
pause
cls
goto:run

:Update
PowerShell -executionpolicy bypass -command "&{. '%~dp0bin/ModuleUpdate.ps1'; FileSync}"
pause
cls
goto:run

:UpdateMenu
mode con:cols=50 lines=15
echo --------------------------------------------------
echo             [35m------[0m[36mUpdate Menu[0m[35m------[0m
echo 1 [46m-[0m           [97mView Installed Modules[0m           [46m-[0m 1
echo 2 [46m-[0m              [97mUpdate Modules[0m                [46m-[0m 2
echo 3 [46m-[0m                [97mFile Sync[0m                   [46m-[0m 3
echo 4 [46m-[0m                [97mUninstall[0m                   [46m-[0m 4
echo 5 [46m-[0m                [97mMain Menu[0m                   [46m-[0m 5
echo --------------------------------------------------
set /P c=""
mode con:cols=95 lines=30
cls
if /I "%c%" equ "1" PowerShell -executionpolicy bypass -command "&{. '%~dp0bin/ModuleUpdate.ps1'; ModuleCheck}"
if /I "%c%" equ "2" PowerShell -executionpolicy bypass -command "&{. '%~dp0bin/ModuleUpdate.ps1'; ModuleUpdate}"
if /I "%c%" equ "3" PowerShell -executionpolicy bypass -command "&{. '%~dp0bin/ModuleUpdate.ps1'; FileSync}"
if /I "%c%" equ "4" PowerShell -executionpolicy bypass -command "&{. '%~dp0bin/ModuleUpdate.ps1'; ModuleUninstall}"
if /I "%c%" equ "5" goto:run
pause
cls
mode con:cols=50 lines=15
goto:UpdateMenu

:Finished
exit
