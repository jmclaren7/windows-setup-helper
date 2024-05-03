@ECHO OFF
Call :Admin
cls

REM == Settings You Need To Change =================================
REM Path of the ISO file to extract
set sourceiso=D:\Windows Images\Windows 11 23H2 MCT 2403.iso

REM Directory to extract the ISO to (no trailing slash)
set mediapath=D:\Windows Images\11

REM Directory of extra files to add to the image (no trailing slash)
set extrafiles=D:\Windows Images\Additions

REM Path to the new ISO file
set outputiso=D:\Windows Images\Windows11.iso


REM == Other Settings ==============================================
REM The index of the boot.wim image you want to modify eg: "/Index:2" or "/Name:name"
set wimindex=/Name:"Microsoft Windows Setup (amd64)"

REM == Other Paths =================================================
set helperrepo=%~dp0
set sourcewim=%mediapath%\sources\boot.wim
set mountpath=%temp%\WIMMount
set adk=%ProgramFiles(x86)%\Windows Kits\10\Assessment and Deployment Kit

REM == Defaults for toggle options ==================================
set "auto_extractiso=*"
set "auto_mountwim=*"
set "auto_copyfiles=*"
set "auto_addpackages=*"
set "auto_disabledpi=*"
set "auto_unmountcommit=*"
set "auto_setresolution= "
set "auto_trimimages=*"
set "auto_makeiso=*"
::set "auto_setresolution_detail=Highest"
set "auto_setresolution_detail=1024x768"

REM == Basic Checks ================================================
if not exist "%mediapath%\" ( echo Media path not found, reconfigure batch file & pause & exit)
REM if not exist "%sourcewim%" ( echo Boot.wim not found & pause & exit)
if "%helperrepo:~-1%"=="\" SET helperrepo=%helperrepo:~0,-1%
if not exist "%helperrepo%\Helper\Main.au3" ( echo Main.au3 not found & pause & exit)


REM == Menu ========================================================
:mainmenu
set pauseafter=true
set returnafter=false
cls
echo  Source ISO = %sourceiso%
echo  Media folder = %mediapath%
echo  Output ISO = %outputiso%
echo.
echo  Q^|1. %auto_extractiso%Extract ISO to media folder
echo  W^|2. %auto_mountwim%Mount boot.wim from media folder
echo  E^|3. %auto_copyfiles%Copy Helper files to mounted image
echo  R^|4. %auto_addpackages%Add packages to mounted image (requires ADK)
echo  T^|5. %auto_disabledpi%Apply registry settings to disable DPI scaling in PE
echo  Y^|6. %auto_unmountcommit%Unmount and commit changes to WIM
echo  U^|7. %auto_setresolution%(Not Working)Use Bcdedit to set media boot resolution to: %auto_setresolution_detail% (G to change)
echo  I^|8. %auto_trimimages%Trim boot.wim (Trims other indexes and removes unused files)
echo  O^|9. %auto_makeiso%Make ISO from media folder (requires ADK)
echo.
echo  F. Automatically run the * steps (enter a step # to toggle its inclusion)
echo.
echo  B. Browse mounted image folder (%mountpath%)
echo  X. Discard changes and unmount WIM 
echo  A. Get image information

REM echo  E. Convert install.esd to install.wim
REM echo  R. Mount and browse install.wim
echo.  
echo  Select letter of the function to run...
choice /C 1234567890QWERTYUIOPFBXAG /N
goto option%errorlevel%

:option1
if "%auto_extractiso%"=="*" ( set "auto_extractiso= " ) else ( set "auto_extractiso=*" )
goto mainmenu
:option2
if "%auto_mountwim%"=="*" ( set "auto_mountwim= " ) else ( set "auto_mountwim=*" )
goto mainmenu
:option3
if "%auto_copyfiles%"=="*" ( set "auto_copyfiles= " ) else ( set "auto_copyfiles=*" )
goto mainmenu
:option4
if "%auto_addpackages%"=="*" ( set "auto_addpackages= " ) else ( set "auto_addpackages=*" )
goto mainmenu
:option5
if "%auto_disabledpi%"=="*" ( set "auto_disabledpi= " ) else ( set "auto_disabledpi=*" )
goto mainmenu
:option6
if "%auto_unmountcommit%"=="*" ( set "auto_unmountcommit= " ) else ( set "auto_unmountcommit=*" )
goto mainmenu
:option7
if "%auto_setresolution%"=="*" ( set "auto_setresolution= " ) else ( set "auto_setresolution=*" )
goto mainmenu
:option8
if "%auto_trimimages%"=="*" ( set "auto_trimimages= " ) else ( set "auto_trimimages=*" )
goto mainmenu
:option9
if "%auto_makeiso%"=="*" ( set "auto_makeiso= " ) else ( set "auto_makeiso=*" )
goto mainmenu
:option10
REM Not used
goto mainmenu

:: Q
:option11
goto extractiso
:: W
:option12
goto mountwim
:: E
:option13
goto copyfiles
:: R
:option14
goto addpackages
:: T
:option15
goto disabledpi
:: Y
:option16
goto setresolution
:: U
:option17
goto unmountcommit
:: I
:option18
goto trimimages
:: O
:option19
goto makeiso
:: P
:option20
REM Not used
goto mainmenu
:: F
:option21
goto automatic
:: B
:option22
goto browsemount
:: X
:option23
goto unmountdiscard
:: A
:option24
goto getinfo
:: G
:option25
goto toggle_resolution_detail

REM == Toggles ======================================================
:toggle_resolution_detail
if %auto_setresolution_detail%==Highest ( set auto_setresolution_detail=800x600 & goto mainmenu )
if %auto_setresolution_detail%==800x600 ( set auto_setresolution_detail=1024x600 & goto mainmenu )
if %auto_setresolution_detail%==1024x600 ( set auto_setresolution_detail=1024x768 & goto mainmenu )
if %auto_setresolution_detail%==1024x768 ( set auto_setresolution_detail=Highest & goto mainmenu )
goto mainmenu

REM == Automatic ===================================================
:automatic

echo.
echo [96mRunning steps automatically[0m
echo.

set pauseafter=false
set returnafter=true
if "%auto_extractiso%"=="*" ( call :extractiso )
if "%auto_mountwim%"=="*" ( call :mountwim )
if "%auto_copyfiles%"=="*" ( call :copyfiles )
if "%auto_addpackages%"=="*" ( call :addpackages )
if "%auto_disabledpi%"=="*" ( call :disabledpi )
if "%auto_unmountcommit%"=="*" ( call :unmountcommit )
if "%auto_setresolution%"=="*" ( call :setresolution )
if "%auto_trimimages%"=="*" ( call :trimimages )
if "%auto_makeiso%"=="*" ( call :makeiso )

echo.
pause
goto mainmenu


REM == Extract From ISO ============================================
:extractiso

echo.
echo [96mExtracting ISO to media folder[0m
echo.

rmdir /s /q "%mediapath%"
mkdir "%mediapath%"
"%~dp0Helper\Tools\7-Zip\7z.exe" x -y -o"%mediapath%" "%sourceiso%"

echo.
if %pauseafter%==true ( pause )
if %returnafter%==true ( exit /B )
goto mainmenu


REM == Mount =======================================================
:mountwim

echo.
echo [96mMounting image to: %mountpath%[0m
echo.

mkdir %mountpath%
Dism /Mount-image /ImageFile:"%sourcewim%" %wimindex% /MountDir:"%mountpath%" /Optimize

if %errorlevel% NEQ 0 ( echo [91mError mounting image, aborting[0m && pause && goto mainmenu )

echo.
if %pauseafter%==true ( pause )
if %returnafter%==true ( exit /B )
goto mainmenu


REM == Browse Mount ===============================================
:browsemount

start explorer "%mountpath%"

echo.
if %returnafter%==true ( exit /B )
goto mainmenu


REM == Add Packages ================================================
:addpackages

echo.
echo [96mAdding packages to mounted image[0m
echo.

set adkpackages=%adk%\Windows Preinstallation Environment\amd64\WinPE_OCs

echo Packages From: %adkpackages%

Dism /Image:"%mountpath%" /Add-Package ^
/PackagePath:"%adkpackages%\WinPE-WMI.cab" /PackagePath:"%adkpackages%\en-us\WinPE-WMI_en-us.cab" ^
/PackagePath:"%adkpackages%\WinPE-NetFx.cab" /PackagePath:"%adkpackages%\en-us\WinPE-NetFx_en-us.cab" ^
/PackagePath:"%adkpackages%\WinPE-Scripting.cab" /PackagePath:"%adkpackages%\en-us\WinPE-Scripting_en-us.cab" ^
/PackagePath:"%adkpackages%\WinPE-PowerShell.cab" /PackagePath:"%adkpackages%\en-us\WinPE-PowerShell_en-us.cab" ^
/PackagePath:"%adkpackages%\WinPE-StorageWMI.cab" /PackagePath:"%adkpackages%\en-us\WinPE-StorageWMI_en-us.cab" ^
/PackagePath:"%adkpackages%\WinPE-SecureBootCmdlets.cab" ^
/PackagePath:"%adkpackages%\WinPE-SecureStartup.cab" /PackagePath:"%adkpackages%\en-us\WinPE-SecureStartup_en-us.cab" ^
/PackagePath:"%adkpackages%\WinPE-DismCmdlets.cab" /PackagePath:"%adkpackages%\en-us\WinPE-DismCmdlets_en-us.cab" ^
/PackagePath:"%adkpackages%\WinPE-EnhancedStorage.cab" /PackagePath:"%adkpackages%\en-us\WinPE-EnhancedStorage_en-us.cab" ^
/PackagePath:"%adkpackages%\WinPE-Dot3Svc.cab" /PackagePath:"%adkpackages%\en-us\WinPE-Dot3Svc_en-us.cab" ^
/PackagePath:"%adkpackages%\WinPE-FMAPI.cab" ^
/PackagePath:"%adkpackages%\WinPE-FontSupport-WinRE.cab" ^
/PackagePath:"%adkpackages%\WinPE-PlatformId.cab" ^
/PackagePath:"%adkpackages%\WinPE-WDS-Tools.cab" /PackagePath:"%adkpackages%\en-us\WinPE-WDS-Tools_en-us.cab" ^
/PackagePath:"%adkpackages%\WinPE-HTA.cab" /PackagePath:"%adkpackages%\en-us\WinPE-HTA_en-us.cab" ^
/PackagePath:"%adkpackages%\WinPE-WinReCfg.cab" /PackagePath:"%adkpackages%\en-us\WinPE-WinReCfg_en-us.cab"
REM /PackagePath:"%adkpackages%\WinPE-Setup.cab" /PackagePath:"%adkpackages%\en-us\WinPE-Setup_en-us.cab" ^
REM /PackagePath:"%adkpackages%\WinPE-Setup-Client.cab" /PackagePath:"%adkpackages%\en-us\WinPE-Setup-Client_en-us.cab"

if %errorlevel% NEQ 0 ( echo [91mError adding packages, aborting[0m && pause && goto mainmenu )

echo.
if %pauseafter%==true ( pause )
if %returnafter%==true ( exit /B )
goto mainmenu


REM == Copy Files ==================================================
:copyfiles

echo.
echo [96mCopying files to mounted image[0m  
echo.

REM Delete Helper folder in mount path if it exists
rmdir /s /q "%mountpath%\Helper"
mkdir "%mountpath%\Helper"

REM Copy files from repository to mounted image
echo Copying "%helperrepo%\Helper"
xcopy /y /e /q "%helperrepo%\Helper" "%mountpath%\Helper"
echo.

echo Copying "%helperrepo%\Windows"
xcopy /y /e /q "%helperrepo%\Windows" "%mountpath%\Windows"
echo.

REM Copy extra files to the mounted image
for /D %%A in ("%extrafiles%*") do (
   echo Copying "%%~fA"
   xcopy /y /e /q "%%~fA" "%mountpath%"
   echo.
)

REM Remove extra files from image
echo Removing logs and extra files
del "%mountpath%\Auto-saved*.xml"
del "%mountpath%\*.log"
del "%mountpath%\Helper\Logon\*.log"

echo.
if %pauseafter%==true ( pause )
if %returnafter%==true ( exit /B )
goto mainmenu


REM == Unmount Commit ==============================================
:unmountcommit

echo.
echo [96mUnmounting and committing changes[0m
echo [93mSave and close any open files that are in the mount path[0m
echo [93mClose any open file explorer windows that are accessing the mount path[0m
echo Continuing will commit any changes to the WIM and attempt to umount the image
echo.
if %pauseafter%==true ( pause ) 

Dism /Unmount-Image /MountDir:"%mountpath%" /commit
if %errorlevel% NEQ 0 ( echo [91mUnmount error, aborting[0m && pause && goto mainmenu )

rmdir /s /q "%mountpath%"

echo.
if %pauseafter%==true ( pause )
if %returnafter%==true ( exit /B )
goto mainmenu


REM == Unmount Discard =============================================
:unmountdiscard

echo.
echo [96mUnmounting and discarding changes[0m
echo.

Dism /Unmount-Image /MountDir:"%mountpath%" /Discard
Dism /Cleanup-Mountpoints 
rmdir /s /q "%mountpath%"

echo.
if %pauseafter%==true ( pause )
if %returnafter%==true ( exit /B )
goto mainmenu


REM == Make ISO ====================================================
:makeiso

echo.
echo [96mCreating ISO[0m
echo Input: "%mediapath%" 
echo Output: "%outputiso%"
echo.

set oscdimg=%adk%\Deployment Tools\amd64\Oscdimg

set BOOTDATA=1#pEF,e,b"%mediapath%\efi\microsoft\boot\efisys.bin"
if exist "%mediapath%\boot\etfsboot.com" (
  set BOOTDATA=2#p0,e,b"%mediapath%\boot\etfsboot.com"#pEF,e,b"%mediapath%\efi\microsoft\boot\efisys.bin"
)

del /F /Q "%outputiso%"
if %errorlevel% NEQ 0 ( echo [91mFailed to delete "%DEST%".[0m && pause && goto mainmenu )

"%oscdimg%\oscdimg" -bootdata:%BOOTDATA% -u1 -udfver102 "%mediapath%" "%outputiso%"
if %errorlevel% NEQ 0 ( echo [91mFailed to create ISO.[0m && pause && goto mainmenu )

echo.
if %pauseafter%==true ( pause )
if %returnafter%==true ( exit /B )
goto mainmenu


REM == Get Information =============================================
:getinfo

echo.

Dism /Get-MountedImageInfo 
Dism /Get-ImageInfo /imagefile:"%sourcewim%"
Dism /Get-ImageInfo /imagefile:"%sourcewim%" %wimindex%

echo.
pause 
goto mainmenu


REM == Convert Install.esd to WIM ====================================
:convertinstallesd

echo.
echo [96mConverting ESD to WIM[0m
echo.

Dism /Export-Image /SourceImageFile:"%mediapath%\sources\install.esd" /SourceIndex:1 /DestinationImageFile:"%mediapath%\sources\install.wim" /Compress:Max
if %errorlevel% NEQ 0 ( echo [91mError converting install.esd, aborting[0m && pause && goto mainmenu )

echo.
pause 
goto mainmenu


REM == Mount Install.wim =============================================
:mountinstallwim

echo.
echo [96mExtracting WinRE[0m
echo.

mkdir %mountpath%
Dism /Mount-image /ReadOnly /imagefile:"%mediapath%\sources\install.wim" /Index:1 /MountDir:"%mountpath%" /Optimize /ReadOnly
if %errorlevel% NEQ 0 ( echo [91mError mounting install.wim, aborting[0m && pause && goto mainmenu )

start explorer "%mediapath%\sources"

echo.
pause
goto mainmenu


REM == Export Image ================================================
:trimimages

set "sourcewimindex=%wimindex:/=/Source%"

echo.
echo [96mExporting boot.wim image using  %sourcewimindex%[0m
echo.

Dism /Export-Image /SourceImageFile:"%sourcewim%" %sourcewimindex% /DestinationImageFile:"%mediapath%\sources\boot2.wim" /Compress:Max
if %errorlevel% NEQ 0 ( echo [91mError exporting, aborting[0m && pause && goto mainmenu )

move /Y "%mediapath%\sources\boot2.wim" "%sourcewim%"
if %errorlevel% NEQ 0 ( echo [91mError overwriting boot.wim, aborting[0m && pause && goto mainmenu )

echo.
if %pauseafter%==true ( pause )
if %returnafter%==true ( exit /B )
goto mainmenu


REM == Set Resolution ===============================================
:setresolution

echo.
echo [96mSet Resolution[0m
echo.
if "%auto_setresolution_detail%"=="Highest" ( 
  bcdedit.exe /store "%mediapath%\boot\bcd" /set {default} highestmode on 
) else (
  echo other
  bcdedit.exe /store "%mediapath%\boot\bcd" /set {default} graphicsresolution %auto_setresolution_detail%
)
echo.
if %pauseafter%==true ( pause )
if %returnafter%==true ( exit /B )
goto mainmenu


REM == Disable DPI Scaling =============================================
:disabledpi

echo.
echo [96mEdit Registry[0m
echo.

reg load HKLM\_WinPE_Default %mountpath%\Windows\System32\config\default
reg add "HKLM\_WinPE_Default\Control Panel\Desktop" /v LogPixels /t REG_DWORD /d 96 /f
reg add "HKLM\_WinPE_Default\Control Panel\Desktop" /v Win8DpiScaling /t REG_DWORD /d 0x00000001 /f
reg add "HKLM\_WinPE_Default\Control Panel\Desktop" /v DpiScalingVer /t REG_DWORD /d 0x00001018 /f
reg unload HKLM\_WinPE_Default

echo.
if %pauseafter%==true ( pause )
if %returnafter%==true ( exit /B )
goto mainmenu


REM == Check For Admin =============================================
:Admin
echo Administrative permissions required. Detecting permissions...
net session >nul 2>&1
if %errorlevel% == 0 (
  echo Success: Elevated permissions confirmed, continuing.
  exit /B
) else (
  echo.
  echo Error: Not running with elevated permissions, please restart script.
  pause >nul
  exit
)