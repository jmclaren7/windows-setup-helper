@ECHO OFF
Call :Admin
cls

REM == Settings You Need To Change =================================
REM The location of the extracted install media (no trailing slash)
set mediapath=D:\Windows Images\11-23H2

REM Extra files that you want to add to the image (no trailing slash, wildcard is added to the end)
set extrafiles=D:\Windows Images\Additions

REM The path where you want to save an ISO file
set outputiso=D:\Windows Images\Windows.iso

REM The index of the boot.wim image you want to modify
REM Usually /Index:2 for an unmodified Windows 10/11 boot.wim
REM Can also be /Name:"Microsoft Windows Setup (amd64)" if you want to use the name
set wimindex=/Name:"Microsoft Windows Setup (amd64)"

REM == Other Paths =================================================
set helperrepo=%~dp0
set sourcewim=%mediapath%\sources\boot.wim
set mountpath=%temp%\WIMMount
set adk=%ProgramFiles(x86)%\Windows Kits\10\Assessment and Deployment Kit


REM == Basic Checks ================================================
if not exist "%mediapath%\" ( echo Media path not found, reconfigure batch file & pause & exit)
if not exist "%sourcewim%" ( echo Boot.wim not found & pause & exit)

REM == Default for toggle options ===================================
set automaticpackages=No
set automaticexport=No

REM == Menu ========================================================
:mainmenu
set pauseafter=true
set returnafter=false
cls
echo.
echo       Media folder = %mediapath%
echo       Mount folder = %mountpath%
echo       Helper files = %helperrepo%
echo       Output ISO = %outputiso%
echo.
echo  1. Mount boot.wim (Windows Image) from media folder
echo  2. Copy Helper files to mounted image
echo  3. Unmount and commit changes to WIM
echo  4. Make ISO from media folder (requires ADK)
echo.
echo  F. Automatically run steps 1,2,3,4 (requires ADK)
echo       (G) Add Packages: %automaticpackages%
echo       (H) Export Overwrite: %automaticexport%
echo.
echo  B. Browse mounted image folder
echo  A. Add packages to mounted image (requires ADK)
echo  X. Unmount, discard changes and cleanup (use if mounted image is stuck)
echo  I. Get image information
echo.
echo  E. Convert install.esd to install.wim
echo  R. Mount and browse install.wim
echo  S. Export and overwrite boot.wim
echo.
echo.  
echo  Enter a selection...
choice /C 1234FGHBAXIERST /N
goto option%errorlevel%

:option1
goto mount
:option2
goto copyfiles
:option3
goto unmountcommit
:option4
goto makeiso
:option5
goto automatic
:option6
goto togglepackages
:option7
goto toggleexport
:option8
goto browsemount
:option9
goto addpackages
:option10
goto unmountdiscard
:option11
goto getinfo
:option12
goto convertinstallesd
:option13
goto mountinstallwim
:option14
goto exportimage
:option15
goto testing


REM == Mount =======================================================
:mount

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

Dism /Image:"%mountpath%" /Add-Package ^
/PackagePath:"%adkpackages%\WinPE-WMI.cab" /PackagePath:"%adkpackages%\en-us\WinPE-WMI_en-us.cab" ^
/PackagePath:"%adkpackages%\WinPE-NetFx.cab" /PackagePath:"%adkpackages%\en-us\WinPE-NetFx_en-us.cab" ^
/PackagePath:"%adkpackages%\WinPE-Scripting.cab" /PackagePath:"%adkpackages%\en-us\WinPE-Scripting_en-us.cab" ^
/PackagePath:"%adkpackages%\WinPE-PowerShell.cab" /PackagePath:"%adkpackages%\en-us\WinPE-PowerShell_en-us.cab" ^
/PackagePath:"%adkpackages%\WinPE-StorageWMI.cab" /PackagePath:"%adkpackages%\en-us\WinPE-StorageWMI_en-us.cab" ^
/PackagePath:"%adkpackages%\WinPE-SecureBootCmdlets.cab" ^
/PackagePath:"%adkpackages%\WinPE-SecureStartup.cab" /PackagePath:"%adkpackages%\en-us\WinPE-SecureStartup_en-us.cab" ^
/PackagePath:"%adkpackages%\WinPE-DismCmdlets.cab" /PackagePath:"%adkpackages%\en-us\WinPE-DismCmdlets_en-us.cab"
/PackagePath:"%adkpackages%\WinPE-EnhancedStorage.cab" /PackagePath:"%adkpackages%\en-us\WinPE-EnhancedStorage_en-us.cab" ^
/PackagePath:"%adkpackages%\WinPE-Dot3Svc.cab" /PackagePath:"%adkpackages%\en-us\WinPE-Dot3Svc_en-us.cab" ^
/PackagePath:"%adkpackages%\WinPE-FMAPI.cab" ^
/PackagePath:"%adkpackages%\WinPE-FontSupport-WinRE.cab ^
/PackagePath:"%adkpackages%\WinPE-PlatformId.cab" ^
/PackagePath:"%adkpackages%\WinPE-WDS-Tools.cab" /PackagePath:"%adkpackages%\en-us\WinPE-WDS-Tools_en-us.cab" ^
/PackagePath:"%adkpackages%\WinPE-HTA.cab" /PackagePath:"%adkpackages%\en-us\WinPE-HTA_en-us.cab" ^
/PackagePath:"%adkpackages%\WinPE-WinReCfg.cab" /PackagePath:"%adkpackages%\en-us\WinPE-WinReCfg_en-us.cab" ^
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
mkdir %mountpath%\Helper

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


REM == Automatic ===================================================
:automatic

echo.
echo [96mRunning steps automatically[0m
echo.

set pauseafter=false
set returnafter=true
call :mount
call :copyfiles
if %automaticpackages%==Yes ( call :addpackages )
call :unmountcommit
if %automaticexport%==Yes ( call :exportimage )
call :makeiso

echo.
pause
goto mainmenu


REM == Get Information =============================================
:getinfo

echo.

Dism /Get-MountedImageInfo 
Dism /Get-ImageInfo /imagefile:"%sourcewim%"

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
:exportimage

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


REM == Testing======================================================
:testing
REM Dism /Delete-Image /ImageFile:boot.wim /Name:"Microsoft Windows Setup (amd64)"
Dism /Image:"%mountpath%" /Cleanup-Image /StartComponentCleanup /ResetBase

echo.
pause
goto mainmenu


REM == Toggle Packages =============================================
:togglepackages

if %automaticpackages%==Yes ( set automaticpackages=No ) else ( set automaticpackages=Yes )

goto mainmenu


REM == Toggle Export ===============================================
:toggleexport

if %automaticexport%==Yes ( set automaticexport=No ) else ( set automaticexport=Yes )

goto mainmenu


REM == Check For Admin =============================================
:Admin
echo Administrative permissions required. Detecting permissions...
net session >nul 2>&1
if %errorlevel% == 0 (
  echo Success: Elevated permissions confirmed, continuing.
  exit /B
) else (
  echo Failure: Not running with elevated permissions, please restart script.
  pause >nul
  exit
)