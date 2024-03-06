@ECHO OFF
Call :Admin
cls

REM == Settings You Need To Change =================================
REM The location of the extracted install media (no trailing slash)
set mediapath=D:\Windows Images\11-23H2

REM Extra files that you want to add to the image (no trailing slash)
set extrafiles1=D:\Windows Images\Additions
set extrafiles2=D:\Windows Images\Additions-Macrium

REM The path where you want to save an ISO file
set outputiso=D:\Windows Images\Windows.iso

REM The index of the boot.wim image you want to modify, usually 2 for an unmodified Windows 10/11 boot.wim
set wimindex=2

REM == Other Paths =================================================
set helperrepo=%~dp0
set sourcewim=%mediapath%\sources\boot.wim
set mountpath=%temp%\WIMPath
set adk=%ProgramFiles(x86)%\Windows Kits\10\Assessment and Deployment Kit


REM == Basic Checks ================================================
if not exist "%mediapath%\" ( echo Media path not found, reconfigure batch file & pause & exit)
if not exist "%sourcewim%" ( echo Boot.wim not found & pause & exit)


REM == Menu ========================================================
:mainmenu
set pauseafter=true
cls
echo.
echo       Media folder = %mediapath%
echo       Mount folder = %mountpath%
echo       Helper files = %helperrepo%
echo       Output ISO = %outputiso%
echo.
echo  1. Mount boot.wim (Windows Image) from media folder
echo  2. Open mounted image in explorer
echo  3. Add packages to mounted image (requires ADK)
echo  4. Copy Helper files to mounted image
echo  5. Unmount and commit changes to WIM
echo  6. Unmount, discard changes and cleanup (use if mounted image is stuck)
echo  7. Make ISO from media folder (requires ADK)
echo  8. Get image information
echo.
echo  F. Automatically run steps 1,4,5,7 (requires ADK)
echo.  
echo  Enter a selection...
choice /C 1234567F8 /N
goto mainmenu%errorlevel%


REM == Mount =======================================================
:mainmenu1

echo.
echo Mounting image to: %mountpath%
echo.

mkdir %mountpath%
Dism /Mount-image /imagefile:"%sourcewim%" /Index:%wimindex% /MountDir:"%mountpath%" /optimize

echo.
if %pauseafter%==true ( pause ) else ( exit /B ) 
goto mainmenu


REM == Explore =====================================================
:mainmenu2

start explorer "%mountpath%"

goto mainmenu


REM == Packages ====================================================
:mainmenu3

echo.
echo Adding packages to image
echo.

set adkpackages=%adk%\Windows Preinstallation Environment\amd64\WinPE_OCs

Dism /Image:"%mountpath%" /Add-Package /PackagePath:"%adkpackages%\WinPE-WMI.cab"
Dism /Image:"%mountpath%" /add-package /packagepath:"%adkpackages%\en-us\WinPE-WMI_en-us.cab"
Dism /Image:"%mountpath%" /Add-Package /PackagePath:"%adkpackages%\WinPE-NetFx.cab"
Dism /Image:"%mountpath%" /add-package /packagepath:"%adkpackages%\en-us\WinPE-NetFx_en-us.cab"
Dism /Image:"%mountpath%" /add-package /packagepath:"%adkpackages%\WinPE-EnhancedStorage.cab"
Dism /Image:"%mountpath%" /add-package /packagepath:"%adkpackages%\en-us\WinPE-EnhancedStorage_en-us.cab"
Dism /Image:"%mountpath%" /add-package /packagepath:"%adkpackages%\WinPE-Scripting.cab"
Dism /Image:"%mountpath%" /add-package /packagepath:"%adkpackages%\en-us\WinPE-Scripting_en-us.cab"
Dism /Image:"%mountpath%" /add-package /packagepath:"%adkpackages%\WinPE-FMAPI.cab"
Dism /Image:"%mountpath%" /add-package /packagepath:"%adkpackages%\WinPE-SecureStartup.cab"
Dism /Image:"%mountpath%" /add-package /packagepath:"%adkpackages%\en-us\WinPE-SecureStartup_en-us.cab"
Dism /Image:"%mountpath%" /add-package /PackagePath:"%adkpackages%\WinPE-PowerShell.cab"
Dism /Image:"%mountpath%" /add-package /PackagePath:"%adkpackages%\en-us\WinPE-PowerShell_en-us.cab"
Dism /Image:"%mountpath%" /add-package /PackagePath:"%adkpackages%\WinPE-StorageWMI.cab"
Dism /Image:"%mountpath%" /add-package /PackagePath:"%adkpackages%\en-us\WinPE-StorageWMI_en-us.cab"
Dism /Image:"%mountpath%" /add-package /PackagePath:"%adkpackages%\WinPE-DismCmdlets.cab"
Dism /Image:"%mountpath%" /add-package /PackagePath:"%adkpackages%\en-us\WinPE-DismCmdlets_en-us.cab"

echo.
pause
goto mainmenu


REM == Copy Files ==================================================
:mainmenu4

echo.
echo Copying files to image
echo.

REM Delete Helper folder in mount path if it exists
rmdir /s /q "%mountpath%\Helper"

REM Copy files from repository to mounted image
robocopy "%helperrepo%\Helper" "%mountpath%\Helper" /e /NFL /NDL
robocopy "%helperrepo%\Windows" "%mountpath%\Windows" /e /NFL /NDL

REM Copy extra files to the mounted image
if exist "%extrafiles1%\" ( robocopy "%extrafiles1%" "%mountpath%" /e /NFL /NDL )
if exist "%extrafiles2%\" ( robocopy "%extrafiles2%" "%mountpath%" /e /NFL /NDL )

REM Remove extra files from image
del "%mountpath%\Auto-saved*.xml"
del "%mountpath%\*.log"
del "%mountpath%\Helper\Logon\*.log"

echo.
if %pauseafter%==true ( pause ) else ( exit /B ) 
goto mainmenu


REM == Unmount Commit ==============================================
:mainmenu5

echo.
echo Save and close any open files that are in the mount path
echo Close any open file explorer windows that are accessing the mount path
echo Continuing will commit any changes to the WIM and attempt to umount the image
echo.
if %pauseafter%==true ( pause ) 

DISM /Unmount-Image /MountDir:"%mountpath%" /commit

echo.
if %pauseafter%==true ( pause ) else ( exit /B ) 
goto mainmenu


REM == Unmount Discard =============================================
:mainmenu6

DISM /Unmount-Image /MountDir:"%mountpath%" /discard
DISM /Cleanup-WIM

echo.
pause
goto mainmenu


REM == Make ISO ====================================================
:mainmenu7

echo.
echo Making ISO
echo Input: "%mediapath%" 
echo Output: "%outputiso%"
echo.

set oscdimg=%adk%\Deployment Tools\amd64\Oscdimg

set BOOTDATA=1#pEF,e,b"%mediapath%\efi\microsoft\boot\efisys.bin"
if exist "%mediapath%\boot\etfsboot.com" (
  set BOOTDATA=2#p0,e,b"%mediapath%\boot\etfsboot.com"#pEF,e,b"%mediapath%\efi\microsoft\boot\efisys.bin"
)

del /F /Q "%outputiso%"
if errorlevel 1 (
  echo ERROR: Failed to delete "%DEST%".
  pause
  goto mainmenu
)

"%oscdimg%\oscdimg" -bootdata:%BOOTDATA% -u1 -udfver102 "%mediapath%" "%outputiso%"
if errorlevel 1 (
  echo ERROR: Failed to create file.
)

echo.
pause
goto mainmenu


REM == Automatic ===================================================
:mainmenu8

set pauseafter=false
call :mainmenu1
call :mainmenu4
call :mainmenu5
call :mainmenu7

echo.
pause
goto mainmenu

REM == Get Information =============================================
:mainmenu9

echo.
Dism /Get-MountedImageInfo 
Dism /Get-ImageInfo /imagefile:"%sourcewim%"

echo.
pause 
goto mainmenu



REM == Check For Admin =============================================
:Admin
echo Administrative permissions required. Detecting permissions...
net session >nul 2>&1
if %errorLevel% == 0 (
  echo Success: Elevated permissions confirmed, continuing.
  exit /B
) else (
  echo Failure: Not running with elevated permissions, please restart script.
  pause >nul
  exit
)