@ECHO OFF
Call :Admin
cls

REM =====================================================================
REM The location where the WIM image is mounted
set target=%TEMP%\NLTmpMnt
set packages=C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment\amd64\WinPE_OCs


Dism /Image:"%target%" /Add-Package /PackagePath:"%packages%\WinPE-WMI.cab"
Dism /Image:"%target%" /add-package /packagepath:"%packages%\en-us\WinPE-WMI_en-us.cab"
Dism /Image:"%target%" /Add-Package /PackagePath:"%packages%\WinPE-NetFx.cab"
Dism /Image:"%target%" /add-package /packagepath:"%packages%\en-us\WinPE-NetFx_en-us.cab"
Dism /Image:"%target%" /add-package /packagepath:"%packages%\WinPE-EnhancedStorage.cab"
Dism /Image:"%target%" /add-package /packagepath:"%packages%\en-us\WinPE-EnhancedStorage_en-us.cab"
Dism /Image:"%target%" /add-package /packagepath:"%packages%\WinPE-Scripting.cab"
Dism /Image:"%target%" /add-package /packagepath:"%packages%\en-us\WinPE-Scripting_en-us.cab"
Dism /Image:"%target%" /add-package /packagepath:"%packages%\WinPE-FMAPI.cab"
Dism /Image:"%target%" /add-package /packagepath:"%packages%\WinPE-SecureStartup.cab"
Dism /Image:"%target%" /add-package /packagepath:"%packages%\en-us\WinPE-SecureStartup_en-us.cab"
Dism /Image:"%target%" /add-package /PackagePath:"%packages%\WinPE-PowerShell.cab"
Dism /Image:"%target%" /add-package /PackagePath:"%packages%\en-us\WinPE-PowerShell_en-us.cab"
Dism /Image:"%target%" /add-package /PackagePath:"%packages%\WinPE-StorageWMI.cab"
Dism /Image:"%target%" /add-package /PackagePath:"%packages%\en-us\WinPE-StorageWMI_en-us.cab"
Dism /Image:"%target%" /add-package /PackagePath:"%packages%\WinPE-DismCmdlets.cab"
Dism /Image:"%target%" /add-package /PackagePath:"%packages%\en-us\WinPE-DismCmdlets_en-us.cab"


pause
exit

REM =====================================================================

:Admin
echo Administrative permissions required. Detecting permissions...
net session >nul 2>&1
if %errorLevel% == 0 (
  echo Success: Elevated permissions confirmed, continuing.
) else (
  echo Failure: Not running with elevated permissions, please restart script.
  pause >nul
  exit
)