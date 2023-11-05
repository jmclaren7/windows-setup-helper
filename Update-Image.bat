@ECHO OFF
Call :Admin
cls

REM =====================================================================
REM The location where the WIM image is mounted, NTLite generally uses either %TEMP%\NLTmpMnt or %TEMP%\NLTmpMnt01
set target=%TEMP%\NLTmpMnt

REM Delete Helper folder in target if it exists for a clean start
rmdir /s /q "%target%\Helper"

REM Copy files from repository to mount location
REM Copy operations use /e /xx to add files
robocopy "%~dp0\Helper" "%target%\Helper" /e /NFL /NDL
robocopy "%~dp0\Windows" "%target%\Windows" /e /NFL /NDL

REM Copy additional files to the mount location (Make sure to include required files like Autoit.exe)
robocopy "D:\Windows Images\Additions" "%target%" /e /NFL /NDL
robocopy "D:\Windows Images\Additions-Macrium" "%target%" /e /NFL /NDL

REM Remove extra files from image
del "%target%\Auto-saved*.xml"
del "%target%\NTLite.log"


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