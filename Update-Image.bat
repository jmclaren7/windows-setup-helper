@ECHO OFF
Call :Admin
cls

REM =====================================================================
REM The location where the WIM image is mounted, NTLite uses %TEMP%\NLTmpMnt01
set target=%TEMP%\NLTmpMnt01

REM Copy files from repository to mount location
REM The first copy operation uses /mir to remove any left over files from old copy operations
robocopy "%~dp0\Helper" "%target%\Helper" /mir /NFL /NDL
robocopy "%~dp0\Windows" "%target%\Windows" /e /NFL /NDL

REM Copy additional files to the image (Make sure to include required files like Autoit.exe)
robocopy "H:\Windows Images\Additions" "%target%" /e /NFL /NDL
robocopy "H:\Windows Images\Additions-Macrium" "%target%" /e /NFL /NDL

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