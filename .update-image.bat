@ECHO OFF
Call :Admin

REM =====================================================================
REM This is the location NTLite uses when mounting an image
set target=%TEMP%\NLTmpMnt01

REM Remove the Helper folder to start fresh
rmdir /s /q "%target%\Helper"

REM Copy files from repository to WinPE image
robocopy "%~dp0\Helper" "%target%\Helper" /mir /NFL /NDL
robocopy "%~dp0\Windows" "%target%\Windows" /e /xx /NFL /NDL

REM Copy additional files to the image (Make sure to include required files like Autoit.exe)
robocopy "H:\Windows Images\Additions" "%target%" /e /xx /NFL /NDL
robocopy "H:\Windows Images\Additions-Macrium" "%target%" /e /xx /NFL /NDL

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
  echo Success: Administrative permissions confirmed, continuing.
) else (
  echo Failure: Not running with elevated permissions, please restart tool.
  pause >nul
  exit
)