@ECHO OFF
Call :Admin

set additions-root=H:\Windows 10 Images
set target=%TEMP%\NLTmpMnt01

del "%image-root%\Auto-saved*.xml"
del "%image-root%\NTLite.log"

robocopy "%~dp0\Helper" "%target%\Helper" /mir
robocopy "%~dp0\Windows" "%target%\Windows" /e /xx

robocopy "%additions-root%\Additions" "%target%" /e /xx
robocopy "%additions-root%\Additions-Macrium" "%target%" /e /xx

pause
exit





:Admin
echo Administrative permissions required. Detecting permissions...
net session >nul 2>&1
if %errorLevel% == 0 (
  echo Success: Administrative permissions confirmed, continuing.
) else (
  echo Failure: Not running with elevated permisions, please restart tool.
  pause >nul
  exit
)