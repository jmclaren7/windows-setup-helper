@ECHO OFF
Call :Admin

set image-root=H:\Windows 10 Images\21H2-2022
set additions-root=H:\Windows 10 Images
set target=%TEMP%\NLTmpMnt01

del "%image-root%\Auto-saved*.xml"
del "%image-root%\NTLite.log"

robocopy "%~dp0\IT" "%target%\IT" /mir
robocopy "%~dp0\Windows" "%target%\Windows" /e /xx

robocopy "%additions-root%\Macrium" "%target%" /e /xx
robocopy "%additions-root%\Additions" "%target%" /e /xx


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