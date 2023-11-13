@ECHO OFF
Call :Admin
cls

REM =====================================================================
REM The location where the WIM image is mounted
set target=%TEMP%\NLTmpMnt
set source=D:\Windows Images\11-23H2\sources\boot.wim

mkdir %target%
DISM /Mount-image /imagefile:"%source%" /Index:1 /MountDir:%target% /optimize


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