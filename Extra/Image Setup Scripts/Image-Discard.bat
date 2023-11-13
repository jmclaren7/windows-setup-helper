@ECHO OFF
Call :Admin
cls

REM =====================================================================
REM The location where the WIM image is mounted
set target=%TEMP%\NLTmpMnt


DISM /Unmount-Image /MountDir:%target% /discard


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