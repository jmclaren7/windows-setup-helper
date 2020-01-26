@echo off
setlocal
:PROMPT
SET /P AREYOUSURE=Are you sure you want to run sysprep and shutdown the system (Y/[N])?
IF /I "%AREYOUSURE%" NEQ "Y" GOTO END

start C:\Windows\System32\Sysprep\sysprep.exe /oobe /shutdown /generalize /unattend:C:\Windows\IT\autounattend.xml

:END
endlocal


