@echo off
REM Disables Windows fast startup feature

reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Power" /v HiberbootEnabled /t reg_dword /d 0 /f
powercfg /change standby-timeout-ac 0
powercfg /change standby-timeout-dc 20
powercfg /change monitor-timeout-ac 30
powercfg /change monitor-timeout-dc 10