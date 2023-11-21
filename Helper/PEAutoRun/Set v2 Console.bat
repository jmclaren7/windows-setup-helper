REM The default (V2) console window can't be controlled over VNC for an unknown reason so use V1
reg.exe add HKCU\Console /v ForceV2 /t REG_DWORD /d 0 /f