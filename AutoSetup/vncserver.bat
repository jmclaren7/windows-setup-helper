ping 0.0.0.0 -n 6
wpeutil disablefirewall
reg import "%~dp0vncserver\vncserversettings.reg"
start "vnc" "%~dp0vncserver\vncserver.exe" 