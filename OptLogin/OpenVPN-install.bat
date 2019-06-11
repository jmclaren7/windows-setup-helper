@echo off
echo Administrative permissions required. Detecting permissions...

net session >nul 2>&1
if %errorLevel% == 0 (
	echo Success: Administrative permissions confirmed.
	bitsadmin.exe /transfer "JobName" http://build.openvpn.net/downloads/releases/latest/openvpn-install-latest-stable.exe %~dp0openvpn-install-latest-stable.exe
	%~dp0openvpn-install-latest-stable.exe /S /SELECT_SHORTCUTS=0 /SELECT_SERVICE=1 /SELECT_OPENVPNGUI=0
	robocopy %~dp0 "C:\Program Files\OpenVPN\config" *.ovpn
	net stop OpenVPNService
	net start OpenVPNService 
	sc config "OpenVPNService" start=auto
	pause
	exit

) else (
        echo Failure: Current permissions inadequate.
	pause
	exit
)