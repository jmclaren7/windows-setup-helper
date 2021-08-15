reg.exe add HKCU\Console /v ForceV2 /t REG_DWORD /d 0 /f
set target=:\sources\$OEM$\$$\IT\WinPE\winpehelper.exe

for %%d in (c d e f g h i j k l m n o p q r s t u v w x y z) do (
	if exist "%%d%target%" (
		start "winpehelper.exe" /WAIT "%%d%target%"
		exit
	)
)