ConsoleWrite("VNCServer" & @CRLF)
For $i=1 to 6
	Ping ("8.8.8.8", 1000)
	If Not @error Then ExitLoop
	Sleep(1000)
Next

FileChangeDir (@ScriptDir)
ConsoleWrite(@WorkingDir & @CRLF)
ConsoleWrite(@ScriptDir & @CRLF)


ConsoleWrite(Run(@ComSpec & " /c " & 'wpeutil.exe disablefirewall', "", @SW_HIDE) & @CRLF)
ConsoleWrite(Run(@ComSpec & " /c " & 'reg.exe import vncserver\vncserversettings.reg', "", @SW_HIDE) & @CRLF)
ConsoleWrite(Run(@ComSpec & " /c " & 'vncserver\vncserver.exe', "", @SW_HIDE) & @CRLF)