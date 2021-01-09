#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Change2CUI=n
#AutoIt3Wrapper_Res_Fileversion=1.0.0.26
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_Language=1033
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

ConsoleWrite("Start "&@ScriptFullPath&@CRLF)
ConsoleWrite("@System="&@SystemDir&@CRLF)
ConsoleWrite("@WorkingDir="&@WorkingDir&@CRLF)

$PathBase = "[DRIVE]:\sources\$OEM$\$$\IT"

For $i = 65 To 90
	$Path = StringReplace($PathBase, "[DRIVE]", Chr($i))

	ConsoleWrite("$Path="&$Path&@CRLF)
	If FileExists($Path & "\AutoIt3.exe") Then
		ConsoleWrite("Found"&@CRLF)

		$Execute = $Path & "\AutoIt3.exe /AutoIt3ExecuteScript " & "Main.au3" & " boot-gui"
		ConsoleWrite("$Execute=" & $Execute & @CRLF)
		$Return = Run(@ComSpec & " /c " & $Execute, $Path, @SW_HIDE)
		ConsoleWrite("$Return="&$Return&@CRLF)

		While 1

			Sleep(100)
		WEnd

	Endif

	sleep(100)
Next