#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Change2CUI=y
#AutoIt3Wrapper_Res_Fileversion=1.0.0.93
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_Language=1033
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

ConsoleWrite("Start "&@ScriptFullPath&@CRLF)
ConsoleWrite("@System="&@SystemDir&@CRLF)
ConsoleWrite("@WorkingDir="&@WorkingDir&@CRLF)

#include <AutoitConstants.au3>

$PathBase = "[DRIVE]:\sources\$OEM$\$$\IT"

For $i = 64 To 90
	$Path = StringReplace($PathBase, "[DRIVE]", Chr($i))
	If $i = 64 Then $Path = @ScriptDir ; 64 is before "a", use this for testing

	ConsoleWrite("$Path="&$Path&@CRLF)
	If FileExists($Path & "\AutoIt3.exe") Then
		ConsoleWrite("Found"&@CRLF)

		$Execute = """" & $Path & "\AutoIt3.exe"" /AutoIt3ExecuteScript " & "Main.au3" & " boot-gui"
		ConsoleWrite("$Execute=" & $Execute & @CRLF)
		$Return = Run($Execute, $Path, @SW_HIDE, 8)
		ConsoleWrite("$Return="&$Return&@CRLF)

		WinSetState(@ScriptFullPath, "", @SW_MINIMIZE)

		While ProcessExists($Return)
			ConsoleWrite(StdoutRead ($Return))

			Sleep(50)
		WEnd

		Exit

	Endif

	sleep(100)
Next