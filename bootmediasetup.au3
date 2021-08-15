#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Change2CUI=y
#AutoIt3Wrapper_Res_Fileversion=1.0.0.89
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_Language=1033
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

ConsoleWrite("Start "&@ScriptFullPath&@CRLF)
ConsoleWrite("@System="&@SystemDir&@CRLF)
ConsoleWrite("@WorkingDir="&@WorkingDir&@CRLF)

#include <AutoitConstants.au3>

$PathBase = "[DRIVE]:\sources\$OEM$\$$\IT"

If @ScriptDir <> @SystemDir OR NOT @Compiled Then
	Msgbox(0,"Error","Must run compiled and from boot media system directory")
	Exit
Endif

If $CmdLineRaw = "" Then
	RegWrite("HKEY_CURRENT_USER\Console", "ForceV2", "REG_DWORD", "0")
	Run("""" & @ScriptFullPath & """ start")
ElseIf $CmdLineRaw = "start" Then

Else
	Exit
EndIf

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