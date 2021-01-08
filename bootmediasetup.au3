#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Change2CUI=n
#AutoIt3Wrapper_Res_Fileversion=1.0.0.18
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_Language=1033
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#pragma compile(AutoItExecuteAllowed, True)

#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>

$Title = "IT Setup Helper - Boot Media Loader"

ConsoleWrite($Title&@CRLF)
ConsoleWrite("@ScriptFullPath="&@ScriptFullPath&@CRLF)
ConsoleWrite("@System="&@SystemDir&@CRLF)
ConsoleWrite("@WorkingDir="&@WorkingDir&@CRLF)

$PathBase = "[DRIVE]:\sources\$OEM$\$$\IT"

For $i = 65 To 90
	$Path = StringReplace($PathBase, "[DRIVE]", Chr($i))

	ConsoleWrite("Checking: "&$Path&@CRLF)
	If FileExists($Path) Then
		ConsoleWrite("Found"&@CRLF)

		;Change Directory
		FileChangeDir ($Path)
		ConsoleWrite("@WorkingDir="&@WorkingDir&@CRLF)

		;Run
		$Execute = @ScriptFullPath & " /AutoIt3ExecuteScript " & "Main.au3" & " boot-gui"
		ConsoleWrite("RunWait: " & $Execute & " (" & $Path & ")" & @CRLF)
		$Return = Run(@ComSpec & " /c " & $Execute, $Path, @SW_HIDE)
		ConsoleWrite("$Return="&$Return&@CRLF)

		While 1

			Sleep(100)
		WEnd


	Endif

	sleep(100)
Next