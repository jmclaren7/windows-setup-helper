While 1
	_ConsoleWrite(WinGetTitle("[Active]"))


	Sleep(2000)
Wend

Func _ConsoleWrite($Data)
	ConsoleWrite(@ScriptName & ": " & $Data & @CRLF)
EndFunc