_ConsoleWrite(@ScriptName)


_ConsoleWrite("@DesktopWidth="&@DesktopWidth)
_ConsoleWrite("@DesktopHeight="&@DesktopHeight)



Func _ConsoleWrite($Data)
	ConsoleWrite($Data & @CRLF)
EndFunc
