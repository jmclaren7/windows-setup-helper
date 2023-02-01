_ConsoleWrite(@ScriptName)


_ConsoleWrite("@DesktopWidth="&@DesktopWidth)
_ConsoleWrite("@DesktopHeight="&@DesktopHeight)
_ConsoleWrite("@ComputerName="&@ComputerName)
_ConsoleWrite("@UserName="&@UserName)
_ConsoleWrite("@OSBuild="&@OSBuild)
_ConsoleWrite("@OSVersion="&@OSVersion)
_ConsoleWrite("@OSType="&@OSType)
_ConsoleWrite("@ComputerName="&@ComputerName)
_ConsoleWrite("@ComputerName="&@ComputerName)

Func _ConsoleWrite($Data)
	ConsoleWrite(@ScriptName & ": " & $Data & @CRLF)
EndFunc