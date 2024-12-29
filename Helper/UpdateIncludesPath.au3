; Add paths to check for includes

$IsPE = StringInStr(@SystemDir, "X:")

If $IsPE Then
	RegWrite("HKEY_CURRENT_USER\Software\AutoIt v3\AutoIt", "Include", "REG_SZ", @ScriptDir & ";" & @ScriptDir & "\IncludeExt\")
EndIf