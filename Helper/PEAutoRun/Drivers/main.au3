; This script loads all driver inf files located in the same folder as this script or it's subdirectories

#include <File.au3>

Global $IsPE = StringInStr(@SystemDir, "X:")


$aFiles = _FileListToArrayRec(@ScriptDir, "*.inf", $FLTAR_FILES, $FLTAR_RECUR, $FLTAR_NOSORT, $FLTAR_FULLPATH)
If @error Then
	ConsoleWrite("No drivers to load or error. @extended = " & @extended & @CRLF)
	Exit
EndIf

For $i = 1 To $aFiles[0]
	$Command = 'drvload "' & $aFiles[$i] & '"'
	ConsoleWrite(@CRLF & $Command)
	If $IsPE Then Run(@ComSpec & " /c " & $Command, @SystemDir, @SW_HIDE)
Next


