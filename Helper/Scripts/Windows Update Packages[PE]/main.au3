#include "File.au3"
#include "WinAPI.au3"
#include "Date.au3"
#include "CommonFunctions.au3"

Local $Drive

Local $aDrivesLetters = DriveGetDrive($DT_ALL)
For $i = 1 To $aDrivesLetters[0]
	$Drive = $aDrivesLetters[$i]

	; If the Windows install is less than 10 minutes old it must be the target
	$TestFile = $aDrivesLetters[$i] & "\Windows\System32\Config\SYSTEM"
	If FileExists($TestFile) And _FileModifiedAge($TestFile) < 600000 Then ExitLoop
Next

$aMSU = _FileListToArray(@ScriptDir, "*.msu", $FLTA_FILES, True)
For $i = 1 To $aMSU[0]
	$Run = 'dism /image:' & $Drive & ' /Add-Package /PackagePath:"' & $aMSU[$i] & '"'
	RunWait(@ComSpec & ' /c ' & $Run, @ScriptDir, @SW_SHOW)

Next