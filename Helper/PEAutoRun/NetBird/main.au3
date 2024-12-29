; This script is an example to start the NetBird service
;     Make sure netbird.exe and wintun.dll are in the folder with this script
;     Add the NetBird setup key below
; Be sure you understand the security implications and that you have your setup key and access control configured correctly

$NetBirdSetupKey = ""

; Advanced/Self-hosted options
$NetBirdAdminURL = ""
$NetBirdMgmtURL = ""
$NetBirdPSK = ""

;==============================================================================

#include <CommonFunctions.au3>

Global $Title = "PENetBird"
Global $LogFullPath = StringTrimRight(@ScriptFullPath, 3) & "log"
Global $IsPE = StringInStr(@WindowsDir, "X:")
If Not $IsPE Then Exit
FileChangeDir(@ScriptDir)
If $NetBirdSetupKey = "" Or Not FileExists("netbird.exe") Then Exit

_Log($Title)
_Log("@WorkingDir=" & @WorkingDir)
_Log("@ScriptDir=" & @ScriptDir)

OnAutoItExitRegister("_Exit")

_UpdateStatusBar("NetBird Wait")

; Wait for network
For $i = 1 To 10
	Ping("8.8.8.8", 1000)
	If Not @error Then ExitLoop
	Sleep(1000)
Next

$Command = "netbird.exe service install"
$Run = _RunWait($Command)
_Log("$Run=" & $Run & "  @error=" & @error & "  $Command=" & $Command)

_UpdateStatusBar("NetBird Start")

$Command = "netbird.exe service start"
$Run = _RunWait($Command)
_Log("$Run=" & $Run & "  @error=" & @error & "  $Command=" & $Command)

$Command = "netbird.exe up"
If $NetBirdSetupKey <> "" Then $Command &= " --setup-key " & $NetBirdSetupKey
If $NetBirdAdminURL <> "" Then $Command &= " --admin-url " & $NetBirdAdminURL
If $NetBirdMgmtURL <> "" Then $Command &= " --management-url " & $NetBirdMgmtURL
If $NetBirdPSK <> "" Then $Command &= " --preshared-key " & $NetBirdPSK

$Run = _RunWait($Command)
_Log("$Run=" & $Run & "  @error=" & @error & "  $Command=" & $Command)

; Loop to update status
While 1
	$Command = "netbird.exe status"
	$Run = _RunWait($Command)
	If StringInStr($Run, "Signal: Connected") Then _UpdateStatusBar("NetBird Up")
	Sleep(2000)
WEnd

Exit

;=========== =========== =========== =========== =========== =========== =========== ===========
;=========== =========== =========== =========== =========== =========== =========== ===========

Func _Exit()

EndFunc   ;==>_Exit

Func _UpdateStatusBar($Text)
	Local $Path = @TempDir & "\Helper_Status_" & $Title & ".txt"

	$hFile = FileOpen($Path, 2)
	FileWrite($hFile, $Text)

	FileClose($hFile)
EndFunc   ;==>_UpdateStatusBar
