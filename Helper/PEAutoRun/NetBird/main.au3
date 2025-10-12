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
#Include "CommonFunctions.au3"

Global $Title = "PENetBird"
Global $LogFullPath = @ScriptFullPath & ".log"
Global $NetBirdConfig = "Settings.ini"
FileChangeDir(@ScriptDir)
_Log($Title)
_Log("@WorkingDir=" & @WorkingDir)
_Log("@ScriptDir=" & @ScriptDir)

Global $IsPE = StringInStr(@WindowsDir, "X:")
If Not $IsPE Then
	_Log("Not running in WinPE.")
	Exit
EndIf

If Not FileExists("netbird.exe") Then
	_Log("Netbird.exe not found.")
	Exit
EndIf

OnAutoItExitRegister("_Exit")

_UpdateStatusBar("NetBird: Wait")

; Wait for network
For $i = 1 To 10
	Ping("8.8.8.8", 1000)
	If Not @error Then ExitLoop
	Sleep(1000)
Next

$ConfigKey = IniRead($NetBirdConfig,"Settings","Key", "")
$KeyURL = IniRead($NetBirdConfig,"Settings","KeyURL", "")
$KeyURLData = ""
If $KeyURL <> "" Then $KeyURLData = BinaryToString(InetRead($KeyURL))

If $KeyURLData <> "" Then
	$NetBirdSetupKey = $KeyURLData
	_Log("Using key from " & $KeyURL)
ElseIf $ConfigKey <> "" Then
	$NetBirdSetupKey = $ConfigKey
	_Log("Using key from " & $NetBirdConfig)
Elseif $NetBirdSetupKey <> "" Then
	_Log("Using built-in key")
Else
	_Log("No key found")
	Exit
EndIf

_Log("$NetBirdSetupKey=" & $NetBirdSetupKey)

$Command = "netbird.exe service install"
_Log("$Command=" & $Command)
$Run = _RunWait($Command)
_Log("$Run=" & $Run & "  @error=" & @error)

_UpdateStatusBar("NetBird: Start")

$Command = "netbird.exe service start"
_Log("$Command=" & $Command)
$Run = _RunWait($Command)
_Log("$Run=" & $Run & "  @error=" & @error)

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
	If StringInStr($Run, "Signal: Connected") Then _UpdateStatusBar("NetBird: Up")
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
