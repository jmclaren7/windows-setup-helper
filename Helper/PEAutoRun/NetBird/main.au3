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

Global $Title = "PENetBird"
Global $IsPE = StringInStr(@WindowsDir, "X:")
If Not $IsPE Then Exit
FileChangeDir(@ScriptDir)
If $NetBirdSetupKey = "" Or Not FileExists("netbird.exe") Then Exit

_Log($Title)
_Log("@WorkingDir=" & @WorkingDir)
_Log("@ScriptDir=" & @ScriptDir)

OnAutoItExitRegister("_Exit")

; Wait for network
For $i = 1 To 10
	Ping("8.8.8.8", 1000)
	If Not @error Then ExitLoop
	Sleep(1000)
Next

$Command = "netbird.exe service install"
$Run = _RunWait($Command)
_Log("$Run=" & $Run & "  @error=" & @error & "  $Command=" & $Command)

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
	If StringInStr($Run, "Signal: Connected") Then _UpdateStatusBar("NetBird Connected")
	Sleep(3000)
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

Func _Log($Data)
	ConsoleWrite(@CRLF & $Title & ": " & $Data)
	FileWriteLine($Title & "_Log.txt", $Data)
EndFunc   ;==>_Log

Func _RunWait($sProgram, $Working = "", $Show = @SW_HIDE, $Opt = 8, $Live = False)
	Local $sData, $iPid

	$iPid = Run($sProgram, $Working, $Show, $Opt)
	If @error Then
		_Log("_RunWait: Couldn't Run " & $sProgram)
		Return SetError(1, 0, 0)
	EndIf

	$sData = _ProcessWaitClose($iPid, $Live)

	Return SetError(0, $iPid, $sData)
EndFunc   ;==>_RunWait

Func _ProcessWaitClose($iPid, $Live = False, $Diag = False)
	Local $sData, $sStdRead

	While 1
		$sStdRead = StdoutRead($iPid)
		If @error Or $sStdRead = "" Then StderrRead($iPid)
		If @error And Not ProcessExists($iPid) Then ExitLoop
		$sStdRead = StringReplace($sStdRead, @CR & @LF & @CR & @LF, @CR & @LF)

		If $Diag Then
			$sStdRead = StringReplace($sStdRead, @CRLF, "_@CRLF")
			$sStdRead = StringReplace($sStdRead, @CR, "@CR" & @CR)
			$sStdRead = StringReplace($sStdRead, @LF, "@LF" & @LF)
			$sStdRead = StringReplace($sStdRead, "_@CRLF", "@CRLF" & @CRLF)
		EndIf

		If $sStdRead <> @CRLF Then
			$sData &= $sStdRead
			If $Live And $sStdRead <> "" Then
				If StringRight($sStdRead, 2) = @CRLF Then $sStdRead = StringTrimRight($sStdRead, 2)
				If StringRight($sStdRead, 1) = @LF Then $sStdRead = StringTrimRight($sStdRead, 1)
				_Log($sStdRead)
			EndIf
		EndIf

		Sleep(5)
	WEnd

	Return $sData
EndFunc   ;==>_ProcessWaitClose
