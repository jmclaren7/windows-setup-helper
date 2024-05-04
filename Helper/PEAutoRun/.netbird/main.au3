#include <WinAPIProc.au3>

; NetBird creates a mesh/overlay network and luckily the NetBird client works well in WinPEx64
; We can use it to join a network on WinPE boot and then use that network to VNC to the WinPE session
; This script is an example to start the NetBird service
;     Add netbird.exe and wintun.dll to the folder with this script
;     Add the setup key below
; Be sure you understand the security implications and that you have your setup key and access control configured correctly

$Key = ""

Global $Title = "PENetBird"
Global $IsPE = StringInStr(@WindowsDir, "X:")
If Not $IsPE Then Exit
FileChangeDir(@ScriptDir)

_Log($Title)
_Log("@WorkingDir="&@WorkingDir)
_Log("@ScriptDir="&@ScriptDir)

OnAutoItExitRegister("_Exit")

; Wait for network
For $i=1 to 10
	Ping ("8.8.8.8", 1000)
	If Not @error Then ExitLoop
	Sleep(1000)
Next

$Command = "netbird.exe service install"
$Run = _RunWait($Command)
_Log("$Run=" & $Run & "  @error=" & @error & "  $Command=" & $Command)

$Command = "netbird.exe service start"
$Run = _RunWait($Command)
_Log("$Run=" & $Run & "  @error=" & @error & "  $Command=" & $Command)

$Command = "netbird.exe up -k " & $Key
$Run = _RunWait($Command)
_Log("$Run=" & $Run & "  @error=" & @error & "  $Command=" & $Command)

; Loop to update status
While 1
	$Command = "netbird.exe status"
	$Run = _RunWait($Command)
	If StringInStr($Run, "Signal: Connected") Then _UpdateStatusBar()

	Sleep(2000)
Wend

Exit

;=========== =========== =========== =========== =========== =========== =========== ===========
;=========== =========== =========== =========== =========== =========== =========== ===========

Func _Exit()

EndFunc

Func _UpdateStatusBar()
	Local $Path = @TempDir & "\Helper_Status_" & $Title & ".txt"

	$hFile = FileOpen ($Path, 2)
	FileWrite($hFile, "NetBird Connected")

	FileClose($hFile)
EndFunc

Func _Log($Data)
	ConsoleWrite($Title & ": " & $Data & @CRLF)
EndFunc


;===============================================================================
; Function Name:    _RunWait
; Description:		Improved version of RunWait that plays nice with my console logging
; Call With:		_RunWait($Run, $Working="")
; Parameter(s):
; Return Value(s):  On Success - Return value of Run() (Should be PID)
; 					On Failure - Return value of Run()
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		01/16/2016  --  v1.1
;===============================================================================
Func _RunWait($sProgram, $Working = "", $Show = @SW_HIDE, $Opt = $STDERR_MERGED, $Live = False)
	Local $sData, $iPid

	$iPid = Run($sProgram, $Working, $Show, $Opt)
	If @error Then
		_ConsoleWrite("_RunWait: Couldn't Run " & $sProgram)
		Return SetError(1, 0, 0)
	EndIf

	$sData = _ProcessWaitClose($iPid, $Live)

	Return SetError(0, $iPid, $sData)
EndFunc   ;==>_RunWait
;===============================================================================
; Function Name:    _ProcessWaitClose
; Description:		ProcessWaitClose that handles stdout from the running process
;					Proccess must have been started with $STDERR_CHILD + $STDOUT_CHILD
; Call With:		_ProcessWaitClose($iPid)
; Parameter(s):
; Return Value(s):  On Success -
; 					On Failure -
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		09/8/2023  --  v1.3
;===============================================================================
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
				_ConsoleWrite($sStdRead)
			EndIf
		EndIf

		Sleep(5)
	WEnd

	Return $sData
EndFunc   ;==>_ProcessWaitClose