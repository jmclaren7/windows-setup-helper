#Include <Array.au3>
#include <WinAPIProc.au3>
#include <WinAPISysWin.au3>
#include <WindowsConstants.au3>


_ArrayDisplay(_GetVisibleWindows())

; #FUNCTION# ====================================================================================================================
; Name ..........: _GetVisibleWindows
; Description ...: Get an array of information for visible windows
; Syntax ........: _GetVisibleWindows([$Options = 1])
; Parameters ....: $GetText             - [optional] True/False - Get window text (slow)
; Return values .: 2D Array of windows and window information
;				   [0][0] contains the number of windows
; Author ........: JohnMC - JohnsCS.com, based on AdamUL's _GetVisibleWindows
; Modified ......: 03/14/2024  --  v2.0
; ===============================================================================================================================
Func _GetVisibleWindows($GetText = False)
	Local $NewCol, $LastCol, $TotalTime, $Timer

	Local $TotalTime = TimerInit()

	; Retrieve a list of windows
	Local $aWinList = WinList()
	If Not IsArray($aWinList) Then Return SetError(0, 0, 0)
	$aWinList[0][1] = "WinHandle"


	; Add window states
	$NewCol = UBound($aWinList, 2)
	_ArrayColInsert($aWinList, $NewCol)
	$aWinList[0][$NewCol] = "State"
	For $i = 1 To $aWinList[0][0]
		$aWinList[$i][$NewCol] = WinGetState($aWinList[$i][1])
	Next


	; Add state descriptions
	$LastCol = $NewCol
	$NewCol = UBound($aWinList, 2)
	_ArrayColInsert($aWinList, $NewCol)
	$aWinList[0][$NewCol] = "StateDesc"
	For $i = 1 To $aWinList[0][0]
		Local $Desc = ""
		If BitAND($aWinList[$i][$LastCol], $WIN_STATE_EXISTS) Then $Desc &= "Exists"
		If BitAND($aWinList[$i][$LastCol], $WIN_STATE_VISIBLE) Then $Desc &= "Visible"
		If BitAND($aWinList[$i][$LastCol], $WIN_STATE_ENABLED) Then $Desc &= "Enabled"
		If BitAND($aWinList[$i][$LastCol], $WIN_STATE_ACTIVE) Then $Desc &= "Active"
		If BitAND($aWinList[$i][$LastCol], $WIN_STATE_MINIMIZED) Then $Desc &= "Minimized"
		If BitAND($aWinList[$i][$LastCol], $WIN_STATE_MAXIMIZED) Then $Desc &= "Maximized"

		$aWinList[$i][$NewCol] = $Desc
	Next


	; Delete undesirable windows
	Local $aNewWinList[0][UBound($aWinList, 2)]
	Local $NewIndex = 0
	For $i = 0 To $aWinList[0][0]
		If $i <> 0 And ($aWinList[$i][0] = "" Or Not BitAND($aWinList[$i][2], $WIN_STATE_VISIBLE)) Then ContinueLoop
		ReDim $aNewWinList[$NewIndex + 1][UBound($aWinList, 2)]
		For $b = 0 To UBound($aWinList, 2) - 1
			$aNewWinList[$NewIndex][$b] = $aWinList[$i][$b]
		Next
		$NewIndex += 1
	Next
	$aNewWinList[0][0] = UBound($aNewWinList) - 1
	$aWinList = $aNewWinList


	;Get Process ID (PID) and add to the array.
	$NewCol = UBound($aWinList, 2)
	_ArrayColInsert($aWinList, $NewCol)
	$aWinList[0][$NewCol] = "PID"
	For $i = 1 To $aWinList[0][0]
		$aWinList[$i][$NewCol] = WinGetProcess($aWinList[$i][1])
	Next


	; Add process name
	$NewCol = UBound($aWinList, 2)
	_ArrayColInsert($aWinList, $NewCol)
	$aWinList[0][$NewCol] = "Name"
	Local $aProcessList = ProcessList()
	For $i = 1 To $aWinList[0][0]
		For $b = 1 To UBound($aProcessList) - 1
			If $aProcessList[$b][1] = $aWinList[$i][4] Then
				$aWinList[$i][$NewCol] = $aProcessList[$b][0]
			EndIf
		Next
	Next


	; Add path using winapi method
	$NewCol = UBound($aWinList, 2)
	_ArrayColInsert($aWinList, $NewCol)
	$aWinList[0][$NewCol] = "Path"
	For $i = 1 To $aWinList[0][0]
		Local $Path = _WinAPI_GetProcessFileName($aWinList[$i][4])
		If $Path = "" Then
			; _WinAPI_EnumProcessModules only helps if the stars align, exe is probably in the system folder anyway
			Local $aEnum = _WinAPI_EnumProcessModules($aWinList[$i][4])
			If Not @error Then
				$TestPath = $aEnum[1]
			Else
				$TestPath = @SystemDir & "\" & $aWinList[$i][5]
			EndIf

			$TestFileAttrib = FileGetAttrib($TestPath)
			If Not @error And Not StringInStr($TestFileAttrib, "D") Then $Path = $TestPath
		EndIf
		$aWinList[$i][$NewCol] = $Path

	Next


	; Add command line string
	$NewCol = UBound($aWinList, 2)
	_ArrayColInsert($aWinList, $NewCol)
	$aWinList[0][$NewCol] = "Command"
	For $i = 1 To $aWinList[0][0]
		$aWinList[$i][$NewCol] = _WinAPI_GetProcessCommandLine($aWinList[$i][4])
	Next


	; Add window position and size
	;   -3200,-3200 is minimized window
	;   -8,-8 is maximized window on 1st display, and x,-8 is maximized window on the nth display were x is the nth display width plus -8 (W + -8).
	$NewCol = UBound($aWinList, 2)
	_ArrayColInsert($aWinList, $NewCol) ; Position (X,Y,W,H)
	$aWinList[0][$NewCol] = "Position"
	For $i = 1 To $aWinList[0][0]
		Local $aWinPosSize = WinGetPos($aWinList[$i][1])
		If Not @error Then
			$aWinList[$i][$NewCol] = $aWinPosSize[0] & "," & $aWinPosSize[1] & "," & $aWinPosSize[2] & "," & $aWinPosSize[3]
		EndIf
	Next


	; Add window style
	$NewCol = UBound($aWinList, 2)
	_ArrayColInsert($aWinList, $NewCol)
	$aWinList[0][$NewCol] = "Style"
	For $i = 1 To $aWinList[0][0]
		Local $tWINDOWINFO = _WinAPI_GetWindowInfo($aWinList[$i][1])
		$aWinList[$i][$NewCol] = DllStructGetData($tWINDOWINFO, 'Style', 1)
	Next


	; Add style descriptions
	$LastCol = $NewCol
	$NewCol = UBound($aWinList, 2)
	_ArrayColInsert($aWinList, $NewCol)
	$aWinList[0][$NewCol] = "StyleDesc"
	For $i = 1 To $aWinList[0][0]
		Local $Desc = ""
		If BitAND($aWinList[$i][$LastCol], $WS_BORDER) Then $Desc &= "Border"
		If BitAND($aWinList[$i][$LastCol], $WS_POPUP) Then $Desc &= "Popup"
		If BitAND($aWinList[$i][$LastCol], $WS_SYSMENU) Then $Desc &= "Sysmenu"
		If BitAND($aWinList[$i][$LastCol], $WS_GROUP) Then $Desc &= "Group"
		If BitAND($aWinList[$i][$LastCol], $WS_SIZEBOX) Then $Desc &= "Sizebox"
		If BitAND($aWinList[$i][$LastCol], $WS_CHILD) Then $Desc &= "Child"
		If BitAND($aWinList[$i][$LastCol], $WS_DLGFRAME) Then $Desc &= "Dialog"
		If BitAND($aWinList[$i][$LastCol], $WS_MINIMIZEBOX) Then $Desc &= "Minbox"
		If BitAND($aWinList[$i][$LastCol], $WS_MAXIMIZEBOX) Then $Desc &= "Maxbox"

		$aWinList[$i][$NewCol] = $Desc
	Next


	; Add window ExStyle
	$NewCol = UBound($aWinList, 2)
	_ArrayColInsert($aWinList, $NewCol)
	$aWinList[0][$NewCol] = "ExStyle"
	For $i = 1 To $aWinList[0][0]
		Local $tWINDOWINFO = _WinAPI_GetWindowInfo($aWinList[$i][1])
		$aWinList[$i][$NewCol] = DllStructGetData($tWINDOWINFO, 'ExStyle', 1)
	Next


	; Add ExStyle descriptions
	$LastCol = $NewCol
	$NewCol = UBound($aWinList, 2)
	_ArrayColInsert($aWinList, $NewCol)
	$aWinList[0][$NewCol] = "ExStyleDesc"
	For $i = 1 To $aWinList[0][0]
		Local $Desc = ""
		If BitAND($aWinList[$i][$LastCol], $WS_EX_TOOLWINDOW) Then $Desc &= "Tool"
		If BitAND($aWinList[$i][$LastCol], $WS_EX_TOPMOST) Then $Desc &= "Top"
		If BitAND($aWinList[$i][$LastCol], $WS_EX_CONTROLPARENT) Then $Desc &= "Parent"
		If BitAND($aWinList[$i][$LastCol], $WS_EX_APPWINDOW) Then $Desc &= "App"
		If BitAND($aWinList[$i][$LastCol], $WS_EX_MDICHILD) Then $Desc &= "Child"
		If BitAND($aWinList[$i][$LastCol], $WS_EX_DLGMODALFRAME) Then $Desc &= "Dialog"

		$aWinList[$i][$NewCol] = $Desc
	Next


	; Add TESTING
	$NewCol = UBound($aWinList, 2)
	_ArrayColInsert($aWinList, $NewCol)
	$aWinList[0][$NewCol] = "Testing"
	For $i = 1 To $aWinList[0][0]
		;$aWinList[$i][$NewCol] = _WinAPI_GetParent($aWinList[$i][1])
		;$aWinList[$i][$NewCol] = _WinAPI_GetParentProcess($aWinList[$i][4])
		;$aWinList[$i][$NewCol] = _WinAPI_GetProcessUser($aWinList[$i][4])[0]
	Next


	; Get Window's text and add to the array.
	$NewCol = UBound($aWinList, 2)
	_ArrayColInsert($aWinList, $NewCol)
	$aWinList[0][$NewCol] = "Text"
	If $GetText Then
		Local $WinDetectHiddenText = Opt("WinDetectHiddenText")
		Opt("WinDetectHiddenText", 1)
		For $i = 1 To $aWinList[0][0]
			$aWinList[$i][$NewCol] = WinGetText($aWinList[$i][1])
		Next
		Opt("WinDetectHiddenText", $WinDetectHiddenText)
	EndIf


	Return SetError(0, TimerDiff($TotalTime), $aWinList)
EndFunc   ;==>_GetVisibleWindows