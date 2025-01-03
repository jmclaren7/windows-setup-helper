; This script creates a basic 'taskbar' to make it easier to identify what is running and to make window switching easier in WinPE

#include <APIDlgConstants.au3>
#include <Array.au3>
#include <ButtonConstants.au3>
#include <Constants.au3>
#include <GuiEdit.au3>
#include <GUIConstantsEx.au3>
#include <GuiButton.au3>
#include <GuiImageList.au3>
#include <WinAPIFiles.au3>
#include <WinAPISys.au3>
#include <WinAPIsysinfoConstants.au3>
#include <WinAPIProc.au3>
#include <WindowsConstants.au3>

#include <CommonFunctions.au3>

Global $Title = "PE Window Bar"
Global $LogFullPath = StringTrimRight(@ScriptFullPath, 3) & "log"
;Global $LogTitle =  " Log"
Global $LogFlushAlways = True
Global $LogLevel = 3
_Log("Start")

OnAutoItExitRegister("_Exit")

; == General Globals =================================
Global $IsPE = StringInStr(@SystemDir, "X:")

; == Button and taskbar sizing/spacing values ========
Global $TaskBarHeight = 30
Global $TaskBarWidth = @DesktopWidth
Global $TaskBarTop = @DesktopHeight - $TaskBarHeight
Global $SpecialButtonWidth = 0
Global $ButtonHSpace = 3
Global $ButtonLeft = $ButtonHSpace + $SpecialButtonWidth + $ButtonHSpace
Global $ButtonWidth = 155
Global $ButtonVSpace = 2
Global $ButtonHeight = $TaskBarHeight - ($ButtonVSpace * 2)
Global $aWindows

; == Indexs for some window properties
Const $WindowTitleIndex = 0
Const $WindowHandleIndex = 1
Const $WindowStateIndex = 2
Const $WindowPIDIndex = 4
Const $WindowNameIndex = 5
Const $WindowPathIndex = 6
Const $WindowStyleIndex = 9
Const $WindowExStyleIndex = 10
Const $WindowButtonIDIndex = 15
Const $WindowContextCloseIDIndex = 16
Const $WindowContextMinimizeIDIndex = 17
Const $WindowContextRestoreIDIndex = 18
Const $WindowContextKillIDIndex = 19

If $IsPE Then
	; Set the area windows are allowed to use when maximized
	_WorkingArea(0, 0, @DesktopWidth, @DesktopHeight - $TaskBarHeight - 1)
Else
	; For testing outside of PE - move the taskbar up a little so it doesnt interfer with the normal taskbar
	$TaskBarTop -= 45
EndIf

; == Dummy GUI for dealing with no windows being active
$DummyForm = GUICreate("Dummy", 10, 10, -100, 100, $WS_POPUP, $WS_EX_TOOLWINDOW)
GUISetState(@SW_SHOWNOACTIVATE)

; == Setup GUI
Global $TaskBarForm = GUICreate($Title, $TaskBarWidth, $TaskBarHeight, 0, $TaskBarTop, $WS_POPUP, BitOR($WS_EX_TOOLWINDOW, $WS_EX_TOPMOST, $WS_EX_NOACTIVATE))
GUISetBkColor(0xE8E8E8)
;$SpecialButton = GUICtrlCreateButton("Close", $ButtonHSpace, $ButtonVSpace, $SpecialButtonWidth, $ButtonHeight, BitOR($BS_CENTER, $BS_VCENTER))
;_GUICtrlSetImage(-1, "C:\Windows\System32\shell32.dll", -16751, 16)
$idThisMenu = GUICtrlCreateContextMenu()
$CloseContext = GUICtrlCreateMenuItem("Close", $idThisMenu)
GUISetState(@SW_SHOWNOACTIVATE)


While 1
	;$TimerMainLoop = TimerInit()

	$nMsg = GUIGetMsg()
	If $nMsg > 0 Then _Log("GUI Event: " & $nMsg)
	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $CloseContext
			_Log("Case Exit")
			Exit

		Case 1 To 9999
			; A taskbar button was pressed, activate/deactivate the related window
			$SearchIndex = _ArraySearch($aWindows, $nMsg, 0, 0, 0, 0, 1, $WindowButtonIDIndex)
			If Not @error Then
				If WinActive($aWindows[$SearchIndex][$WindowHandleIndex]) Then
					_Log("Minimizing: " & $aWindows[$SearchIndex][$WindowTitleIndex] & " - " & $aWindows[$SearchIndex][$WindowHandleIndex], 2)
					WinSetState($aWindows[$SearchIndex][$WindowHandleIndex], "", @SW_MINIMIZE)
				Else
					_Log("Activating: " & $aWindows[$SearchIndex][$WindowTitleIndex] & " - " & $aWindows[$SearchIndex][$WindowHandleIndex], 2)
					WinActivate($aWindows[$SearchIndex][$WindowHandleIndex])
				EndIf
				ContinueLoop
			EndIf
			; A taskbar button's conext menu exit item was pressed
			$SearchIndex = _ArraySearch($aWindows, $nMsg, 0, 0, 0, 0, 1, $WindowContextCloseIDIndex)
			If Not @error Then
				WinClose($aWindows[$SearchIndex][$WindowHandleIndex])
				ContinueLoop
			EndIf
			; A taskbar button's conext menu minimize item was pressed
			$SearchIndex = _ArraySearch($aWindows, $nMsg, 0, 0, 0, 0, 1, $WindowContextMinimizeIDIndex)
			If Not @error Then
				WinSetState($aWindows[$SearchIndex][$WindowHandleIndex], "", @SW_MINIMIZE)
				ContinueLoop
			EndIf
			; A taskbar button's conext menu restore item was pressed
			$SearchIndex = _ArraySearch($aWindows, $nMsg, 0, 0, 0, 0, 1, $WindowContextRestoreIDIndex)
			If Not @error Then
				WinActivate($aWindows[$SearchIndex][$WindowHandleIndex])
				ContinueLoop
			EndIf
			; A taskbar button's conext menu kill item was pressed
			$SearchIndex = _ArraySearch($aWindows, $nMsg, 0, 0, 0, 0, 1, $WindowContextKillIDIndex)
			If Not @error Then
				ProcessClose($aWindows[$SearchIndex][$WindowPIDIndex])
				ContinueLoop
			EndIf


	EndSwitch

	_UpdateTaskBar()

	; Detect minimized window and move out of view
	For $i = 1 To $aWindows[0][0]
		$WindowState = WinGetState($aWindows[$i][$WindowHandleIndex])
		If BitAND($WindowState, $WIN_STATE_MINIMIZED) Then
			$WinPos = WinGetPos($aWindows[$i][$WindowHandleIndex])
			If $WinPos[1] < $TaskBarTop And $IsPE Then
				_Log("Moving minimized window out of view: " & $aWindows[$i][$WindowTitleIndex])
				WinMove($aWindows[$i][$WindowHandleIndex], "", ($i - 1) * 160, $TaskBarTop)
			EndIf
		EndIf
	Next

	; Detect change in active windows that wasnt cause by taskbar action
	$hActiveWindow = WinGetHandle("[Active]")
	$ActiveWindowIndex = _ArraySearch($aWindows, $hActiveWindow, 0, 0, 0, 0, 1, 1)
	If Not @error Then
		; The active window has a button so check it if it isn't already
		If GUICtrlRead ($aWindows[$ActiveWindowIndex][$WindowButtonIDIndex]) <> $GUI_CHECKED Then
			GUICtrlSetState($aWindows[$ActiveWindowIndex][$WindowButtonIDIndex], $GUI_CHECKED)
		EndIf

	Else
		; Active window doesn't have a button so uncheck all buttons
		For $i = 1 To $aWindows[0][0]
			GUICtrlSetState($aWindows[$i][$WindowButtonIDIndex], $GUI_UNCHECKED)
		Next
	EndIf

	;_Log("TimerMainLoop: " & TimerDiff($TimerMainLoop))
	Sleep(10)
WEnd


; =========================================================================
; =========================================================================
Func _Exit()
	_Log("_Exit" & @CRLF)
EndFunc   ;==>_Exit

Func _UpdateTaskBar()
	; Get window information and reverse order
	Local $aWindowsNew = _GetVisibleWindows(False)
	Local $iRows = UBound($aWindowsNew, $UBOUND_ROWS)
	Local $iCols = UBound($aWindowsNew, $UBOUND_COLUMNS)
	Local $aReversed[$iRows][$iCols]
	For $i = 1 To $iRows - 1
		For $j = 0 To $iCols - 1
			$aReversed[$i][$j] = $aWindowsNew[$iRows - $i][$j]
		Next
	Next
	For $j = 0 To $iCols - 1
		$aReversed[0][$j] = $aWindowsNew[0][$j]
	Next
	$aWindowsNew = $aReversed


	; Add columns
	_ArrayColInsert($aWindowsNew, $WindowButtonIDIndex)
	$aWindowsNew[0][$WindowButtonIDIndex] = "ButtonID"
	_ArrayColInsert($aWindowsNew, $WindowContextCloseIDIndex)
	$aWindowsNew[0][$WindowContextCloseIDIndex] = "ContextCloseID"
	_ArrayColInsert($aWindowsNew, $WindowContextMinimizeIDIndex)
	$aWindowsNew[0][$WindowContextMinimizeIDIndex] = "ContextMinimizeID"
	_ArrayColInsert($aWindowsNew, $WindowContextRestoreIDIndex)
	$aWindowsNew[0][$WindowContextRestoreIDIndex] = "ContextMaximizeID"
	_ArrayColInsert($aWindowsNew, $WindowContextKillIDIndex)
	$aWindowsNew[0][$WindowContextKillIDIndex] = "ContextKillID"


	; Delete windows we don't want in taskbar
	Local $aWindowsTrimed[0][UBound($aWindowsNew,2)]
	_ArrayAdd($aWindowsTrimed, _ArrayExtract($aWindowsNew, 0, 0))
	For $i = 1 To $aWindowsNew[0][0]
		;If $aWindowsNew[$i][1] = $TaskBarForm Then ContinueLoop
		If $aWindowsNew[$i][$WindowPathIndex] = "X:\sources\setup.exe" And $aWindowsNew[$i][$WindowTitleIndex] = "Install Windows" Then ContinueLoop
		If BitAND($aWindowsNew[$i][$WindowExStyleIndex], $WS_EX_TOOLWINDOW) Then ContinueLoop
		If Not $IsPE And BitAND($aWindowsNew[$i][$WindowStyleIndex], $WS_POPUP) Then ContinueLoop
		_ArrayAdd($aWindowsTrimed, _ArrayExtract($aWindowsNew, $i, $i))
	Next
	$aWindowsTrimed[0][0] = Ubound($aWindowsTrimed) - 1
	$aWindowsNew = $aWindowsTrimed


	; Setup $aWindows on first pass to match columns
	If Not IsArray($aWindows) Then Global $aWindows[$WindowHandleIndex][UBound($aWindowsNew, 2)]

	;_ArrayDisplay($aWindows,"$aWindows")
	;_ArrayDisplay($aWindowsNew,"$aWindowsNew")

	; If change is detected do things
	If $aWindowsNew[0][0] <> $aWindows[0][0] Then
		_Log("Change Detected")
		Local $aWindowsUpdated = $aWindows

		; Check for a change in windows
		If $aWindows[0][0] <> $aWindowsNew[0][0] Then
			; Delete all buttons
			For $i = 1 To $aWindows[0][0]
				GUICtrlDelete($aWindows[$i][$WindowButtonIDIndex])
			Next


			; Delete windows that no longer exist
			Local $aWindowsTrimed[0][UBound($aWindows,2)]
			_ArrayAdd($aWindowsTrimed, _ArrayExtract($aWindowsNew, 0, 0))
			For $i = 1 To $aWindows[0][0]
				$Index = _ArraySearch($aWindowsNew, $aWindows[$i][$WindowHandleIndex], 1, 0, 0, 2, 1, 1) ; Search for window handle
				If Not @error Then _ArrayAdd($aWindowsTrimed, _ArrayExtract($aWindowsNew, $Index, $Index))
			Next
			$aWindowsTrimed[0][0] = Ubound($aWindowsTrimed) - 1
			$aWindows = $aWindowsTrimed


			; Add new windows to the end of $aWindows
			For $i = 1 To $aWindowsNew[0][0]
				$SearchIndex = _ArraySearch($aWindows, $aWindowsNew[$i][$WindowHandleIndex], 1, 0, 1, 0, 1, 1)
				If @error Then
					;_Log("Add Window: " & $i)
					_ArrayAdd($aWindows, _ArrayExtract($aWindowsNew, $i, $i))
				EndIf
			Next
			$aWindows[0][0] = UBound($aWindows) - 1




			; Create buttons
			For $i = 1 To $aWindows[0][0]
				$aWindows[$i][$WindowButtonIDIndex] = GUICtrlCreateRadio(" ", $ButtonLeft + (($i - 1) * ($ButtonWidth + $ButtonHSpace)), $ButtonVSpace, $ButtonWidth, $ButtonHeight, BitOR($GUI_SS_DEFAULT_RADIO,$BS_LEFT,$BS_PUSHLIKE))
				GUICtrlSetTip(-1, $aWindows[$i][0])
				$idThisMenu = GUICtrlCreateContextMenu($aWindows[$i][$WindowButtonIDIndex])
				$aWindows[$i][$WindowContextKillIDIndex] = GUICtrlCreateMenuItem("Kill", $idThisMenu)
				$aWindows[$i][$WindowContextRestoreIDIndex] = GUICtrlCreateMenuItem("Restore", $idThisMenu)
				$aWindows[$i][$WindowContextMinimizeIDIndex] = GUICtrlCreateMenuItem("Minimize", $idThisMenu)
				$aWindows[$i][$WindowContextCloseIDIndex] = GUICtrlCreateMenuItem("Close", $idThisMenu)

				; Checks for special cases where we set a different icon
				$HelperTitle = "Windows Setup Helper v"
				If StringLeft($aWindows[$i][$WindowTitleIndex], StringLen($HelperTitle)) = $HelperTitle Then
					_GUICtrlSetImage($aWindows[$i][$WindowButtonIDIndex], "X:\sources\setup.exe", 0, 16)
				Else
					_GUICtrlSetImage($aWindows[$i][$WindowButtonIDIndex], $aWindows[$i][$WindowPathIndex], 0, 16)
				EndIf
			Next
		EndIf

	EndIf

	; Update button labels
	For $i = 1 To $aWindows[0][0]
		$ButtonText = GUICtrlRead($aWindows[$i][$WindowButtonIDIndex], 1)
		$SearchIndex = _ArraySearch($aWindowsNew, $aWindows[$i][$WindowHandleIndex], 1, 0, 1, 0, 1, 1)
		If Not @error Then
			$Title = $aWindowsNew[$SearchIndex][$WindowTitleIndex]
			$Label = _GenerateLabelText($Title)
			If $ButtonText <> $Label Then GUICtrlSetData($aWindows[$i][$WindowButtonIDIndex], $Label)
		EndIf
	Next

	Return 1
EndFunc

Func _GenerateLabelText($Text)
	;If StringLen($Text) > 32 Then
	;	$Text = StringLeft($Text, 12) & "..." & StringRight($Text, 10)
	;Else
	If StringLen($Text) > 22 Then
		$Text = StringLeft($Text, 22)
	EndIf

	Return $Text
EndFunc


Func _GUICtrlSetImage($hButton, $sFileIco, $iIndIco = 0, $iSize = 16)
	Switch $iSize
		Case 16, 24, 32
			$iSize = $iSize
		Case Else
			$iSize = $iSize
	EndSwitch
	Local $hImage = _GUIImageList_Create($iSize, $iSize, 5, 3, 6)
	If @error Then Return SetError(1, @error, $hImage)
	_GUIImageList_AddIcon($hImage, $sFileIco, $iIndIco)
	If @error Then Return SetError(2, @error, $hImage)
	_GUICtrlButton_SetImageList($hButton, $hImage, 0)
	If @error Then Return SetError(3, @error, $hImage)
	Return $hImage
EndFunc   ;==>_GUICtrlSetImage


Func _WorkingArea($iLeft = Default, $iTop = Default, $iWidth = Default, $iHeight = Default)
    Local Static $tWorkArea = 0
    If IsDllStruct($tWorkArea) Then
        _WinAPI_SystemParametersInfo($SPI_SETWORKAREA, 0, DllStructGetPtr($tWorkArea), $SPIF_SENDCHANGE)
        $tWorkArea = 0
    Else
        $tWorkArea = DllStructCreate($tagRECT)
        _WinAPI_SystemParametersInfo($SPI_GETWORKAREA, 0, DllStructGetPtr($tWorkArea))

        Local $tCurrentArea = DllStructCreate($tagRECT)
        Local $aArray[4] = [$iLeft, $iTop, $iWidth, $iHeight]
        For $i = 0 To 3
            If $aArray[$i] = Default Or $aArray[$i] < 0 Then
                $aArray[$i] = DllStructGetData($tWorkArea, $i + 1)
            EndIf
            DllStructSetData($tCurrentArea, $i + 1, $aArray[$i])
            $aArray[$i] = DllStructGetData($tWorkArea, $i + 1)
        Next
        _WinAPI_SystemParametersInfo($SPI_SETWORKAREA, 0, DllStructGetPtr($tCurrentArea), $SPIF_SENDCHANGE)
        $aArray[2] -= $aArray[0]
        $aArray[3] -= $aArray[1]
        Local $aReturn[4] = [$aArray[2], $aArray[3], $aArray[0], $aArray[1]]
        Return $aReturn
    EndIf
EndFunc   ;==>_WorkingArea
