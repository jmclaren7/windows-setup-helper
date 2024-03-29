#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Change2CUI=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

;#RequireAdmin

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
Global $MenuButtonWidth = 70
Global $ButtonHSpace = 3
Global $ButtonLeft = $ButtonHSpace + $MenuButtonWidth + $ButtonHSpace
Global $ButtonWidth = 150
Global $ButtonVSpace = 2
Global $ButtonHeight = $TaskBarHeight - ($ButtonVSpace * 2)
Global $aWindows

; == Indexs for some window properties
Const $WindowTitleIndex = 0
Const $WindowHandleIndex = 1
Const $WindowStateIndex = 2
Const $WindowPIDIndex = 3
Const $WindowNameIndex = 4
Const $WindowPathIndex = 5
Const $WindowStyleIndex = 6
Const $WindowExStyleIndex = 7
Const $WindowButtonIDIndex = 8


If $IsPE Then
	_WorkingArea(0, 0, @DesktopWidth, @DesktopHeight - $TaskBarHeight - 1)
Else
	$TaskBarTop -= 45
EndIf

; == Dummy GUI for dealing with no windows being active
$DummyForm = GUICreate("Dummy", 10, 10, -100, 100, $WS_POPUP, $WS_EX_TOOLWINDOW)
GUISetState(@SW_SHOWNOACTIVATE)

; == Setup GUI
Global $TaskBarForm = GUICreate($Title, $TaskBarWidth, $TaskBarHeight, 0, $TaskBarTop, $WS_POPUP, BitOR($WS_EX_TOOLWINDOW, $WS_EX_TOPMOST, $WS_EX_NOACTIVATE))
GUISetBkColor(0xE8E8E8)
$SpecialButton = GUICtrlCreateButton("Close", $ButtonHSpace, $ButtonVSpace, $MenuButtonWidth, $ButtonHeight, BitOR($BS_CENTER, $BS_VCENTER))
_GUICtrlSetImage(-1, "C:\Windows\System32\shell32.dll", -16751, 16)
GUISetState(@SW_SHOWNOACTIVATE)



_UpdateTaskBar()
AdlibRegister("_UpdateTaskBar", 200)

While 1
	;$TimerMainLoop = TimerInit()

	$nMsg = GUIGetMsg()
	If $nMsg > 0 Then _Log("GUI Event: " & $nMsg)
	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $SpecialButton
			_Log("Case Exit")
			Exit

		Case 1 To 9999
			; A taskbar button was pressed, activate/deactivate the related window
			$SearchIndex = _ArraySearch($aWindows, $nMsg, 0, 0, 0, 0, 1, 15)
			If Not @error Then
				If WinGetHandle("[Active]") = $aWindows[$SearchIndex][$WindowHandleIndex] Then
					_Log("Minimizing: " & $aWindows[$SearchIndex][0] & " - " & $aWindows[$SearchIndex][$WindowHandleIndex], 2)
					WinSetState($aWindows[$SearchIndex][$WindowHandleIndex], "", @SW_MINIMIZE)
				Else
					_Log("Activating: " & $aWindows[$SearchIndex][0] & " - " & $aWindows[$SearchIndex][$WindowHandleIndex], 2)
					WinActivate($aWindows[$SearchIndex][$WindowHandleIndex])
				EndIf

			EndIf


	EndSwitch


	; Detect minimized window and move out of view
	For $i = 1 To $aWindows[0][0]
		_Log("WinGetState: " & $i & " vs " & Ubound($aWindows))
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
		; Active active window doesn't have a button so uncheck all buttons
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
	Local $aWindowsNew = __GetVisibleWindows(False)
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


	; Add column for button ID
	_ArrayColInsert($aWindowsNew, $WindowButtonIDIndex)
	$aWindowsNew[0][$WindowButtonIDIndex] = "ButtonID"


	; Delete windows we don't want in taskbar
	Local $aWindowsTrimed[0][UBound($aWindowsNew,2)]
	_ArrayAdd($aWindowsTrimed, _ArrayExtract($aWindowsNew, 0, 0))
	For $i = 1 To $aWindowsNew[0][0]
		;If $aWindowsNew[$i][1] = $TaskBarForm Then ContinueLoop
		If $aWindowsNew[$i][6] = "X:\sources\setup.exe" Then ContinueLoop
		If BitAND($aWindowsNew[$i][$WindowExStyleIndex], $WS_EX_TOOLWINDOW) Then ContinueLoop
		If Not $IsPE And BitAND($aWindowsNew[$i][$WindowStyleIndex], $WS_POPUP) Then ContinueLoop
		_ArrayAdd($aWindowsTrimed, _ArrayExtract($aWindowsNew, $i, $i))
	Next
	$aWindowsTrimed[0][0] = Ubound($aWindowsTrimed) - 1
	$aWindowsNew = $aWindowsTrimed


	; Setup $aWindows on first pass to match columns
	If Not IsArray($aWindows) Then Global $aWindows[1][UBound($aWindowsNew, 2)]

	;_ArrayDisplay($aWindows,"$aWindows")
	;_ArrayDisplay($aWindowsNew,"$aWindowsNew")
	;Exit

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
				$Index = _ArraySearch($aWindowsNew, $aWindows[$i][1], 1, 0, 0, 2, 1, 1) ; Search for window handle
				If Not @error Then _ArrayAdd($aWindowsTrimed, _ArrayExtract($aWindowsNew, $Index, $Index))
			Next
			$aWindowsTrimed[0][0] = Ubound($aWindowsTrimed) - 1
			$aWindows = $aWindowsTrimed


			; Add new windows to the end of $aWindows
			For $i = 1 To $aWindowsNew[0][0]
				$SearchIndex = _ArraySearch($aWindows, $aWindowsNew[$i][1], 1, 0, 1, 0, 1, 1)
				If @error Then
					;_Log("Add Window: " & $i)
					_ArrayAdd($aWindows, _ArrayExtract($aWindowsNew, $i, $i))
				EndIf
			Next
			$aWindows[0][0] = UBound($aWindows) - 1


			; Create buttons
			For $i = 1 To $aWindows[0][0]
				;_Log("Adding button: " & $i)
				$Label = "";_GenerateLabelText($aWindows[$i][0])
				$aWindows[$i][$WindowButtonIDIndex] = GUICtrlCreateRadio(" " & $Label, $ButtonLeft + (($i - 1) * ($ButtonWidth + $ButtonHSpace)), $ButtonVSpace, $ButtonWidth, $ButtonHeight, BitOR($GUI_SS_DEFAULT_RADIO,$BS_LEFT,$BS_PUSHLIKE))
				GUICtrlSetTip(-1, $aWindows[$i][0])
				; Checks for special cases where we set a different icon
				If $aWindows[$i][$WindowPIDIndex] = @AutoItPID Then
					_GUICtrlSetImage(-1, "X:\sources\setup.exe", 0, 16)
				Else
					_GUICtrlSetImage(-1, $aWindows[$i][$WindowPathIndex], 0, 16)
				EndIf
			Next
		EndIf

	EndIf

	; Update button labels
	For $i = 1 To $aWindows[0][0]
		$ButtonText = GUICtrlRead($aWindows[$i][$WindowButtonIDIndex], 1)
		$SearchIndex = _ArraySearch($aWindowsNew, $aWindows[$i][$WindowHandleIndex], 1, 0, 1, 0, 1, 1)
		If Not @error Then
			$Title = $aWindowsNew[$SearchIndex][0]
			$Label = _GenerateLabelText($Title)
			If $ButtonText <> $Label Then GUICtrlSetData($aWindows[$i][$WindowButtonIDIndex], $Label)
		EndIf
	Next

	Return 1
EndFunc


Func _GenerateLabelText($Text)
	If StringLen($Text) > 32 Then
		$Text = StringLeft($Text, 12) & "..." & StringRight($Text, 10)
	ElseIf StringLen($Text) > 22 Then
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

; #FUNCTION# ====================================================================================================================
; Name ..........: _GetVisibleWindows
; Description ...: Get an array of information for visible windows
; Syntax ........: _GetVisibleWindows([$Options = 1])
; Parameters ....: $GetText             - [optional] True/False - Get window text (slow)
; Return values .: 2D Array of windows and window information
;				   [0][0] contains the number of windows
; Author ........: JohnMC - JohnsCS.com, based on AdamUL's _GetVisibleWindows
; Modified ......:
; ===============================================================================================================================
Func __GetVisibleWindows($GetText = False)
	Local $NewCol, $LastCol, $TotalTime, $Timer

	Local $TotalTime = TimerInit()

	; 0,1 Retrieve a list of windows
	Local $aWinList = WinList()
	If Not IsArray($aWinList) Then Return SetError(0, 0, 0)
	$aWinList[0][1] = "WinHandle"

	; 2 Add window states
	$NewCol = UBound($aWinList, 2)
	_ArrayColInsert($aWinList, $NewCol)
	$aWinList[0][$NewCol] = "State"
	For $i = 1 To $aWinList[0][0]
		$aWinList[$i][$NewCol] = WinGetState($aWinList[$i][1])
	Next

	; Delete undesirable windows
	Local $WindowStateIndex = 2
	Local $aNewWinList[0][UBound($aWinList, 2)]
	Local $NewIndex = 0
	For $i = 0 To $aWinList[0][0]
		If $i <> 0 And ($aWinList[$i][0] = "" Or Not BitAND($aWinList[$i][$WindowStateIndex], $WIN_STATE_VISIBLE)) Then ContinueLoop
		ReDim $aNewWinList[$NewIndex + 1][UBound($aWinList, 2)]
		For $b = 0 To UBound($aWinList, 2) - 1
			$aNewWinList[$NewIndex][$b] = $aWinList[$i][$b]
		Next
		$NewIndex += 1
	Next
	$aNewWinList[0][0] = UBound($aNewWinList) - 1
	$aWinList = $aNewWinList

	; 3 Get Process ID (PID) and add to the array.
	$NewCol = UBound($aWinList, 2)
	_ArrayColInsert($aWinList, $NewCol)
	$aWinList[0][$NewCol] = "PID"
	For $i = 1 To $aWinList[0][0]
		$aWinList[$i][$NewCol] = WinGetProcess($aWinList[$i][1])
	Next

	; 4 Add process name
	Local $PIDIndex = 3
	$NewCol = UBound($aWinList, 2)
	_ArrayColInsert($aWinList, $NewCol)
	$aWinList[0][$NewCol] = "Name"
	Local $aProcessList = ProcessList()
	For $i = 1 To $aWinList[0][0]
		For $b = 1 To UBound($aProcessList) - 1
			If $aProcessList[$b][1] = $aWinList[$i][$PIDIndex] Then
				$aWinList[$i][$NewCol] = $aProcessList[$b][0]
			EndIf
		Next
	Next

	; 5 Add path using winapi method
	$NewCol = UBound($aWinList, 2)
	_ArrayColInsert($aWinList, $NewCol)
	$aWinList[0][$NewCol] = "Path"
	Local $ProccessNameIndex = 3
	For $i = 1 To $aWinList[0][0]
		Local $Path = _WinAPI_GetProcessFileName($aWinList[$i][$ProccessNameIndex])
		; No path might mean the process is elevated so let's try some other things...
		If $Path = "" Then
			; Might only help if the stars align
			Local $aEnum = _WinAPI_EnumProcessModules($aWinList[$i][$ProccessNameIndex])
			If Not @error Then
				$TestPath = $aEnum[1]
				; The exe might be in the system folder (elevated cmd or task manager)
			Else
				$TestPath = @SystemDir & "\" & $aWinList[$i][5]
			EndIf

			$TestFileAttrib = FileGetAttrib($TestPath)
			If Not @error And Not StringInStr($TestFileAttrib, "D") Then $Path = $TestPath
		EndIf
		$aWinList[$i][$NewCol] = $Path

	Next

	; 6,7 Add window style and exstyle
	$NewCol = UBound($aWinList, 2)
	_ArrayColInsert($aWinList, $NewCol)
	$aWinList[0][$NewCol] = "Style"
	$NewCol2 = UBound($aWinList, 2)
	_ArrayColInsert($aWinList, $NewCol2)
	$aWinList[0][$NewCol2] = "ExStyle"
	For $i = 1 To $aWinList[0][0]
		Local $tWINDOWINFO = _WinAPI_GetWindowInfo($aWinList[$i][1])
		$aWinList[$i][$NewCol] = DllStructGetData($tWINDOWINFO, 'Style', 1)
		$aWinList[$i][$NewCol2] = DllStructGetData($tWINDOWINFO, 'ExStyle', 1)
	Next

	Return SetError(0, TimerDiff($TotalTime), $aWinList)
EndFunc   ;==>_GetVisibleWindows
;===============================================================================
; Function Name:   	_Log()
; Description:		Console & File Loging
; Call With:		_Log($Text, $iLevel)
; Parameter(s): 	$sMessage - Text to print
;					$iLevel - The level *this* message
;								Use 1 for critical or always shown (default), 2+ for debuging
;
; Return Value(s):  The original message, if $iLevel is greater than $LogLevel returns an empty string
; Notes:			Some options are configured with global variables
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		11/20/2023 --  V1.0 Added function to CommoneFunctions.au3
;===============================================================================
; Write to the log, prepend a timestamp, create a custom log GUI
Func _Log($sMessage, $iLevel = 1)
	; Global options
	If Not IsDeclared("LogLevel") Then Global $LogLevel = 1 ; Only show messages this level or below
	If Not IsDeclared("LogTitle") Then Global $LogTitle = "" ; Title to use for log GUI, no title will skip the GUI
	If Not IsDeclared("LogWindowStart") Then Global $LogWindowStart = -1 ; -1 for center
	If Not IsDeclared("LogWindowSize") Then Global $LogWindowSize = 750 ; Starting width, height will be .6 of this value
	If Not IsDeclared("LogFullPath") Then Global $LogFullPath = "" ; The path of the log file, empty value will not log to file
	If Not IsDeclared("LogFileMaxSize") Then Global $LogFileMaxSize = 1024 ; Size limit for log in KB
	If Not IsDeclared("LogFlushAlways") Then Global $LogFlushAlways = False

	Local $LogFileMaxSize_Bytes = $LogFileMaxSize * 1024
	Local $sTime = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & "> "
	Local $sLogLine = $sTime & $sMessage

	; Do not log this message if $iLevel is greater than global $LogLevel
	If $iLevel > $LogLevel Then Return ""

	; Send to console
	ConsoleWrite($sLogLine & @CRLF)

	; Append message to custom GUI if $LogTitle is set
	If $LogTitle <> "" Then
		If Not IsDeclared("_hLogEdit") Then
			; The GUI doesn't exist, create it
			Global $_hLogWindow = GUICreate($LogTitle, $LogWindowSize, Round($LogWindowSize * 0.6), $LogWindowStart, $LogWindowStart, BitOR($GUI_SS_DEFAULT_GUI, $WS_SIZEBOX))
			Global $_hLogEdit = GUICtrlCreateEdit("", 0, 0, $LogWindowSize, Round($LogWindowSize * 0.6), BitOR($ES_MULTILINE, $ES_WANTRETURN, $WS_VSCROLL, $WS_HSCROLL))
			GUICtrlSetFont(-1, 10, 400, 0, "Consolas")
			GUICtrlSetColor(-1, 0xFFFFFF)
			GUICtrlSetBkColor(-1, 0x000000)
			GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM)
			GUISetState(@SW_SHOW)
			_GUICtrlEdit_AppendText($_hLogEdit, $sLogLine)
		Else
			; Update an existing GUI
			_GUICtrlEdit_BeginUpdate($_hLogEdit)
			_GUICtrlEdit_AppendText($_hLogEdit, @CRLF & $sLogLine)
			_GUICtrlEdit_LineScroll($_hLogEdit, -StringLen($sLogLine), _GUICtrlEdit_GetLineCount($_hLogEdit))
			_GUICtrlEdit_EndUpdate($_hLogEdit)
		EndIf
	EndIf

	; Append message to file
	If $LogFullPath <> "" Then
		If Not IsDeclared("_hLogFile") Then Global $_hLogFile = FileOpen($LogFullPath, $FO_APPEND)

		; Limit log size
		If $LogFileMaxSize > 0 Then
			Local $iCurrentSize = FileGetPos($_hLogFile) ; + StringLen($sLogLine)

			If $iCurrentSize > $LogFileMaxSize_Bytes Then
				; Rewrite desired data to begining of file, drop trailing data, flush to disk.
				FileSetPos($_hLogFile, 0, $FILE_BEGIN)
				$sLogLine = FileRead($_hLogFile) & $sLogLine
				$sLogLine = StringRight($sLogLine, $LogFileMaxSize_Bytes - 1024)
				FileSetPos($_hLogFile, 0, $FILE_BEGIN)
				FileWrite($_hLogFile, $sLogLine & @CRLF)
				FileSetEnd($_hLogFile)
				FileFlush($_hLogFile)

			Else
				; Write to file
				FileWrite($_hLogFile, $sLogLine & @CRLF)
				If $LogFlushAlways Then FileFlush($_hLogFile)

			EndIf

		EndIf

	EndIf

	Return $sMessage
EndFunc   ;==>_Log
