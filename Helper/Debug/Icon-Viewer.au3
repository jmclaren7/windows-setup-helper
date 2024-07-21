#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Fileversion=0.98.7.0
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductName=Icon-Viewer
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=None
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
; Icon Viewer
; View and save icons from resource files like shell32.dll or explorer.exe
; By JohnMC - JohnsCS.com - https://github.com/jmclaren7/autoit-scripts

#include <ComboConstants.au3>
#include <StructureConstants.au3> ;
#include <GuiComboBox.au3>
#include <GuiListView.au3> ;
#include <GuiImageList.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiEdit.au3>


Global $Title = "Icon-Viewer"
Global $IconGUI, $IconListView
_Log(@SystemDir)

; ==== GUI setup
$IconGUI = GUICreate($Title, 801, 587, -1, -1, BitOR($GUI_SS_DEFAULT_GUI,$WS_MAXIMIZEBOX,$WS_SIZEBOX,$WS_THICKFRAME,$WS_TABSTOP))
$IconListView = GUICtrlCreateListView("", 1, 34, 798, 550, BitOR($GUI_SS_DEFAULT_LISTVIEW,$LVS_NOCOLUMNHEADER,$LVS_NOSORTHEADER,$LVS_NOLABELWRAP,$WS_VSCROLL))
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 0, 50)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKRIGHT+$GUI_DOCKTOP+$GUI_DOCKBOTTOM)
$IconListView_0 = GUICtrlCreateListViewItem("", $IconListView)
$IconFileCombo = GUICtrlCreateCombo("", 4, 6, 408, 25)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKTOP+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)
$BrowseButton = GUICtrlCreateButton("Browse", 420, 2, 60, 28)
GUICtrlSetResizing(-1, $GUI_DOCKLEFT+$GUI_DOCKTOP+$GUI_DOCKWIDTH+$GUI_DOCKHEIGHT)

$ContextMenu = GUICtrlCreateContextMenu($IconListView)
$SaveContext = GUICtrlCreateMenuItem("Save...", $ContextMenu)
GUISetIcon("C:\Windows\System32\Shell32.dll", -327, $IconGUI)

_GUICtrlListView_SetView ( $IconListView, 2)
GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")

$EnterDummy = GUICtrlCreateDummy()
Local $aAccelKeys[1][2]
$aAccelKeys[0][0] = "{ENTER}"
$aAccelKeys[0][1] = $EnterDummy
GUISetAccelerators($aAccelKeys)

GUISetState(@SW_SHOW)

Local $tInfo
_GUICtrlComboBox_GetComboBoxInfo($IconFileCombo, $tInfo)
Local $hCombo = DllStructGetData($tInfo, "hEdit")

GUICtrlSetData($IconFileCombo, "Shell32.dll|Imageres.dll|wmploc.dll|netshell.dll|DDORes.dll", "Shell32.dll")

_Log("GUI Ready: " & @ScriptName)

_LoadIcons(GUICtrlRead($IconFileCombo))

_Log("Done")

; ==== GUI loop
While 1
	$Msg = GUIGetMsg()
	If $Msg > 0 Then _Log("$Msg=" & $Msg)
	Switch $Msg
		Case $GUI_EVENT_CLOSE
			Exit

		Case $BrowseButton
			_Log("$BrowseButton")
			$NewIconFile = FileOpenDialog($Title, "", "Icon Files (*.exe;*.dll;*.ico)|All (*.*)", 1, "", $IconGUI)
			If Not @error Then
				_Log("$NewIconFile="&$NewIconFile)
				$SelectComboText = _GUICtrlComboBox_SelectString($IconFileCombo, $NewIconFile)
				If $SelectComboText = -1 Then
					$NewStringIndex = _GUICtrlComboBox_AddString($IconFileCombo, $NewIconFile)
					_GUICtrlComboBox_SetCurSel($IconFileCombo , $NewStringIndex)
				EndIf
				_LoadIcons($NewIconFile)
			EndIf

		Case $IconListView
			_Log("$IconListView")

		Case $IconFileCombo
			_Log("$IconFileCombo")
			_LoadIcons(GUICtrlRead($IconFileCombo))

		Case $EnterDummy
			_Log("$EnterDummy")
			$Focus = ControlGetHandle($IconGUI, "", ControlGetFocus ($IconGUI))

			If $Focus = $hCombo Then
				_Log("$IconFileCombo")
				_LoadIcons(GUICtrlRead($IconFileCombo))

			EndIf

		Case $SaveContext
			_Log("$SaveContext")
			_Log(_GUICtrlListView_GetSelectedIndices($IconListView))

	EndSwitch

	Sleep(10)
Wend


Func _LoadIcons($IconFile)
	_Log("_LoadIcons: " & $IconFile)
	_Timer()

	; Check if the file exists, if it doesn't check the system32 folder
	If Not FileExists($IconFile) Then
		$IconFile = @SystemDir & "\" & $IconFile
		If Not FileExists($IconFile) Then Return SetError(1)
	EndIf

	; Remove all existing list view items
	_GUICtrlListView_DeleteAllItems($IconListView)

	If Not FileExists($IconFile) Then Return SetError(1)

	; Create the image list, this is attached to the listview control later
	Local $hImageList = _GUIImageList_Create(32, 32, 5, 3, 256, 512)

	; Get icon count and names from file
	Local $aIconNames = _WinAPI_EnumResourceNames($IconFile, BitOr($RT_GROUP_ICON,$RT_GROUP_CURSOR)) ; BitOr($RT_GROUP_ICON,$RT_GROUP_CURSOR)
	If @error Then
		_Log("_WinAPI_EnumResourceNames Error: " & @error)
		Return SetError(1, @error, 0)
	Endif

	; Add each icon to the image list
	For $Index = 1 To $aIconNames[0]
		$ErrorCheck = _GUIImageList_AddIcon($hImageList, $IconFile, $Index - 1, True)
		If $ErrorCheck = -1 Then _Log("_GUIImageList_AddIcon Error: $Index=" & $Index & "  $aIconNames[$Index]=" & $aIconNames[$Index] & "  $ErrorCheck=" & $ErrorCheck)
	Next
	_Log("Timer: " & _Timer())

	; Attach image list to listview control
	_GUICtrlListView_SetImageList($IconListView, $hImageList, 1)

	; Populate listview with text and icon
	For $Index = 1 To $aIconNames[0]
		$Text = "" & $Index & @CRLF & " (-" & $aIconNames[$Index] & ")"
		$ErrorCheck = _GUICtrlListView_AddItem($IconListView, $Text, $Index - 1)
		If $ErrorCheck = -1 Then _Log("  _GUICtrlListView_AddItem Error: $Index=" & $Index & "  $aIconNames[$Index]=" & $aIconNames[$Index] & "  $ErrorCheck=" & $ErrorCheck)
	Next

	_Log("Timer: " & _Timer())

	Return $aIconNames[0]
EndFunc   ;==>_LoadIcons

Func _Exit()
	_Log("_Exit")
EndFunc



;~ Func _SaveToFile($sFile, $iIconId, $iWidth = 32, $iHeight = 32)
;~ 	Local $sOutputFile = $sIconPath & '\' & StringReplace($sFile, '.dll', '_') & StringReplace($iIconId & '.ico', ' ', '')
;~ 	ConsoleWrite($sFile & @CRLF)
;~ 	Local $hIcon = _WinAPI_ShellExtractIcon($sFile, $iIconId, $iWidth, $iHeight)
;~ 	_WinAPI_SaveHICONToFile($sOutputFile, $hIcon)
;~ 	;ShellExecute($sFile)
;~ 	_WinAPI_DestroyIcon($hIcon)
;~ EndFunc   ;==>_SaveToFile



; Used for detecting events related to the listview
Func WM_NOTIFY($hWnd, $iMsg, $wParam, $lParam)
	Local $tNMHDR = DllStructCreate($tagNMHDR, $lParam)
	Local $iCode = DllStructGetData($tNMHDR, "Code")

	Switch $iCode
		Case $NM_CLICK, $NM_DBLCLK, $NM_RCLICK, $NM_RETURN, $NM_SETFOCUS
			Local $tInfo = DllStructCreate($tagNMITEMACTIVATE, $lParam)
			Local $Index = DllStructGetData($tInfo, "Index") + 1 ; +1 to translate to 1 based icon index

			_Log("WM_NOTIFY: $hWnd=" & $hWnd & "  $iMsg=" & $iMsg & "  $wParam=" & $wParam & "  $lParam=" & $lParam & "  $Index=" & $Index)
			_Log("GetSelectedIndices: " & _GUICtrlListView_GetSelectedIndices($IconListView))

			Global $MenuMsg = 2000 + $Index
	EndSwitch

	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY


; #FUNCTION# ====================================================================================================================
; Name ..........: _Timer
; Description ...: Creates a quick to use timer
; Syntax ........: _Timer([$Reset = True])
; Parameters ....: $Reset               - [optional] If True, time will reset
; Return values .: The time in miliseconds since last reset or init
; Author ........: JohnMC - JohnsCS.com
; Modified ......: 03/14/2024  --  v1.0
; ===============================================================================================================================
Func _Timer($Reset = True)
	If Not IsDeclared("_Timer_Handle") Then
		Global $_Timer_Handle
		$_Timer_Handle = TimerInit()
		Return True
	EndIf
	Local $Return = Round(TimerDiff($_Timer_Handle), 1)
	If $Reset Then $_Timer_Handle = TimerInit()
	Return $Return
EndFunc   ;==>_Timer

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
; Date/Last Change:	4/26/2024 -- Fixed global handling, added minimize window on start
;					5/6/2024 -- Added $bOverWriteLast, changed the way line returns work on consolewrite
;===============================================================================
; Write to the log, prepend a timestamp, create a custom log GUI
Func _Log($sMessage, $iLevel = Default, $bOverWriteLast = Default)
	Static Local $_hLogFile

	; Defaults
	If $iLevel = Default Then $iLevel = 1
	If $bOverWriteLast = Default Then $bOverWriteLast = False

	; Global options
	Global $LogLevel, $LogTitle, $LogWindowStart, $LogWindowSize, $LogFullPath, $LogFileMaxSize, $LogFlushAlways

	; If $LogTitle is empty, skip the GUI
	If $LogLevel = "" Then $LogLevel = 1 ; Only show messages this level or below
	If $LogWindowStart = "" Then Global $LogWindowStart = -1 ; -1 for center, -# for minimized with position being the absolute value
	If $LogWindowSize = "" Then Global $LogWindowSize = 750 ; Starting width, height will be .6 of this value
	If $LogFullPath = "" Then Global $LogFullPath = "" ; The path of the log file, empty value will not log to file
	If $LogFileMaxSize = "" Then Global $LogFileMaxSize = 1024 ; Size limit for log in KB
	If $LogFlushAlways = "" Then Global $LogFlushAlways = False ; Flush log to disk after each update

	Local $LogFileMaxSize_Bytes = $LogFileMaxSize * 1024
	Local $sTime = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & "> "
	Local $sLogLine = $sTime & $sMessage
	Local $Minimize = False

	; Do not log this message if $iLevel is greater than global $LogLevel
	If $iLevel > $LogLevel Then Return ""

	; Send to console
	If $bOverWriteLast And Not @Compiled Then
		; Do Nothing
	ElseIf $bOverWriteLast Then
		ConsoleWrite(@CR & $sLogLine)
	Else
		ConsoleWrite(@CRLF & $sLogLine)
	EndIf

	; Append message to custom GUI if $LogTitle is set
	If $LogTitle <> "" Then
		If Not IsDeclared("_hLogEdit") Then
			; The GUI doesn't exist, create it
			If $LogWindowStart < -1 Then
				$LogWindowStart = Abs($LogWindowStart)
				$Minimize = True
			EndIf
			Global $_hLogWindow = GUICreate($LogTitle, $LogWindowSize, Round($LogWindowSize * 0.6), $LogWindowStart, $LogWindowStart, BitOR($GUI_SS_DEFAULT_GUI, $WS_SIZEBOX))
			Global $_hLogEdit = GUICtrlCreateEdit("", 0, 0, $LogWindowSize, Round($LogWindowSize * 0.6), BitOR($ES_MULTILINE, $ES_WANTRETURN, $WS_VSCROLL, $WS_HSCROLL))
			GUICtrlSetFont(-1, 10, 400, 0, "Consolas")
			GUICtrlSetColor(-1, 0xFFFFFF)
			GUICtrlSetBkColor(-1, 0x000000)
			GUICtrlSetResizing(-1, $GUI_DOCKLEFT + $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKBOTTOM)
			_GUICtrlEdit_SetReadOnly($_hLogEdit, True)
			GUISetState(@SW_SHOW, $_hLogWindow)
			If $Minimize Then GUISetState(@SW_MINIMIZE, $_hLogWindow)
			_GUICtrlEdit_AppendText($_hLogEdit, $sLogLine)
		Else
			; Update an existing GUI
			_GUICtrlEdit_BeginUpdate($_hLogEdit)

			If $bOverWriteLast Then
				Local $sFullText = _GUICtrlEdit_GetText($_hLogEdit)
				;Msgbox(0,"",$sFullText)
				$sFullText = StringLeft($sFullText, StringInStr($sFullText, @CRLF, 0, -1) - 1)
				;Msgbox(0,"",$sFullText)
				_GUICtrlEdit_SetText($_hLogEdit, $sFullText)

			EndIf
			_GUICtrlEdit_AppendText($_hLogEdit, @CRLF & $sLogLine)
			_GUICtrlEdit_LineScroll($_hLogEdit, -StringLen($sLogLine), _GUICtrlEdit_GetLineCount($_hLogEdit))
			_GUICtrlEdit_EndUpdate($_hLogEdit)
		EndIf
	EndIf

	; Append message to file
	If $LogFullPath <> "" And Not $bOverWriteLast Then
		If $_hLogFile = "" Then $_hLogFile = FileOpen($LogFullPath, $FO_APPEND)

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