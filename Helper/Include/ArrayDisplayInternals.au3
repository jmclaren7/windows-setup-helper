#include-once

#include "AutoItConstants.au3"
#include "MsgBoxConstants.au3"
#include "StringConstants.au3"

; #INDEX# =======================================================================================================================
; Title .........: Internal UDF Library for AutoIt3 _ArrayDisplay() and _DebugArrayDisplay()
; AutoIt Version : 3.3.18.0
; Description ...: Internal functions for the Array.au3 and Debug.au3
; Author(s) .....: Melba23, jpm, LarsJ, pixelsearch
; ===============================================================================================================================

#Region Global Variables and Constants

; #VARIABLES# ===================================================================================================================
; for use with the notify handler

Global $_g_bUserFunc_ArrayDisplay = False
Global $_g_hListView_ArrayDisplay
Global $_g_iTranspose_ArrayDisplay
Global $_g_iDisplayRow_ArrayDisplay
Global $_g_aArray_ArrayDisplay
Global $_g_iDims_ArrayDisplay
Global $_g_nRows_ArrayDisplay
Global $_g_nCols_ArrayDisplay
Global $_g_iItem_Start_ArrayDisplay
Global $_g_iItem_End_ArrayDisplay
Global $_g_iSubItem_Start_ArrayDisplay
Global $_g_iSubItem_End_ArrayDisplay
Global $_g_aIndex_ArrayDisplay
Global $_g_aIndexes_ArrayDisplay[1]
Global $_g_iSortDir_ArrayDisplay
Global $_g_asHeader_ArrayDisplay
Global $_g_aNumericSort_ArrayDisplay

Global $ARRAYDISPLAY_ROWPREFIX = "#"
Global $ARRAYDISPLAY_NUMERICSORT = "*"
; ===============================================================================================================================

; #CONSTANTS# ===================================================================================================================
Global Const $ARRAYDISPLAY_COLALIGNLEFT = 0 ; (default) Column text alignment - left
Global Const $ARRAYDISPLAY_TRANSPOSE = 1 ; Transposes the array (2D only)
Global Const $ARRAYDISPLAY_COLALIGNRIGHT = 2 ; Column text alignment - right
Global Const $ARRAYDISPLAY_COLALIGNCENTER = 4 ; Column text alignment - center
Global Const $ARRAYDISPLAY_VERBOSE = 8 ; Verbose - display MsgBox on error and splash screens during processing of large arrays
Global Const $ARRAYDISPLAY_NOROW = 64 ; No 'Row' column displayed
Global Const $ARRAYDISPLAY_CHECKERROR = 128 ; return if @error <> 0

Global Const $_ARRAYCONSTANT_tagLVITEM = "struct;uint Mask;int Item;int SubItem;uint State;uint StateMask;ptr Text;int TextMax;int Image;lparam Param;" & _
		"int Indent;int GroupID;uint Columns;ptr pColumns;ptr piColFmt;int iGroup;endstruct"
; ===============================================================================================================================

#EndRegion Global Variables and Constants

#Region Functions list

; #CURRENT# =====================================================================================================================
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __ArrayDisplay_Share
; __ArrayDisplay_SortIndexes
; __ArrayDisplay_NotifyHandler
; __ArrayDisplay_GetSortColStruct
; __ArrayDisplay_SortArrayStruct
; __ArrayDisplay_Transpose
; __ArrayDisplay_HeaderSetItemFormat
; __ArrayDisplay_GetItemText
; __ArrayDisplay_GetItemTextStringSelected
; __ArrayDisplay_JustifyColumn
; __ArrayDisplay_OnExit_CleanUp
; ===============================================================================================================================

#EndRegion Functions list

Func __ArrayDisplay_Share(Const ByRef $aArray, $sTitle = Default, $sArrayRange = Default, $iFlags = Default, $vUser_Separator = Default, $sHeader = Default, $iDesired_Colwidth = Default, $hUser_Function = Default, $bDebug = True, Const $_iScriptLineNumber = @ScriptLineNumber, Const $_iCallerError = @error, Const $_iCallerExtended = @extended)
	Local $sMsgBoxTitle = (($bDebug) ? ("_DebugArrayDisplay") : ("_ArrayDisplay"))

	; to avoid user_function using _DebugArrayDislay() recursion
	If $_g_bUserFunc_ArrayDisplay Then
		$hUser_Function = Default
		$bDebug = False
	EndIf
	If Not IsKeyword($hUser_Function) = $KEYWORD_DEFAULT Then
		$_g_bUserFunc_ArrayDisplay = True
	EndIf

	; Default values
	If $sTitle = Default Then $sTitle = $sMsgBoxTitle
	If $sArrayRange = Default Then $sArrayRange = ""
	If $iFlags = Default Then $iFlags = 0
	If $vUser_Separator = Default Then $vUser_Separator = ""
	If $sHeader = Default Then $sHeader = ""

	Local $iMin_ColWidth = 55
	Local $iMax_ColWidth = 350
;~  If $iDesired_Colwidth = Default Then ; do not change predefined $iMin_ColWidth, $iMax_ColWidth
	If $iDesired_Colwidth > 0 Then $iMax_ColWidth = $iDesired_Colwidth
	If $iDesired_Colwidth < 0 Then $iMin_ColWidth = -$iDesired_Colwidth
	If $iMax_ColWidth = Default Then $iMax_ColWidth = 350
	If $iMax_ColWidth > 4095 Then $iMax_ColWidth = 4095 ; needed as the structure inside the notify handler is declared as Static
	If $hUser_Function = Default Then $hUser_Function = 0

	; Check for transpose, column align, verbosity and "Row" column visibility
	$_g_iTranspose_ArrayDisplay = BitAND($iFlags, $ARRAYDISPLAY_TRANSPOSE)
	Local $iColAlign = BitAND($iFlags, 6) ; 0 = Left (default); 2 = Right; 4 = Center
	Local $iVerbose = Int(BitAND($iFlags, $ARRAYDISPLAY_VERBOSE))
	$_g_iDisplayRow_ArrayDisplay = Int(BitAND($iFlags, $ARRAYDISPLAY_NOROW) = 0)

	__ArrayDisplay_CheckArray_Range($iFlags, $iVerbose, $bDebug, $sMsgBoxTitle, $sTitle, $aArray, $sArrayRange, $_iScriptLineNumber, $_iCallerError)
	If @error Then Return SetError(@error, @extended, 0)

	#Region Check custom header

	; Determine copy separator
	Local $iCW_ColWidth = Number($vUser_Separator)

	; Get current separator character
	Local $sCurr_Separator = Opt("GUIDataSeparatorChar")

	; Set default user separator if required
	If $vUser_Separator = "" Then $vUser_Separator = $sCurr_Separator

	; Split custom header on separator
	$_g_asHeader_ArrayDisplay = StringSplit($sHeader, $sCurr_Separator, $STR_NOCOUNT) ; No count element
	If UBound($_g_asHeader_ArrayDisplay) = 0 Then Dim $_g_asHeader_ArrayDisplay[1] = [""]
	$sHeader = "Row"
	Local $iIndex = $_g_iSubItem_Start_ArrayDisplay
	If $_g_iTranspose_ArrayDisplay Then
		; All default headers
		$sHeader = "Row"
		For $j = 0 To $_g_nCols_ArrayDisplay - 1
			$sHeader &= $sCurr_Separator & $ARRAYDISPLAY_ROWPREFIX & " " & $j + $_g_iSubItem_Start_ArrayDisplay
		Next
	Else
		; Create custom header with available items
		If $_g_asHeader_ArrayDisplay[0] Then
			; Set as many as available
			For $iIndex = $_g_iSubItem_Start_ArrayDisplay To $_g_iSubItem_End_ArrayDisplay
				; Check custom header available
				If $iIndex >= UBound($_g_asHeader_ArrayDisplay) Then ExitLoop
				If StringRight($_g_asHeader_ArrayDisplay[$iIndex], 1) = $ARRAYDISPLAY_NUMERICSORT Then
					$_g_asHeader_ArrayDisplay[$iIndex] = StringTrimRight($_g_asHeader_ArrayDisplay[$iIndex], 1) ; remove "*" from right
					$_g_aNumericSort_ArrayDisplay[$iIndex - $_g_iSubItem_Start_ArrayDisplay] = 1 ; 1 (numeric sort) or empty (natural sort)
				EndIf

				$sHeader &= $sCurr_Separator & $_g_asHeader_ArrayDisplay[$iIndex]
			Next
		EndIf
		; Add default headers to fill to end
		For $j = $iIndex To $_g_iSubItem_End_ArrayDisplay
			$sHeader &= $sCurr_Separator & "Col " & $j
		Next
	EndIf
	; Remove "Row" header if not needed
	If Not $_g_iDisplayRow_ArrayDisplay Then $sHeader = StringTrimLeft($sHeader, 4)

	#EndRegion Check custom header

	#Region Generate Sort index for columns

	__ArrayDisplay_SortIndexes(0, -1)

	; compute the time to generate one colum info to the sorting
	Local $hTimer = TimerInit()
	__ArrayDisplay_SortIndexes(1, 1)
	Local $fTimer = TimerDiff($hTimer)
	If $fTimer * $_g_nCols_ArrayDisplay < 1000 Then
		; 		__ArrayDisplay_SortIndexes(-1)
		__ArrayDisplay_SortIndexes(2, $_g_nCols_ArrayDisplay)
;~ 		If $bDebug Then ConsoleWrite("Sorting all indexes = " & TimerDiff($hTimer) & @CRLF & @CRLF)
	Else
;~ 		If $bDebug Then ConsoleWrite("Sorting one index = " & TimerDiff($hTimer) & @CRLF)
	EndIf

	#EndRegion Generate Sort index for columns

	#Region GUI and Listview generation

	; Display splash dialog if required
	If $iVerbose And ($_g_nRows_ArrayDisplay * $_g_nCols_ArrayDisplay) > 1000 Then
		SplashTextOn($sMsgBoxTitle, "Preparing display" & @CRLF & @CRLF & "Please be patient", 300, 100)
	EndIf

	; GUI Constants
	Local Const $_ARRAYCONSTANT_GUI_DOCKBOTTOM = 64
	Local Const $_ARRAYCONSTANT_GUI_DOCKBORDERS = 102
	Local Const $_ARRAYCONSTANT_GUI_DOCKHEIGHT = 512
	Local Const $_ARRAYCONSTANT_GUI_DOCKLEFT = 2
	Local Const $_ARRAYCONSTANT_GUI_DOCKRIGHT = 4
	Local Const $_ARRAYCONSTANT_GUI_DOCKHCENTER = 8
	Local Const $_ARRAYCONSTANT_GUI_EVENT_CLOSE = -3
	Local Const $_ARRAYCONSTANT_GUI_EVENT_ARRAY = 1
	Local Const $_ARRAYCONSTANT_GUI_FOCUS = 256
	Local Const $_ARRAYCONSTANT_SS_CENTER = 0x1
	Local Const $_ARRAYCONSTANT_SS_CENTERIMAGE = 0x0200
	Local Const $_ARRAYCONSTANT_LVM_GETITEMRECT = (0x1000 + 14)
	Local Const $_ARRAYCONSTANT_LVM_GETITEMSTATE = (0x1000 + 44)
	Local Const $_ARRAYCONSTANT_LVM_GETSELECTEDCOUNT = (0x1000 + 50)
	Local Const $_ARRAYCONSTANT_LVM_SETEXTENDEDLISTVIEWSTYLE = (0x1000 + 54)
	Local Const $_ARRAYCONSTANT_LVS_EX_GRIDLINES = 0x1
	Local Const $_ARRAYCONSTANT_LVIS_SELECTED = 0x0002
	Local Const $_ARRAYCONSTANT_LVS_SHOWSELALWAYS = 0x8
	Local Const $_ARRAYCONSTANT_LVS_OWNERDATA = 0x1000
	Local Const $_ARRAYCONSTANT_LVS_EX_FULLROWSELECT = 0x20
	Local Const $_ARRAYCONSTANT_LVS_EX_DOUBLEBUFFER = 0x00010000 ; Paints via double-buffering, which reduces flicker
	Local Const $_ARRAYCONSTANT_WS_EX_CLIENTEDGE = 0x0200
	Local Const $_ARRAYCONSTANT_WS_MAXIMIZEBOX = 0x00010000
	Local Const $_ARRAYCONSTANT_WS_MINIMIZEBOX = 0x00020000
	Local Const $_ARRAYCONSTANT_WS_SIZEBOX = 0x00040000
	Local Const $_ARRAYCONSTANT_WS_EX_TOPMOST = 0x00000008

	; Set coord mode 1
	Local $iCoordMode = Opt("GUICoordMode", 1)

	; Set lower button border
	Local $iButtonBorder = (($bDebug) ? (40) : (20))

	; Create GUI
	Local $iOrgWidth = 210, $iHeight = 200, $iMinSize = 250
	Local $hGUI = GUICreate($sTitle, $iOrgWidth, $iHeight, Default, Default, BitOR($_ARRAYCONSTANT_WS_SIZEBOX, $_ARRAYCONSTANT_WS_MINIMIZEBOX, $_ARRAYCONSTANT_WS_MAXIMIZEBOX), $_ARRAYCONSTANT_WS_EX_TOPMOST)
	GUICtrlCreateLabel("@ArrayDisplayInternals@GUIidentifier@", 0, -10, 0, 0) ; 3.3.17.2 ; argumentum
	Local $aiGUISize = WinGetClientSize($hGUI)
	; Create ListView
	Local $idListView = GUICtrlCreateListView($sHeader, 0, 0, $aiGUISize[0], $aiGUISize[1] - $iButtonBorder, BitOR($_ARRAYCONSTANT_LVS_SHOWSELALWAYS, $_ARRAYCONSTANT_LVS_OWNERDATA))
	$_g_hListView_ArrayDisplay = GUICtrlGetHandle($idListView)
	GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_LVM_SETEXTENDEDLISTVIEWSTYLE, $_ARRAYCONSTANT_LVS_EX_GRIDLINES, $_ARRAYCONSTANT_LVS_EX_GRIDLINES)
	GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_LVM_SETEXTENDEDLISTVIEWSTYLE, $_ARRAYCONSTANT_LVS_EX_FULLROWSELECT, $_ARRAYCONSTANT_LVS_EX_FULLROWSELECT)
	GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_LVM_SETEXTENDEDLISTVIEWSTYLE, $_ARRAYCONSTANT_LVS_EX_DOUBLEBUFFER, $_ARRAYCONSTANT_LVS_EX_DOUBLEBUFFER)
	GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_LVM_SETEXTENDEDLISTVIEWSTYLE, $_ARRAYCONSTANT_WS_EX_CLIENTEDGE, $_ARRAYCONSTANT_WS_EX_CLIENTEDGE)
	Local $hHeader = HWnd(GUICtrlSendMsg($idListView, (0x1000 + 31), 0, 0)) ; $LVM_GETHEADER _GUICtrlListView_GetHeader($idListView)
	; Set resizing
	GUICtrlSetResizing($idListView, $_ARRAYCONSTANT_GUI_DOCKBORDERS)

	; Fill listview
	Local $iColFill = $_g_nCols_ArrayDisplay + $_g_iDisplayRow_ArrayDisplay
	; Align columns if required - $iColAlign = 2 for Right and 4 for Center
	If $iColAlign Then
		; Loop through columns
		For $i = 0 To $iColFill - 1
			__ArrayDisplay_JustifyColumn($idListView, $i, $iColAlign / 2)
		Next
	EndIf

	GUICtrlSendMsg($idListView, (0x1000 + 47), $_g_nRows_ArrayDisplay, 0) ; $LVM_SETITEMCOUNT

	; Get row height
	Local $tRECT = DllStructCreate("struct; long Left;long Top;long Right;long Bottom; endstruct") ; $tagRECT
	DllCall("user32.dll", "struct*", "SendMessageW", "hwnd", $_g_hListView_ArrayDisplay, "uint", $_ARRAYCONSTANT_LVM_GETITEMRECT, "wparam", 0, "struct*", $tRECT)
	; Set required GUI height
	Local $aiWin_Pos = WinGetPos($hGUI)
	Local $aiLV_Pos = ControlGetPos($hGUI, "", $idListView)
	$iHeight = (($_g_nRows_ArrayDisplay + 3) * (DllStructGetData($tRECT, "Bottom") - DllStructGetData($tRECT, "Top"))) + $aiWin_Pos[3] - $aiLV_Pos[3]
	; Check min/max height
	If $iHeight > @DesktopHeight - 100 Then
		$iHeight = @DesktopHeight - 100
	ElseIf $iHeight < $iMinSize Then
		$iHeight = $iMinSize
	EndIf

	If $iVerbose Then SplashOff()

	; Sorting information
	$_g_iSortDir_ArrayDisplay = 0x00000400 ; $HDF_SORTUP
	Local $iColumn = 0, $iColumnPrev = -1
	If $_g_iDisplayRow_ArrayDisplay Then
		$iColumnPrev = $iColumn
		__ArrayDisplay_HeaderSetItemFormat($hHeader, $iColumn, 0x00004000 + $_g_iSortDir_ArrayDisplay + $iColAlign / 2) ; $HDF_STRING
	EndIf
	$_g_aIndex_ArrayDisplay = $_g_aIndexes_ArrayDisplay[0]

	#EndRegion GUI and Listview generation

	; Register WM_NOTIFY message handler through subclassing
	Local $p__ArrayDisplay_NotifyHandler = DllCallbackGetPtr(DllCallbackRegister("__ArrayDisplay_NotifyHandler", "lresult", "hwnd;uint;wparam;lparam;uint_ptr;dword_ptr"))
	DllCall("comctl32.dll", "bool", "SetWindowSubclass", "hwnd", $hGUI, "ptr", $p__ArrayDisplay_NotifyHandler, "uint_ptr", 0, "dword_ptr", 0)   ; $iSubclassId = 0, $pData = 0

	#Region Adjust dialog width

	Local $iWidth = 40, $iColWidth = 0, $aiColWidth[$iColFill]
	; Get required column widths to fit items
	Local $iColWidthHeader, $iMin_ColW = 55
	For $i = 0 To $iColFill - 1
		If $i > 0 Then $iMin_ColW = $iMin_ColWidth ; to be use only for #col > 0
		GUICtrlSendMsg($idListView, (0x1000 + 30), $i, -1)                            ; $LVM_SETCOLUMNWIDTH $LVSCW_AUTOSIZE
		$iColWidth = GUICtrlSendMsg($idListView, (0x1000 + 29), $i, 0)                    ; $LVM_GETCOLUMNWIDTH
		; Check width of header if set
		If $sHeader <> "" Then
			If $iColWidth = 0 Then ExitLoop
			GUICtrlSendMsg($idListView, (0x1000 + 30), $i, -2)                            ; $LVM_SETCOLUMNWIDTH $LVSCW_AUTOSIZE_USEHEADER
			$iColWidthHeader = GUICtrlSendMsg($idListView, (0x1000 + 29), $i, 0)              ; $GETCOLUMNWIDTH
			; Set minimum if required
			If $iColWidth < $iMin_ColW And $iColWidthHeader < $iMin_ColW Then
				GUICtrlSendMsg($idListView, (0x1000 + 30), $i, $iMin_ColW)                ; $LVM_SETCOLUMNWIDTH
				$iColWidth = $iMin_ColW
			ElseIf $iColWidthHeader < $iColWidth Then
				GUICtrlSendMsg($idListView, (0x1000 + 30), $i, $iColWidth)                    ; $LVM_SETCOLUMNWIDTH
			Else
				$iColWidth = $iColWidthHeader
			EndIf
		Else
			; Set minimum if required
			If $iColWidth < $iMin_ColW Then
				GUICtrlSendMsg($idListView, (0x1000 + 30), $i, $iMin_ColW)                ; $LVM_SETCOLUMNWIDTH
				$iColWidth = $iMin_ColW
			EndIf
		EndIf
		; Add to total width
		$iWidth += $iColWidth
		; Store  value
		$aiColWidth[$i] = $iColWidth
	Next
	; Now check max size
	If $iWidth > @DesktopWidth - 100 Then
		; Apply max col width limit to reduce width
		$iWidth = 40
		For $i = 0 To $iColFill - 1
			If $aiColWidth[$i] > $iMax_ColWidth Then
				; Reset width
				GUICtrlSendMsg($idListView, (0x1000 + 30), $i, $iMax_ColWidth)        ; $LVM_SETCOLUMNWIDTH
				$iWidth += $iMax_ColWidth
			Else
				; Retain width
				$iWidth += $aiColWidth[$i]
			EndIf
			If $i < 20 And $bDebug Then ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $iWidth = ' & $iWidth & " $i = " & $i & @CRLF)                      ;### Debug Console
		Next
	EndIf

	; Check max/min width
	If $iWidth > @DesktopWidth - 100 Then
		$iWidth = @DesktopWidth - 100
	ElseIf $iWidth < $iMinSize Then
		$iWidth = $iMinSize
	EndIf

	#EndRegion Adjust dialog width

	; Allow for borders with vertical scrollbar
	Local $iScrollBarSize = 0
	If $iHeight = (@DesktopHeight - 100) Then $iScrollBarSize = 15

	; Resize dialog
	WinMove($hGUI, "", (@DesktopWidth - $iWidth + $iScrollBarSize) / 2, (@DesktopHeight - $iHeight) / 2, $iWidth + $iScrollBarSize, $iHeight)

	; Resize ListView
	$aiGUISize = WinGetClientSize($hGUI)
	GUICtrlSetPos($idListView, 0, 0, $iWidth, $aiGUISize[1] - $iButtonBorder)

	#Region Create bottom infos

	; Create data display
	Local $sDisplayData = "[" & $_g_nRows_ArrayDisplay & "]"
	If $_g_iDims_ArrayDisplay = 2 Then
		$sDisplayData &= " [" & $_g_nCols_ArrayDisplay & "]"
	EndIf

	; Create tooltip data
	Local $sTipData = ""
	If $sArrayRange Then
		If $sTipData Then $sTipData &= " - "
		$sTipData &= "Range set " & $sArrayRange
	EndIf
	If $_g_iTranspose_ArrayDisplay Then
		If $sTipData Then $sTipData &= " - "
		$sTipData &= "Transposed"
	EndIf

	Local $iButtonWidth_1 = $aiGUISize[0] / 2
	Local $iButtonWidth_2 = $aiGUISize[0] / 3
	Local $idBtn_Copy_ID = 9999, $idBtn_Copy_Data = 99999, $idLbl_Data = 99999, $idBtn_User_Func = 99999, $idBtn_Exit_Script = 99999
	If $bDebug Then
		; Create buttons
		$idBtn_Copy_ID = GUICtrlCreateButton("Copy Data && Hdr/Row", 0, $aiGUISize[1] - $iButtonBorder, $iButtonWidth_1, 20)
		$idBtn_Copy_Data = GUICtrlCreateButton("Copy Data Only", $iButtonWidth_1, $aiGUISize[1] - $iButtonBorder, $iButtonWidth_1, 20)
		Local $iButtonWidth_Var = $iButtonWidth_1
		Local $iOffset = $iButtonWidth_1
		If IsFunc($hUser_Function) Then
			; Create UserFunc button if function passed
			$idBtn_User_Func = GUICtrlCreateButton("Run User Func", $iButtonWidth_2, $aiGUISize[1] - 20, $iButtonWidth_2, 20)
			$iButtonWidth_Var = $iButtonWidth_2
			$iOffset = $iButtonWidth_2 * 2
		EndIf
		; Create Exit button and data label
		$idBtn_Exit_Script = GUICtrlCreateButton("Exit Script", $iOffset, $aiGUISize[1] - 20, $iButtonWidth_Var, 20)
		$idLbl_Data = GUICtrlCreateLabel($sDisplayData, 0, $aiGUISize[1] - 20, $iButtonWidth_Var, 18, BitOR($_ARRAYCONSTANT_SS_CENTER, $_ARRAYCONSTANT_SS_CENTERIMAGE))
	Else
		$idLbl_Data = GUICtrlCreateLabel($sDisplayData, 0, $aiGUISize[1] - 20, $aiGUISize[0], 18, BitOR($_ARRAYCONSTANT_SS_CENTER, $_ARRAYCONSTANT_SS_CENTERIMAGE))
	EndIf
	; Change label colour and create tooltip if required
	If $_g_iTranspose_ArrayDisplay Or $sArrayRange Then
		GUICtrlSetColor($idLbl_Data, 0xFF0000)
		GUICtrlSetTip($idLbl_Data, $sTipData)
	EndIf
	GUICtrlSetResizing($idBtn_Copy_ID, $_ARRAYCONSTANT_GUI_DOCKLEFT + $_ARRAYCONSTANT_GUI_DOCKBOTTOM + $_ARRAYCONSTANT_GUI_DOCKHEIGHT)
	GUICtrlSetResizing($idBtn_Copy_Data, $_ARRAYCONSTANT_GUI_DOCKRIGHT + $_ARRAYCONSTANT_GUI_DOCKBOTTOM + $_ARRAYCONSTANT_GUI_DOCKHEIGHT)
	GUICtrlSetResizing($idLbl_Data, $_ARRAYCONSTANT_GUI_DOCKLEFT + $_ARRAYCONSTANT_GUI_DOCKBOTTOM + $_ARRAYCONSTANT_GUI_DOCKHEIGHT)
	GUICtrlSetResizing($idBtn_User_Func, $_ARRAYCONSTANT_GUI_DOCKHCENTER + $_ARRAYCONSTANT_GUI_DOCKBOTTOM + $_ARRAYCONSTANT_GUI_DOCKHEIGHT)
	GUICtrlSetResizing($idBtn_Exit_Script, $_ARRAYCONSTANT_GUI_DOCKRIGHT + $_ARRAYCONSTANT_GUI_DOCKBOTTOM + $_ARRAYCONSTANT_GUI_DOCKHEIGHT)

	#EndRegion Create bottom infos

	; Display dialog
	GUISetState(@SW_SHOW, $hGUI)

	; Check if sort clicking can take a while
	If $fTimer > 1000 And Not $sArrayRange Then
		Beep(750, 250)
		ToolTip("Sorting Action can take as long as " & Ceiling($fTimer / 1000) & " sec" & @CRLF & @CRLF & "Please be patient when you click to sort a column", 50, 50, $sMsgBoxTitle, $TIP_WARNINGICON, $TIP_BALLOON)
		Sleep(3000)
		ToolTip("")
	EndIf

	#Region GUI Handling events

	; Switch to GetMessage mode
	Local $iOnEventMode = Opt("GUIOnEventMode", 0), $aMsg

	__ArrayDisplay_OnExit_CleanUp(1, $hGUI, $iCoordMode, $iOnEventMode, $_iCallerError, $_iCallerExtended, $p__ArrayDisplay_NotifyHandler)
	OnAutoItExitRegister(__ArrayDisplay_OnExit_CleanUp) ; 3.3.17.2 ; fix tray exit hung ; argumentum

	WinSetOnTop($hGUI, "", 1) ; 3.3.17.2 ; bring to front ; argumentum
	WinSetOnTop($hGUI, "", 0)

	While 1

		$aMsg = GUIGetMsg($_ARRAYCONSTANT_GUI_EVENT_ARRAY) ; Variable needed to check which "Copy" button was pressed
		If $aMsg[1] = $hGUI Then
			Switch $aMsg[0]
				Case $_ARRAYCONSTANT_GUI_EVENT_CLOSE
					ExitLoop

				Case $idBtn_Copy_ID, $idBtn_Copy_Data
					; Count selected rows
					Local $iSel_Count = GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_LVM_GETSELECTEDCOUNT, 0, 0)
					; Display splash dialog if required
					If $iVerbose And (Not $iSel_Count) And ($_g_iItem_End_ArrayDisplay - $_g_iItem_Start_ArrayDisplay) * ($_g_iSubItem_End_ArrayDisplay - $_g_iSubItem_Start_ArrayDisplay) > 10000 Then
						SplashTextOn($sMsgBoxTitle, "Copying data" & @CRLF & @CRLF & "Please be patient", 300, 100)
					EndIf
					; Generate clipboard text
					Local $sClip = "", $sItem, $aSplit, $iFirstCol = 0
					If $aMsg[0] = $idBtn_Copy_Data And $_g_iDisplayRow_ArrayDisplay Then $iFirstCol = 1
					; Add items
					For $i = 0 To GUICtrlSendMsg($idListView, 0X1004, 0, 0) - 1 ; $LVM_GETITEMCOUNT
						; Skip if copying selected rows and item not selected
						If $iSel_Count And Not (GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_LVM_GETITEMSTATE, $i, $_ARRAYCONSTANT_LVIS_SELECTED) <> 0) Then
							ContinueLoop
						EndIf
						$sItem = __ArrayDisplay_GetItemTextStringSelected($idListView, $i, $iFirstCol)
						If $aMsg[0] = $idBtn_Copy_ID And Not $_g_iDisplayRow_ArrayDisplay Then
							; Add row data
							$sItem = $ARRAYDISPLAY_ROWPREFIX & " " & ($i + $_g_iItem_Start_ArrayDisplay) & $sCurr_Separator & $sItem
						EndIf
						If $iCW_ColWidth Then
							; Expand columns
							$aSplit = StringSplit($sItem, $sCurr_Separator)
							$sItem = ""
							For $j = 1 To $aSplit[0]
								$sItem &= StringFormat("%-" & $iCW_ColWidth + 1 & "s", StringLeft($aSplit[$j], $iCW_ColWidth))
							Next
						Else
							; Use defined separator
							$sItem = StringReplace($sItem, $sCurr_Separator, $vUser_Separator)
						EndIf
						$sClip &= $sItem & @CRLF
					Next
					$sItem = $sHeader
					; Add header line if required
					If $aMsg[0] = $idBtn_Copy_ID Then
						$sItem = $sHeader
						If Not $_g_iDisplayRow_ArrayDisplay Then
							; Add "Row" to header
							$sItem = "Row" & $sCurr_Separator & $sItem
						EndIf
						If $iCW_ColWidth Then
							$aSplit = StringSplit($sItem, $sCurr_Separator)
							$sItem = ""
							For $j = 1 To $aSplit[0]
								$sItem &= StringFormat("%-" & $iCW_ColWidth + 1 & "s", StringLeft($aSplit[$j], $iCW_ColWidth))
							Next
						Else
							$sItem = StringReplace($sItem, $sCurr_Separator, $vUser_Separator)
						EndIf
						$sClip = $sItem & @CRLF & $sClip
					EndIf
					;Send to clipboard
					ClipPut($sClip)
					; Remove splash if used
					SplashOff()
					; Refocus ListView
					GUICtrlSetState($idListView, $_ARRAYCONSTANT_GUI_FOCUS)

				Case $idListView
					$iColumn = GUICtrlGetState($idListView)
					If Not IsArray($_g_aIndexes_ArrayDisplay[$iColumn + Not $_g_iDisplayRow_ArrayDisplay]) Then
						; in case all indexes have not been set during start creation
						__ArrayDisplay_SortIndexes($iColumn + Not $_g_iDisplayRow_ArrayDisplay)
					EndIf

					If $iColumn <> $iColumnPrev Then
						__ArrayDisplay_HeaderSetItemFormat($hHeader, $iColumnPrev, 0x00004000 + $iColAlign / 2) ; $HDF_STRING
						If $_g_iDisplayRow_ArrayDisplay And $iColumn = 0 Then
							$_g_aIndex_ArrayDisplay = $_g_aIndexes_ArrayDisplay[0]
						Else
							$_g_aIndex_ArrayDisplay = $_g_aIndexes_ArrayDisplay[$iColumn + Not $_g_iDisplayRow_ArrayDisplay]
						EndIf
					EndIf
					; $_g_iSortDir_ArrayDisplay = ($iColumn = $iColumnPrev) ? $_g_iSortDir_ArrayDisplay = $HDF_SORTUP ? $HDF_SORTDOWN : $HDF_SORTUP : $HDF_SORTUP
					$_g_iSortDir_ArrayDisplay = ($iColumn = $iColumnPrev) ? $_g_iSortDir_ArrayDisplay = 0x00000400 ? 0x00000200 : 0x00000400 : 0x00000400 ; $HDF_SORTUP
					__ArrayDisplay_HeaderSetItemFormat($hHeader, $iColumn, 0x00004000 + $_g_iSortDir_ArrayDisplay + $iColAlign / 2)  ; $HDF_STRING
					GUICtrlSendMsg($idListView, (0x1000 + 140), $iColumn, 0) ; $LVM_SETSELECTEDCOLUMN
					GUICtrlSendMsg($idListView, (0x1000 + 47), $_g_nRows_ArrayDisplay, 0) ; $LVM_SETITEMCOUNT
					$iColumnPrev = $iColumn

				Case $idBtn_User_Func
					; Get selected indices
					Local $aiSelItems[1] = [0]
					For $i = 0 To GUICtrlSendMsg($idListView, 0x1004, 0, 0) - 1 ; $LVM_GETITEMCOUNT
						If (GUICtrlSendMsg($idListView, $_ARRAYCONSTANT_LVM_GETITEMSTATE, $i, $_ARRAYCONSTANT_LVIS_SELECTED) <> 0) Then
							$aiSelItems[0] += 1
							ReDim $aiSelItems[$aiSelItems[0] + 1]
							$aiSelItems[$aiSelItems[0]] = $i + $_g_iItem_Start_ArrayDisplay
						EndIf
					Next

					; Pass array and selection to user function
					$hUser_Function($_g_aArray_ArrayDisplay, $aiSelItems)
					$_g_bUserFunc_ArrayDisplay = False

					OnAutoItExitUnRegister(__ArrayDisplay_OnExit_CleanUp) ; 3.3.17.2 ; fix tray exit hung ; argumentum
					__ArrayDisplay_CleanUp($hGUI, $iCoordMode, $iOnEventMode, $_iCallerError, $_iCallerExtended, $p__ArrayDisplay_NotifyHandler)
					;;__ArrayDisplay_Share($aArray, $sTitle, $sArrayRange, $iFlags, $vUser_Separator, $sHeader, $iMax_ColWidth, $hUser_Function, $bDebug, $_iScriptLineNumber, $_iCallerError, $_iCallerExtended)
					Return SetError($_iCallerError, $_iCallerExtended, -1)

				Case $idBtn_Exit_Script
					; Clear up
					GUIDelete($hGUI)
					Exit
			EndSwitch
		EndIf
	WEnd

	#EndRegion GUI Handling events

	OnAutoItExitUnRegister(__ArrayDisplay_OnExit_CleanUp) ; 3.3.17.2 ; fix tray exit hung ; argumentum
	__ArrayDisplay_CleanUp($hGUI, $iCoordMode, $iOnEventMode, $_iCallerError, $_iCallerExtended, $p__ArrayDisplay_NotifyHandler)

	Return SetError($_iCallerError, $_iCallerExtended, 1)
EndFunc   ;==>__ArrayDisplay_Share

Func __ArrayDisplay_CheckArray_Range($iFlags, $iVerbose, $bDebug, $sMsgBoxTitle, $sTitle, Const ByRef $aArray, $sArrayRange, Const $_iScriptLineNumber, Const $_iCallerError)

	#Region Check valid array

	Local $sMsg = "", $iError = 0
	If IsArray($aArray) Then
		$_g_aArray_ArrayDisplay = $aArray
		$_g_iDims_ArrayDisplay = UBound($_g_aArray_ArrayDisplay, $UBOUND_DIMENSIONS)
		If $_g_iDims_ArrayDisplay = 1 Then $_g_iTranspose_ArrayDisplay = 0
		$_g_nRows_ArrayDisplay = UBound($_g_aArray_ArrayDisplay, $UBOUND_ROWS)
		$_g_nCols_ArrayDisplay = ($_g_iDims_ArrayDisplay = 2) ? UBound($_g_aArray_ArrayDisplay, $UBOUND_COLUMNS) : 1

		; Split custom header on separator
		Dim $_g_aNumericSort_ArrayDisplay[$_g_nCols_ArrayDisplay]

		; Dimension checking
		If $_g_iDims_ArrayDisplay > 2 Then
			$sMsg = "Larger than 2D array passed to function"
			$iError = 2
		EndIf
		If $_iCallerError Then
			If BitAND($iFlags, $ARRAYDISPLAY_CHECKERROR) Then
				If $bDebug Then
					; Call _DebugReport() if available
					If IsDeclared("__g_sReportCallBack_DebugReport_Debug") Then
						$sMsg = "@@ Debug( " & $_iScriptLineNumber & ") : @error = " & $_iCallerError & " in " & $sMsgBoxTitle & "( '" & $sTitle & "' )"
						Execute('$__g_sReportCallBack_DebugReport_Debug("' & $sMsg & '")')
					EndIf
					$iError = 3
				Else
					$sMsg = "@error = " & $_iCallerError & " when calling the function"
					If $_iScriptLineNumber > 0 Then $sMsg &= " at line " & $_iScriptLineNumber
					$iError = 3
				EndIf
			EndIf
		EndIf
	Else
		$sMsg = "No array variable passed to function"
		$iError = 1
	EndIf
	If $iError Then
		If $iVerbose And MsgBox($MB_SYSTEMMODAL + $MB_ICONERROR + $MB_YESNO, _
				$sMsgBoxTitle & "() Error: " & $sTitle, $sMsg & @CRLF & @CRLF & "Exit the script?") = $IDYES Then
			Exit
		Else
			Return SetError($iError, 0, 0)
		EndIf
	EndIf

	#EndRegion Check valid array

	#Region Check array range

	; Declare variables
	$_g_iItem_Start_ArrayDisplay = 0
	$_g_iItem_End_ArrayDisplay = $_g_nRows_ArrayDisplay - 1
	$_g_iSubItem_Start_ArrayDisplay = 0
	$_g_iSubItem_End_ArrayDisplay = (($_g_iDims_ArrayDisplay = 2) ? ($_g_nCols_ArrayDisplay - 1) : (0))

	Local $avRangeSplit
	; Check for range settings
	If $sArrayRange Then
		; Split into separate dimension sections
		Local $vTmp, $aArray_Range = StringRegExp($sArrayRange & "||", "(?U)(.*)\|", $STR_REGEXPARRAYGLOBALMATCH)
		; Rows range
		If $aArray_Range[0] Then
			$avRangeSplit = StringSplit($aArray_Range[0], ":")
			If @error Then
				$_g_iItem_End_ArrayDisplay = Number($aArray_Range[0])
			Else
				$_g_iItem_Start_ArrayDisplay = Number($avRangeSplit[1])
				If $avRangeSplit[2] <> "" Then
					$_g_iItem_End_ArrayDisplay = Number($avRangeSplit[2])
				EndIf
			EndIf
		EndIf
		; Check row bounds
		If $_g_iItem_Start_ArrayDisplay < 0 Then $_g_iItem_Start_ArrayDisplay = 0
		If $_g_iItem_End_ArrayDisplay >= $_g_nRows_ArrayDisplay Then $_g_iItem_End_ArrayDisplay = $_g_nRows_ArrayDisplay - 1
		If ($_g_iItem_Start_ArrayDisplay > $_g_iItem_End_ArrayDisplay) And ($_g_iItem_End_ArrayDisplay > 0) Then
			$vTmp = $_g_iItem_Start_ArrayDisplay
			$_g_iItem_Start_ArrayDisplay = $_g_iItem_End_ArrayDisplay
			$_g_iItem_End_ArrayDisplay = $vTmp
		EndIf

		; Columns range
		If $_g_iDims_ArrayDisplay = 2 And $aArray_Range[1] Then
			$avRangeSplit = StringSplit($aArray_Range[1], ":")
			If @error Then
				$_g_iSubItem_End_ArrayDisplay = Number($aArray_Range[1])
			Else
				$_g_iSubItem_Start_ArrayDisplay = Number($avRangeSplit[1])
				If $avRangeSplit[2] <> "" Then
					$_g_iSubItem_End_ArrayDisplay = Number($avRangeSplit[2])
				EndIf
			EndIf
			; Check column bounds
			If $_g_iSubItem_Start_ArrayDisplay > $_g_iSubItem_End_ArrayDisplay Then
				$vTmp = $_g_iSubItem_Start_ArrayDisplay
				$_g_iSubItem_Start_ArrayDisplay = $_g_iSubItem_End_ArrayDisplay
				$_g_iSubItem_End_ArrayDisplay = $vTmp
			EndIf
			If $_g_iSubItem_Start_ArrayDisplay < 0 Then $_g_iSubItem_Start_ArrayDisplay = 0
			If $_g_iSubItem_End_ArrayDisplay >= $_g_nCols_ArrayDisplay Then $_g_iSubItem_End_ArrayDisplay = $_g_nCols_ArrayDisplay - 1
		EndIf
	EndIf

	If $sArrayRange Or $_g_iTranspose_ArrayDisplay Then $_g_aArray_ArrayDisplay = __ArrayDisplay_CreateSubArray()

	#EndRegion Check array range

EndFunc   ;==>__ArrayDisplay_CheckArray_Range

Func __ArrayDisplay_CleanUp($hGUI, $iCoordMode, $iOnEventMode, $_iCallerError, $_iCallerExtended, $p__ArrayDisplay_NotifyHandler)
	$_g_bUserFunc_ArrayDisplay = False ; to retore default behaviour

	; Cleanup
	DllCall("comctl32.dll", "bool", "RemoveWindowSubclass", "hwnd", $hGUI, "ptr", $p__ArrayDisplay_NotifyHandler, "uint_ptr", 0)   ; $iSubclassId = 0

	; Release resources in case of big array used
	$_g_aIndex_ArrayDisplay = 0
	Dim $_g_aIndexes_ArrayDisplay[1]

	GUIDelete($hGUI)
	Opt("GUICoordMode", $iCoordMode) ; Reset original Coord mode
	Opt("GUIOnEventMode", $iOnEventMode) ; Reset original GUI mode

	Return SetError($_iCallerError, $_iCallerExtended, 1)
EndFunc   ;==>__ArrayDisplay_CleanUp

; #ADDITIONAL Functions to speed up _ArrayDisplay() or _DebugArrayDisplay()# ====================================================
; Thanks Larsj

Func __ArrayDisplay_NotifyHandler($hWnd, $iMsg, $wParam, $lParam, $iSubclassId, $pData)
	If $iMsg <> 0x004E Then Return DllCall("comctl32.dll", "lresult", "DefSubclassProc", "hwnd", $hWnd, "uint", $iMsg, "wparam", $wParam, "lparam", $lParam)[0]   ; 0x004E = $WM_NOTIFY

	Local Static $tagNMHDR = "struct;hwnd hWndFrom;uint_ptr IDFrom;INT Code;endstruct"
	Local Static $tagNMLVDISPINFO = $tagNMHDR & ";" & $_ARRAYCONSTANT_tagLVITEM

	Local $tNMLVDISPINFO = DllStructCreate($tagNMLVDISPINFO, $lParam)
	Switch HWnd(DllStructGetData($tNMLVDISPINFO, "hWndFrom"))
		Case $_g_hListView_ArrayDisplay
			Switch DllStructGetData($tNMLVDISPINFO, "Code")
				Case -177 ; $LVN_GETDISPINFOW
					Local Static $tText = DllStructCreate("wchar[4096]"), $pText = DllStructGetPtr($tText)
					Local $iItem = DllStructGetData($tNMLVDISPINFO, "Item")
					Local $iRow = ($_g_iSortDir_ArrayDisplay = 0x00000400) ? $_g_aIndex_ArrayDisplay[$iItem] : $_g_aIndex_ArrayDisplay[$_g_nRows_ArrayDisplay - 1 - $iItem]
					Local $iCol = DllStructGetData($tNMLVDISPINFO, "SubItem")

;~ 					ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $iRow = ' & $iRow & '  $iCol= ' & $iCol & '  $iItem = ' & $iItem & @CRLF) ;### Debug Console

					Local $sTemp
					If $_g_iDisplayRow_ArrayDisplay = 0 Then
						$sTemp = __ArrayDisplay_GetData($iRow, $iCol)
						DllStructSetData($tText, 1, $sTemp)
						DllStructSetData($tNMLVDISPINFO, "Text", $pText)
					Else
						If $iCol = 0 Then
							If $_g_iTranspose_ArrayDisplay Then
								Local $sCaptionCplt = ""
								If $iRow + $_g_iItem_Start_ArrayDisplay < UBound($_g_asHeader_ArrayDisplay) _
										And StringStripWS($_g_asHeader_ArrayDisplay[$iRow + $_g_iItem_Start_ArrayDisplay], 1 + 2) <> "" Then
									$sCaptionCplt = " (" & StringStripWS($_g_asHeader_ArrayDisplay[$iRow + $_g_iItem_Start_ArrayDisplay], 1 + 2)
									If StringRight($sCaptionCplt, 1) = $ARRAYDISPLAY_NUMERICSORT Then $sCaptionCplt = StringTrimRight($sCaptionCplt, 1)
									$sCaptionCplt &= ")"
								EndIf
								DllStructSetData($tText, 1, "Col " & ($iRow + $_g_iItem_Start_ArrayDisplay) & $sCaptionCplt)
							Else
								DllStructSetData($tText, 1, $ARRAYDISPLAY_ROWPREFIX & " " & $iRow + $_g_iItem_Start_ArrayDisplay)
							EndIf
							DllStructSetData($tNMLVDISPINFO, "Text", $pText)
						Else
							$sTemp = __ArrayDisplay_GetData($iRow, $iCol - 1)
							DllStructSetData($tText, 1, $sTemp)
							DllStructSetData($tNMLVDISPINFO, "Text", $pText)
						EndIf
					EndIf
					Return

			EndSwitch
	EndSwitch

	Return DllCall("comctl32.dll", "lresult", "DefSubclassProc", "hwnd", $hWnd, "uint", $iMsg, "wparam", $wParam, "lparam", $lParam)[0]
	#forceref $iSubclassId, $pData
EndFunc   ;==>__ArrayDisplay_NotifyHandler

Func __ArrayDisplay_GetData($iRow, $iCol, $bSmall = False)
	Local $sTemp
	If $_g_iDims_ArrayDisplay = 2 Then
		$sTemp = $_g_aArray_ArrayDisplay[$iRow][$iCol]
	Else
		$sTemp = $_g_aArray_ArrayDisplay[$iRow]
	EndIf
	Switch VarGetType($sTemp)
		Case "Array"
			Local $sSubscript = ""
			For $i = 1 To UBound($sTemp, 0)
				$sSubscript &= "[" & UBound($sTemp, $i) & "]"
			Next

			$sTemp = "{Array" & $sSubscript & "}"
		Case "Map"
			$sTemp = "{Map[" & UBound($sTemp) & "]}"
		Case "Object"
			$sTemp = "{Object}"
	EndSwitch

	Local $iMax = (($bSmall) ? (35) : (4095))
	If StringLen($sTemp) > $iMax Then $sTemp = StringLeft($sTemp, $iMax - 4) & " ..."

	Return $sTemp
EndFunc   ;==>__ArrayDisplay_GetData

Func __ArrayDisplay_SortIndexes($iColStart, $iColEnd = $iColStart)
	Dim $_g_aIndex_ArrayDisplay[$_g_nRows_ArrayDisplay]
;~ 	Local $hTimer
	If $iColEnd = -1 Then
		; column (0) already sorted
		Dim $_g_aIndexes_ArrayDisplay[$_g_nCols_ArrayDisplay + $_g_iDisplayRow_ArrayDisplay + 1]
;~ 		$hTimer = TimerInit()

		For $i = 0 To $_g_nRows_ArrayDisplay - 1
			$_g_aIndex_ArrayDisplay[$i] = $i
		Next

		$_g_aIndexes_ArrayDisplay[0] = $_g_aIndex_ArrayDisplay
;~ 		ConsoleWrite("Sorting array col#0 = " & TimerDiff($hTimer) & @CRLF)
	EndIf

	If $iColStart = -1 Then
		; to index all columns
		$iColStart = 1
		$iColEnd = $_g_nCols_ArrayDisplay
	EndIf

	If $iColStart Then
		; Index aArray columns
		Local $tIndex
		For $i = $iColStart To $iColEnd
;~ 			$hTimer = TimerInit()
			$tIndex = __ArrayDisplay_GetSortColStruct($_g_aArray_ArrayDisplay, $i - 1)

			For $j = 0 To $_g_nRows_ArrayDisplay - 1
				$_g_aIndex_ArrayDisplay[$j] = DllStructGetData($tIndex, 1, $j + 1)
			Next

			$_g_aIndexes_ArrayDisplay[$i] = $_g_aIndex_ArrayDisplay
;~ 			If $i < 20 Then ConsoleWrite("Sorting array col#" & $i & " = " & TimerDiff($hTimer) & @CRLF)
		Next
	EndIf

EndFunc   ;==>__ArrayDisplay_SortIndexes

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __ArrayDisplay_GetSortColStruct
; Description ...: Index based sorting of a 2D array by one or more columns
; Syntax.........: __ArrayDisplay_GetSortColStruct( $aArray, $iCol )
; Parameters ....: $aArray The 2D array to be sorted by index
;					$iCol Index column used in sorting
; Return values .: the sorting index as a DllStruct of integers
; Author ........: Larj : FAS_Sort2DArrayAu3()
; Modified.......: jpm : extension to 1D Array
; Remarks .......:
; ===============================================================================================================================
Func __ArrayDisplay_GetSortColStruct(Const ByRef $aArray, $iCol)

	If UBound($aArray, $UBOUND_DIMENSIONS) < 1 Or UBound($aArray, $UBOUND_DIMENSIONS) > 2 Then
;~ 		ConsoleWrite("$aArray is not a 1D or 2D array variable" & @CRLF)
		Return SetError(6, 0, 0)
	EndIf

	Return __ArrayDisplay_SortArrayStruct($aArray, $iCol)
EndFunc   ;==>__ArrayDisplay_GetSortColStruct

Func __ArrayDisplay_SortArrayStruct(Const ByRef $aArray, $iCol)
	Local $iDims = UBound($aArray, $UBOUND_DIMENSIONS)
	Local $tIndex = DllStructCreate("uint[" & $_g_nRows_ArrayDisplay & "]")
	Local $pIndex = DllStructGetPtr($tIndex)
	Static $hDll = DllOpen("kernel32.dll")
	Static $hDllComp = DllOpen("shlwapi.dll")

	Local $iLo, $iHi, $iMi, $r, $nVal1, $nVal2

	; Sorting by one column
	For $i = 1 To $_g_nRows_ArrayDisplay - 1
		$iLo = 0
		$iHi = $i - 1
		Do
			$iMi = Int(($iLo + $iHi) / 2)
			If Not $_g_iTranspose_ArrayDisplay And $_g_aNumericSort_ArrayDisplay[$iCol] Then ; Numeric sort
				If $iDims = 1 Then
					$nVal1 = Number($aArray[$i])
					$nVal2 = Number($aArray[DllStructGetData($tIndex, 1, $iMi + 1)])
				Else
					$nVal1 = Number($aArray[$i][$iCol])
					$nVal2 = Number($aArray[DllStructGetData($tIndex, 1, $iMi + 1)][$iCol])
				EndIf
				$r = $nVal1 < $nVal2 ? -1 : $nVal1 > $nVal2 ? 1 : 0
			Else ; Natural sort
				If $iDims = 1 Then
					$r = DllCall($hDllComp, 'int', 'StrCmpLogicalW', 'wstr', String($aArray[$i]), 'wstr', String($aArray[DllStructGetData($tIndex, 1, $iMi + 1)]))[0]
				Else
					$r = DllCall($hDllComp, 'int', 'StrCmpLogicalW', 'wstr', String($aArray[$i][$iCol]), 'wstr', String($aArray[DllStructGetData($tIndex, 1, $iMi + 1)][$iCol]))[0]
				EndIf
			EndIf
			Switch $r
				Case -1
					$iHi = $iMi - 1
				Case 1
					$iLo = $iMi + 1
				Case 0
					ExitLoop
			EndSwitch
		Until $iLo > $iHi
		DllCall($hDll, "none", "RtlMoveMemory", "struct*", $pIndex + ($iMi + 1) * 4, "struct*", $pIndex + $iMi * 4, "ulong_ptr", ($i - $iMi) * 4)
		DllStructSetData($tIndex, 1, $i, $iMi + 1 + ($iLo = $iMi + 1))
	Next

	Return $tIndex
EndFunc   ;==>__ArrayDisplay_SortArrayStruct

Func __ArrayDisplay_CreateSubArray()
	; the returned subarray is transposed
	Local $nRows = $_g_iItem_End_ArrayDisplay - $_g_iItem_Start_ArrayDisplay + 1
	Local $nCols = $_g_iSubItem_End_ArrayDisplay - $_g_iSubItem_Start_ArrayDisplay + 1

	Local $iRow = -1, $iCol, $iTemp, $aTemp
	If $_g_iTranspose_ArrayDisplay Then
		Dim $aTemp[$nCols][$nRows]
		For $i = $_g_iItem_Start_ArrayDisplay To $_g_iItem_End_ArrayDisplay
			$iRow += 1
			$iCol = -1
			For $j = $_g_iSubItem_Start_ArrayDisplay To $_g_iSubItem_End_ArrayDisplay
				$iCol += 1
				$aTemp[$iCol][$iRow] = $_g_aArray_ArrayDisplay[$i][$j]
			Next
		Next

		$iTemp = $_g_iItem_Start_ArrayDisplay
		$_g_iItem_Start_ArrayDisplay = $_g_iSubItem_Start_ArrayDisplay
		$_g_iSubItem_Start_ArrayDisplay = $iTemp

		$iTemp = $_g_iItem_End_ArrayDisplay
		$_g_iItem_End_ArrayDisplay = $_g_iSubItem_End_ArrayDisplay
		$_g_iSubItem_End_ArrayDisplay = $iTemp

		$_g_nRows_ArrayDisplay = $nCols
		$_g_nCols_ArrayDisplay = $nRows
	Else
		If $_g_iDims_ArrayDisplay = 1 Then
			Dim $aTemp[$nRows]
			For $i = $_g_iItem_Start_ArrayDisplay To $_g_iItem_End_ArrayDisplay
				$iRow += 1
				$aTemp[$iRow] = $_g_aArray_ArrayDisplay[$i]
			Next
		Else
			Dim $aTemp[$nRows][$nCols]
			For $i = $_g_iItem_Start_ArrayDisplay To $_g_iItem_End_ArrayDisplay
				$iRow += 1
				$iCol = -1
				For $j = $_g_iSubItem_Start_ArrayDisplay To $_g_iSubItem_End_ArrayDisplay
					$iCol += 1
					$aTemp[$iRow][$iCol] = $_g_aArray_ArrayDisplay[$i][$j]
				Next
			Next

			$_g_nCols_ArrayDisplay = $nCols
		EndIf

		$_g_nRows_ArrayDisplay = $nRows
	EndIf

;~ 	_DebugArrayDisplay($aTemp, "Subarray")
	Return $aTemp
EndFunc   ;==>__ArrayDisplay_CreateSubArray

Func __ArrayDisplay_OnExit_CleanUp($iInit = 0, $hGUI = "", $iCoordMode = "", $iOnEventMode = "", $_iCallerError = "", $_iCallerExtended = "", $p__ArrayDisplay_NotifyHandler = "")
	#forceref $iInit
	Local Static $hGUI_Saved = "", $iCoordMode_Saved = "", $iOnEventMode_Saved = "", $_iCallerError_Saved = "", $_iCallerExtended_Saved = "", $_p__ArrayDisplay_NotifyHandler_Saved = ""
	If Int(Eval("iInit")) Then
		If IsHWnd($hGUI) Then
			$hGUI_Saved = $hGUI
			$iCoordMode_Saved = $iCoordMode
			$iOnEventMode_Saved = $iOnEventMode
			$_iCallerError_Saved = $_iCallerError
			$_iCallerExtended_Saved = $_iCallerExtended
			$_p__ArrayDisplay_NotifyHandler_Saved = $p__ArrayDisplay_NotifyHandler
		EndIf
	Else
		If $hGUI_Saved Then __ArrayDisplay_CleanUp($hGUI_Saved, $iCoordMode_Saved, $iOnEventMode_Saved, $_iCallerError_Saved, $_iCallerExtended_Saved, $_p__ArrayDisplay_NotifyHandler_Saved)
	EndIf
EndFunc   ;==>__ArrayDisplay_OnExit_CleanUp

; #DUPLICATED Functions to avoid big #include "GuiHeader.au3"# ==================================================================
; Functions have been simplified (unicode inprocess) according to __ArrayDisplay_Share() needs

Func __ArrayDisplay_HeaderSetItemFormat($hWnd, $iIndex, $iFormat)
	Local Static $tHDItem = DllStructCreate("uint Mask;int XY;ptr Text;handle hBMP;int TextMax;int Fmt;lparam Param;int Image;int Order;uint Type;ptr pFilter;uint State") ; $tagHDITEM
	DllStructSetData($tHDItem, "Mask", 0x00000004) ; $HDI_FORMAT
	DllStructSetData($tHDItem, "Fmt", $iFormat)
	Local $aResult = DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $hWnd, "uint", 0x120C, "wparam", $iIndex, "struct*", $tHDItem) ; $HDM_SETITEMW
	Return $aResult[0] <> 0
EndFunc   ;==>__ArrayDisplay_HeaderSetItemFormat

; #DUPLICATED Functions to avoid big #include "GuiListView.au3"# ================================================================
; Functions have been simplified (unicode inprocess) according to __ArrayDisplay_Share() needs

Func __ArrayDisplay_GetItemText($idListView, $iIndex, $iSubItem = 0)
	Local $tBuffer = DllStructCreate("wchar Text[4096]")
	Local $pBuffer = DllStructGetPtr($tBuffer)
	Local $tItem = DllStructCreate($_ARRAYCONSTANT_tagLVITEM)
	DllStructSetData($tItem, "SubItem", $iSubItem)
	DllStructSetData($tItem, "TextMax", 4096)
	DllStructSetData($tItem, "Text", $pBuffer)
	;Global Const $LVM_GETITEMTEXTW = (0x1000 + 115) ; 0X1073
	If IsHWnd($idListView) Then
		DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $idListView, "uint", 0x1073, "wparam", $iIndex, "struct*", $tItem)
	Else
		Local $pItem = DllStructGetPtr($tItem)
		GUICtrlSendMsg($idListView, 0x1073, $iIndex, $pItem)
	EndIf

	Return DllStructGetData($tBuffer, "Text")
EndFunc   ;==>__ArrayDisplay_GetItemText

Func __ArrayDisplay_GetItemTextStringSelected($idListView, $iItem, $iFirstCol)
	Local $sRow = "", $sSeparatorChar = Opt('GUIDataSeparatorChar')
	Local $iSelected = $iItem ; get row

	; GetColumnCount
	Local $hHeader = HWnd(GUICtrlSendMsg($idListView, 0x101F, 0, 0)) ; $LVM_GETHEADER
	Local $nCol = DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $hHeader, "uint", 0x1200, "wparam", 0, "lparam", 0)[0] ; $HDM_GETITEMCOUNT

	For $x = $iFirstCol To $nCol - 1
		$sRow &= __ArrayDisplay_GetItemText($idListView, $iSelected, $x) & $sSeparatorChar
	Next

	Return StringTrimRight($sRow, 1)
EndFunc   ;==>__ArrayDisplay_GetItemTextStringSelected

Func __ArrayDisplay_JustifyColumn($idListView, $iIndex, $iAlign = -1)
	;Local $aAlign[3] = [$LVCFMT_LEFT, $LVCFMT_RIGHT, $LVCFMT_CENTER]

	Local $tColumn = DllStructCreate("uint Mask;int Fmt;int CX;ptr Text;int TextMax;int SubItem;int Image;int Order;int cxMin;int cxDefault;int cxIdeal") ; $tagLVCOLUMN
	If $iAlign < 0 Or $iAlign > 2 Then $iAlign = 0
	DllStructSetData($tColumn, "Mask", 0x01) ; $LVCF_FMT
	DllStructSetData($tColumn, "Fmt", $iAlign)
	Local $pColumn = DllStructGetPtr($tColumn)
	Local $iRet = GUICtrlSendMsg($idListView, 0x1060, $iIndex, $pColumn)  ; $_ARRAYCONSTANT_LVM_SETCOLUMNW
	Return $iRet <> 0
EndFunc   ;==>__ArrayDisplay_JustifyColumn
