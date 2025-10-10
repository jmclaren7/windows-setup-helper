#include-once

#include "GuiCtrlInternals.au3"
#include "SendMessage.au3"
#include "StructureConstants.au3"    ; $tagPOINT
#include "ToolTipConstants.au3"      ; $TTM_ACTIVATE
#include "WinAPIConv.au3"            ; _WinAPI_MultiByteToWideChar()

#include "WinAPISysInternals.au3"    ; _WinAPI_IsClassName()

; #INDEX# =======================================================================================================================
; Title .........: ToolTip
; AutoIt Version : 3.3.18.0
; Description ...: Functions that assist with ToolTip control management.
;                  ToolTip controls are pop-up windows that display text.  The text usually describes a tool, which is  either  a
;                  window, such as a child window or control, or an application-defined rectangular area within a window's client
;                  area.
; Author(s) .....: Paul Campbell (PaulIA)
; ===============================================================================================================================

; #VARIABLES# ===================================================================================================================
Global $__g_tTTBuffer = DllStructCreate("wchar Text[4096]")
; ===============================================================================================================================

; #CONSTANTS# ===================================================================================================================
Global Const $_TOOLTIPCONSTANTS_ClassName = "tooltips_class32"
Global Const $_TT_ghTTDefaultStyle = BitOR($TTS_ALWAYSTIP, $TTS_NOPREFIX)
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _GUIToolTip_Activate
; _GUIToolTip_AddTool
; _GUIToolTip_AdjustRect
; _GUIToolTip_Deactivate
; _GUIToolTip_BitsToTTF
; _GUIToolTip_Create
; _GUIToolTip_Deactivate
; _GUIToolTip_DelTool
; _GUIToolTip_Destroy
; _GUIToolTip_EnumTools
; _GUIToolTip_GetBubbleHeight
; _GUIToolTip_GetBubbleSize
; _GUIToolTip_GetBubbleWidth
; _GUIToolTip_GetCurrentTool
; _GUIToolTip_GetDelayTime
; _GUIToolTip_GetMargin
; _GUIToolTip_GetMarginEx
; _GUIToolTip_GetMaxTipWidth
; _GUIToolTip_GetText
; _GUIToolTip_GetTipBkColor
; _GUIToolTip_GetTipTextColor
; _GUIToolTip_GetTitleBitMap
; _GUIToolTip_GetTitleText
; _GUIToolTip_GetToolCount
; _GUIToolTip_GetToolInfo
; _GUIToolTip_HitTest
; _GUIToolTip_NewToolRect
; _GUIToolTip_Pop
; _GUIToolTip_PopUp
; _GUIToolTip_SetDelayTime
; _GUIToolTip_SetMargin
; _GUIToolTip_SetMaxTipWidth
; _GUIToolTip_SetTipBkColor
; _GUIToolTip_SetTipTextColor
; _GUIToolTip_SetTitle
; _GUIToolTip_SetToolInfo
; _GUIToolTip_SetWindowTheme
; _GUIToolTip_ToolExists
; _GUIToolTip_ToolToArray
; _GUIToolTip_TrackActivate
; _GUIToolTip_TrackPosition
; _GUIToolTip_Update
; _GUIToolTip_UpdateTipText
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; $tagNMTTDISPINFO
; $tagTOOLINFO
; $tagTTGETTITLE
; $tagTTHITTESTINFO
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: $tagNMTTDISPINFO
; Description ...: Contains information used in handling the $TTN_GETDISPINFOW notification message
; Fields ........: $tagNMHDR - Contains information about a notification message
;                  pText     - Pointer to a string that will be displayed as the ToolTip text.  If Instance specifies an instance
;                  +handle, this member must be the identifier of a string resource.
;                  aText     - Buffer that receives the ToolTip text.  An application can copy the text to this buffer instead of
;                  +specifying a string address or string resource.
;                  Instance  - Handle to the instance that contains a string resource to be used as the ToolTip text. If pText is
;                  +the address of the ToolTip text string, this member must be 0.
;                  Flags     - Flags that indicates how to interpret the IDFrom member:
;                  |$TTF_IDISHWND   - If this flag is set, IDFrom is the tool's handle. Otherwise, it is the tool's identifier.
;                  |$TTF_RTLREADING - Specifies right to left text
;                  |$TTF_DI_SETITEM - If you add this flag to Flags while processing the notification, the ToolTip  control  will
;                  +retain the supplied information and not request it again.
;                  Param     - Application-defined data associated with the tool
; Author ........: Paul Campbell (PaulIA)
; Remarks .......: You need to point the pText array to your own private buffer when the text used in the ToolTip text exceeds 80
;                  +characters in length.  The system automatically strips the accelerator from all strings passed to  a  ToolTip
;                  control, unless the control has the $TTS_NOPREFIX style.
; ===============================================================================================================================
Global Const $tagNMTTDISPINFO = $tagNMHDR & ";ptr pText;wchar aText[80];ptr Instance;uint Flags;lparam Param"

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: $tagTOOLINFO
; Description ...: Contains information about a tool in a ToolTip contr
; Fields ........: Size    - Size of this structure, in bytes
;                  Flags    - Flags that control the ToolTip display. This member can be a combination of the following values:
;                  |$TTF_ABSOLUTE    - Positions the ToolTip at the same coordinates provided by $TTM_TRACKPOSITION
;                  |$TTF_CENTERTIP   - Centers the ToolTip below the tool specified by the ID member
;                  |$TTF_IDISHWND    - Indicates that the ID member is the window handle to the tool
;                  |$TTF_PARSELINKS  - Indicates that links in the tooltip text should be parsed
;                  |$TTF_RTLREADING  - Indicates that the ToolTip text will be displayed in the opposite direction
;                  |$TTF_SUBCLASS    - Indicates that the ToolTip control should subclass the tool's window to intercept messages
;                  |$TTF_TRACK       - Positions the ToolTip next to the tool to which it corresponds
;                  |$TTF_TRANSPARENT - Causes the ToolTip control to forward mouse event messages to the parent window
;                  hWnd     - Handle to the window that contains the tool
;                  ID       - Application-defined identifier of the tool
;                  Left     - X position of upper left corner of bounding rectangle
;                  Top      - Y position of upper left corner of bounding rectangle
;                  Right    - X position of lower right corner of bounding rectangle
;                  Bottom   - Y position of lower right corner of bounding rectangle
;                  hInst    - Handle to the instance that contains the string resource for the too
;                  Text     - Pointer to the buffer that contains the text for the tool
;                  Param    - A 32-bit application-defined value that is associated with the tool
;                  Reserved - Reserved
; Author ........: Paul Campbell (PaulIA)
; Remarks .......:
; ===============================================================================================================================
Global Const $tagTOOLINFO = "struct; uint Size;uint Flags;hwnd hWnd;uint_ptr ID;" & $tagRECT & ";handle hInst;ptr Text;lparam Param;ptr Reserved; endstruct"

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: $tagTTGETTITLE
; Description ...: Provides information about the title of a tooltip control
; Fields ........: Size     - Size of this structure, in bytes
;                  Bitmap   - The tooltip icon
;                  TitleMax - Specifies the number of characters in the title
;                  Title    - Pointer to a wide character string that contains the title
; Author ........: Paul Campbell (PaulIA)
; Remarks .......:
; ===============================================================================================================================
Global Const $tagTTGETTITLE = "dword Size;uint Bitmap;uint TitleMax;ptr Title"

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: $tagTTHITTESTINFO
; Description ...: Contains information that a ToolTip control uses to determine whether a point is in the bounding rectangle of the specified tool
; Fields ........: Tool     - Handle to the tool or window with the specified tool
;                  X        - X position to be tested, in client coordinates
;                  Y        - Y position to be tested, in client coordinates
;                  Size     - Size of a TOOLINFO structure
;                  Flags    - Flags that control the ToolTip display. This member can be a combination of the following values:
;                  |$TTF_ABSOLUTE    - Positions the ToolTip at the same coordinates provided by $TTM_TRACKPOSITION
;                  |$TTF_CENTERTIP   - Centers the ToolTip below the tool specified by the ID member
;                  |$TTF_IDISHWND    - Indicates that the ID member is the window handle to the tool
;                  |$TTF_PARSELINKS  - Indicates that links in the tooltip text should be parsed
;                  |$TTF_RTLREADING  - Indicates that the ToolTip text will be displayed in the opposite direction
;                  |$TTF_SUBCLASS    - Indicates that the ToolTip control should subclass the tool's window to intercept messages
;                  |$TTF_TRACK       - Positions the ToolTip next to the tool to which it corresponds
;                  |$TTF_TRANSPARENT - Causes the ToolTip control to forward mouse event messages to the parent window
;                  hWnd     - Handle to the window that contains the tool
;                  ID       - Application-defined identifier of the tool
;                  Left     - X position of upper left corner of bounding rectangle
;                  Top      - Y position of upper left corner of bounding rectangle
;                  Right    - X position of lower right corner of bounding rectangle
;                  Bottom   - Y position of lower right corner of bounding rectangle
;                  hInst    - Handle to the instance that contains the string resource for the too
;                  Text     - Pointer to the buffer that contains the text for the tool
;                  Param    - A 32-bit application-defined value that is associated with the tool
;                  Reserved - Reserved
; Author ........: Paul Campbell (PaulIA)
; Remarks .......:
; ===============================================================================================================================
Global Const $tagTTHITTESTINFO = "hwnd Tool;" & $tagPOINT & ";" & $tagTOOLINFO

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......: Bob Marotte (BrewManNH)
; ===============================================================================================================================
Func _GUIToolTip_Activate($hTool)
	_SendMessage($hTool, $TTM_ACTIVATE, True)
EndFunc   ;==>_GUIToolTip_Activate

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......: Bob Marotte (BrewManNH)
; ===============================================================================================================================
Func _GUIToolTip_AddTool($hTool, $hWnd, $sText, $iID = 0, $iLeft = 0, $iTop = 0, $iRight = 0, $iBottom = 0, $iFlags = Default, $iParam = 0, $hInst = 0)
	If __GUICtrl_CheckHandleOutProcess($hWnd) Then
		Return SetError(6, 0, "") ; Processes not in same AutoIt Mode
	EndIf

	Local $tBuffer, $pBuffer
	If $iFlags = Default Then $iFlags = BitOR($TTF_SUBCLASS, $TTF_IDISHWND)
	If $sText <> -1 Then
		$tBuffer = $__g_tTTBuffer
		$pBuffer = DllStructGetPtr($tBuffer)
		DllStructSetData($tBuffer, "Text", $sText)
	Else
		$tBuffer = 0
		$pBuffer = -1 ; LPSTR_TEXTCALLBACK
	EndIf
	Local $tToolInfo = DllStructCreate($tagTOOLINFO)
	Local $iToolInfo = DllStructGetSize($tToolInfo)
	DllStructSetData($tToolInfo, "Size", $iToolInfo)
	DllStructSetData($tToolInfo, "Flags", $iFlags)
	DllStructSetData($tToolInfo, "hWnd", $hWnd)
	DllStructSetData($tToolInfo, "ID", $iID)
	DllStructSetData($tToolInfo, "Left", $iLeft)
	DllStructSetData($tToolInfo, "Top", $iTop)
	DllStructSetData($tToolInfo, "Right", $iRight)
	DllStructSetData($tToolInfo, "Bottom", $iBottom)
	DllStructSetData($tToolInfo, "hInst", $hInst)
	DllStructSetData($tToolInfo, "Param", $iParam)
	DllStructSetData($tToolInfo, "Text", $pBuffer)
	Local $iRet = __GUICtrl_SendMsg($hTool, $TTM_ADDTOOLW, 0, $tToolInfo, $tBuffer, False, 10, False, -1)

	Return $iRet <> 0
EndFunc   ;==>_GUIToolTip_AddTool

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_AdjustRect($hTool, ByRef $tRECT, $bLarger = True)
	__GUICtrl_SendMsg($hTool, $TTM_ADJUSTRECT, $bLarger, $tRECT, 0, True)

	Return $tRECT
EndFunc   ;==>_GUIToolTip_AdjustRect

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......: Bob Marotte (BrewManNH)
; ===============================================================================================================================
Func _GUIToolTip_BitsToTTF($iFlags)
	Local $iN = ""
	If BitAND($iFlags, $TTF_IDISHWND) <> 0 Then $iN &= "TTF_IDISHWND,"
	If BitAND($iFlags, $TTF_CENTERTIP) <> 0 Then $iN &= "TTF_CENTERTIP,"
	If BitAND($iFlags, $TTF_RTLREADING) <> 0 Then $iN &= "TTF_RTLREADING,"
	If BitAND($iFlags, $TTF_SUBCLASS) <> 0 Then $iN &= "TTF_SUBCLASS,"
	If BitAND($iFlags, $TTF_TRACK) <> 0 Then $iN &= "TTF_TRACK,"
	If BitAND($iFlags, $TTF_ABSOLUTE) <> 0 Then $iN &= "TTF_ABSOLUTE,"
	If BitAND($iFlags, $TTF_TRANSPARENT) <> 0 Then $iN &= "TTF_TRANSPARENT,"
	If BitAND($iFlags, $TTF_PARSELINKS) <> 0 Then $iN &= "TTF_PARSELINKS,"
	Return StringTrimRight($iN, 1)
EndFunc   ;==>_GUIToolTip_BitsToTTF

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......: Gary Frost
; ===============================================================================================================================
Func _GUIToolTip_Create($hWnd, $iStyle = $_TT_ghTTDefaultStyle)
	If $hWnd And Not IsHWnd($hWnd) Then Return SetError(1, 0, 0) ; Invalid Window handle for _GUIToolTip_Create 1st parameter

	If $hWnd And DllCall("user32.dll", "dword", "GetWindowThreadProcessId", "hwnd", $hWnd, "dword*", 0)[2] <> @AutoItPID Then ; Not _WinAPI_InProcess
		Return SetError(4, 0, 0)
	EndIf

	Return _WinAPI_CreateWindowEx($hWnd, $_TOOLTIPCONSTANTS_ClassName, "", $iStyle, 0, 0, 0, 0, $hWnd)
EndFunc   ;==>_GUIToolTip_Create

; #FUNCTION# ====================================================================================================================
; Author ........: Bob Marotte (BrewManNH)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_Deactivate($hTool)
	_SendMessage($hTool, $TTM_ACTIVATE, False)
EndFunc   ;==>_GUIToolTip_Deactivate

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......: Bob Marotte (BrewManNH)
; ===============================================================================================================================
Func _GUIToolTip_DelTool($hTool, $hWnd, $iID = 0)
	If __GUICtrl_CheckHandleOutProcess($hWnd) Then
		Return SetError(6, 0, "") ; Processes not in same AutoIt Mode
	EndIf

	Local $tToolInfo = DllStructCreate($tagTOOLINFO)
	Local $iToolInfo = DllStructGetSize($tToolInfo)
	DllStructSetData($tToolInfo, "Size", $iToolInfo)
	DllStructSetData($tToolInfo, "ID", $iID)
	DllStructSetData($tToolInfo, "hWnd", $hWnd)
	__GUICtrl_SendMsg($hTool, $TTM_DELTOOLW, 0, $tToolInfo)

EndFunc   ;==>_GUIToolTip_DelTool

; #FUNCTION# ====================================================================================================================
; Author ........: Gary Frost
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_Destroy(ByRef $hTool)
	If Not _WinAPI_IsClassName($hTool, $_TOOLTIPCONSTANTS_ClassName) Then Return SetError(2, 0, False)

	Local $iDestroyed = 0
	If IsHWnd($hTool) Then
		If DllCall("user32.dll", "dword", "GetWindowThreadProcessId", "hwnd", $hTool, "dword*", 0)[2] = @AutoItPID Then ; _WinAPI_InProcess
			$iDestroyed = _WinAPI_DestroyWindow($hTool)
		Else
			; Not Allowed to Destroy Other Applications Control(s)
			Return SetError(1, 0, False)
		EndIf
	EndIf
	If $iDestroyed Then $hTool = 0

	Return $iDestroyed <> 0
EndFunc   ;==>_GUIToolTip_Destroy

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_EnumTools($hTool, $iIndex)
	If __GUICtrl_CheckHandleOutProcess($hTool) Then
		Return SetError(6, 0, "") ; Processes not in same AutoIt Mode
	EndIf

	Local $tBuffer = $__g_tTTBuffer

	Local $tToolInfo = DllStructCreate($tagTOOLINFO)
	Local $iToolInfo = DllStructGetSize($tToolInfo)
	DllStructSetData($tToolInfo, "Size", $iToolInfo)
;~ 	Local $bResult = __GUICtrl_SendMsg($hTool, $TTM_ENUMTOOLSW, $iIndex, $tToolInfo, 0, True)
	Local $bResult = __GUICtrl_SendMsg($hTool, $TTM_ENUMTOOLSW, $iIndex, $tToolInfo, $tBuffer, True, 10, True, -1)

	Return _GUIToolTip_ToolToArray($hTool, $tToolInfo, ($bResult = True))
EndFunc   ;==>_GUIToolTip_EnumTools

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......: Bob Marotte (BrewManNH)
; ===============================================================================================================================
Func _GUIToolTip_GetBubbleHeight($hTool, $hWnd, $iID, $iFlags = Default)
	If $iFlags = Default Then $iFlags = BitOR($TTF_IDISHWND, $TTF_SUBCLASS)
	Local $iHeight = _GUIToolTip_GetBubbleSize($hTool, $hWnd, $iID, $iFlags)
	Local $iError = @error
	$iHeight = _WinAPI_HiWord($iHeight)
	Return SetError($iError, @extended, $iHeight)
EndFunc   ;==>_GUIToolTip_GetBubbleHeight

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......: Bob Marotte (BrewManNH)
; ===============================================================================================================================
Func _GUIToolTip_GetBubbleSize($hTool, $hWnd, $iID, $iFlags = Default)
	If __GUICtrl_CheckHandleOutProcess($hWnd) Then
		Return SetError(6, 0, "") ; Processes not in same AutoIt Mode
	EndIf

	If $iFlags = Default Then $iFlags = BitOR($TTF_IDISHWND, $TTF_SUBCLASS)
	Local $tToolInfo = DllStructCreate($tagTOOLINFO)
	Local $iToolInfo = DllStructGetSize($tToolInfo)
	DllStructSetData($tToolInfo, "Size", $iToolInfo)
	DllStructSetData($tToolInfo, "hWnd", $hWnd)
	DllStructSetData($tToolInfo, "ID", $iID)
	DllStructSetData($tToolInfo, "Flags", $iFlags)
	Local $iRet = __GUICtrl_SendMsg($hTool, $TTM_GETBUBBLESIZE, 0, $tToolInfo)

	Return $iRet
EndFunc   ;==>_GUIToolTip_GetBubbleSize

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......: Bob Marotte (BrewManNH)
; ===============================================================================================================================
Func _GUIToolTip_GetBubbleWidth($hTool, $hWnd, $iID, $iFlags = Default)
	If $iFlags = Default Then $iFlags = BitOR($TTF_IDISHWND, $TTF_SUBCLASS)
	Local $iWidth = _GUIToolTip_GetBubbleSize($hTool, $hWnd, $iID, $iFlags)
	Local $iError = @error
	$iWidth = _WinAPI_LoWord($iWidth)
	Return SetError($iError, @extended, $iWidth)
EndFunc   ;==>_GUIToolTip_GetBubbleWidth

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_GetCurrentTool($hTool)
	If __GUICtrl_CheckHandleOutProcess($hTool) Then
		Return SetError(6, 0, "") ; Processes not in same AutoIt Mode
	EndIf

	Local $tToolInfo = DllStructCreate($tagTOOLINFO)
	Local $iToolInfo = DllStructGetSize($tToolInfo)
	DllStructSetData($tToolInfo, "Size", $iToolInfo)
	Local $bResult = __GUICtrl_SendMsg($hTool, $TTM_GETCURRENTTOOLW, 0, $tToolInfo, 0, True)

	Return _GUIToolTip_ToolToArray($hTool, $tToolInfo, $bResult = True)
EndFunc   ;==>_GUIToolTip_GetCurrentTool

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......: Bob Marotte (BrewManNH)
; ===============================================================================================================================
Func _GUIToolTip_GetDelayTime($hTool, $iDuration)
	Return _SendMessage($hTool, $TTM_GETDELAYTIME, $iDuration)
EndFunc   ;==>_GUIToolTip_GetDelayTime

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_GetMargin($hTool)
	Local $aMargin[4]

	Local $tRECT = _GUIToolTip_GetMarginEx($hTool)
	$aMargin[0] = DllStructGetData($tRECT, "Left")
	$aMargin[1] = DllStructGetData($tRECT, "Top")
	$aMargin[2] = DllStructGetData($tRECT, "Right")
	$aMargin[3] = DllStructGetData($tRECT, "Bottom")

	Return $aMargin
EndFunc   ;==>_GUIToolTip_GetMargin

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_GetMarginEx($hTool)
	Local $tRECT = DllStructCreate($tagRECT)
	__GUICtrl_SendMsg($hTool, $TTM_GETMARGIN, 0, $tRECT, 0, True)

	Return $tRECT
EndFunc   ;==>_GUIToolTip_GetMarginEx

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_GetMaxTipWidth($hTool)
	Return _SendMessage($hTool, $TTM_GETMAXTIPWIDTH)
EndFunc   ;==>_GUIToolTip_GetMaxTipWidth

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_GetText($hTool, $hWnd, $iID)
	If __GUICtrl_CheckHandleOutProcess($hWnd) Then
		Return SetError(6, 0, "") ; Processes not in same AutoIt Mode
	EndIf

	Local $tBuffer = $__g_tTTBuffer

	Local $tToolInfo = DllStructCreate($tagTOOLINFO)
	Local $iToolInfo = DllStructGetSize($tToolInfo)
	DllStructSetData($tToolInfo, "Size", $iToolInfo)
	DllStructSetData($tToolInfo, "hWnd", $hWnd)
	DllStructSetData($tToolInfo, "ID", $iID)
	__GUICtrl_SendMsg($hTool, $TTM_GETTEXTW, DllStructGetSize($tBuffer), $tToolInfo, $tBuffer, False, 10, True, -1)

	Return DllStructGetData($tBuffer, "Text")
EndFunc   ;==>_GUIToolTip_GetText

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_GetTipBkColor($hTool)
	Return _SendMessage($hTool, $TTM_GETTIPBKCOLOR)
EndFunc   ;==>_GUIToolTip_GetTipBkColor

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_GetTipTextColor($hTool)
	Return _SendMessage($hTool, $TTM_GETTIPTEXTCOLOR)
EndFunc   ;==>_GUIToolTip_GetTipTextColor

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_GetTitleBitMap($hTool)
	If __GUICtrl_CheckHandleOutProcess($hTool) Then
		Return SetError(6, 0, "") ; Processes not in same AutoIt Mode JPM to check
	EndIf

	Local $tBuffer = $__g_tTTBuffer
	Local $tTitle = DllStructCreate($tagTTGETTITLE)
	Local $iTitle = DllStructGetSize($tTitle)
	DllStructSetData($tTitle, "Size", $iTitle)
	DllStructSetData($tTitle, "TitleMax", DllStructGetSize($tBuffer))
	__GUICtrl_SendMsg($hTool, $TTM_GETTITLE, 0, $tTitle, $tBuffer, True, 4, False, -1)

	Return DllStructGetData($tTitle, "Bitmap")
EndFunc   ;==>_GUIToolTip_GetTitleBitMap

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_GetTitleText($hTool)
	Local $tBuffer = $__g_tTTBuffer
	Local $tTitle = DllStructCreate($tagTTGETTITLE)
	Local $iTitle = DllStructGetSize($tTitle)
	DllStructSetData($tTitle, "TitleMax", DllStructGetSize($tBuffer))
	DllStructSetData($tTitle, "Size", $iTitle)
	__GUICtrl_SendMsg($hTool, $TTM_GETTITLE, 0, $tTitle, $tBuffer, False, 4, True, -1)

	Return DllStructGetData($tBuffer, "Text")
EndFunc   ;==>_GUIToolTip_GetTitleText

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_GetToolCount($hTool)
	Return _SendMessage($hTool, $TTM_GETTOOLCOUNT)
EndFunc   ;==>_GUIToolTip_GetToolCount

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_GetToolInfo($hTool, $hWnd, $iID)
	If __GUICtrl_CheckHandleOutProcess($hWnd) Then
		Return SetError(6, 0, "") ; Processes not in same AutoIt Mode
	EndIf

	Local $aTool, $iCount = _GUIToolTip_GetToolCount($hTool)
	For $i = 1 To $iCount
		; to go around for Flag not return by $TTM_GETTOOLINFOW
		$aTool = _GUIToolTip_EnumTools($hTool, $i - 1)
		If $aTool[2] = $iID Then Return $aTool
	Next

	Local $tBuffer = $__g_tTTBuffer

	Local $tToolInfo = DllStructCreate($tagTOOLINFO)
	Local $iToolInfo = DllStructGetSize($tToolInfo)
	DllStructSetData($tToolInfo, "Size", $iToolInfo)
	DllStructSetData($tToolInfo, "hWnd", $hWnd)
	DllStructSetData($tToolInfo, "ID", $iID)
;~ 	DllStructSetData($tToolInfo, "Text", DllStructGetPtr($tBuffer))
;~ 	Local $bResult = _SendMessage($hTool, $TTM_GETTOOLINFOW, 0, $tToolInfo, 0, True)
	Local $bResult = __GUICtrl_SendMsg($hTool, $TTM_GETTOOLINFOW, 0, $tToolInfo, $tBuffer, True, 10, True, -1)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $bResult = ' & $bResult & @CRLF & '>Error code: ' & @error & '    Extended code: ' & @extended & ' (0x' & Hex(@extended) & ')' & @CRLF) ;### Debug Console

	Return _GUIToolTip_ToolToArray($hTool, $tToolInfo, ($bResult = True))
EndFunc   ;==>_GUIToolTip_GetToolInfo

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_HitTest($hTool, $hWnd, $iX, $iY)
	If __GUICtrl_CheckHandleOutProcess($hWnd) Then
		Return SetError(6, 0, "") ; Processes not in same AutoIt Mode
	EndIf

	Local $tHitTest = DllStructCreate($tagTTHITTESTINFO)
	Local $tToolInfo = DllStructCreate($tagTOOLINFO)
	Local $iToolInfo = DllStructGetSize($tToolInfo)
	DllStructSetData($tHitTest, "Tool", $hWnd)
	DllStructSetData($tHitTest, "X", $iX)
	DllStructSetData($tHitTest, "Y", $iY)
	DllStructSetData($tHitTest, "Size", $iToolInfo)
;~ 	Local $bResult = __GUICtrl_SendMsg($hTool, $TTM_HITTESTW, 0, $tHitTest, 0, True)
	Local $bResult = __GUICtrl_SendMsg($hTool, $TTM_HITTESTW, 0, $tHitTest, 0, True, 10, False, -1)

	DllStructSetData($tToolInfo, "Size", DllStructGetData($tHitTest, "Size"))
	DllStructSetData($tToolInfo, "Flags", DllStructGetData($tHitTest, "Flags"))
	DllStructSetData($tToolInfo, "hWnd", DllStructGetData($tHitTest, "hWnd"))
	DllStructSetData($tToolInfo, "ID", DllStructGetData($tHitTest, "ID"))
	DllStructSetData($tToolInfo, "Left", DllStructGetData($tHitTest, "Left"))
	DllStructSetData($tToolInfo, "Top", DllStructGetData($tHitTest, "Top"))
	DllStructSetData($tToolInfo, "Right", DllStructGetData($tHitTest, "Right"))
	DllStructSetData($tToolInfo, "Bottom", DllStructGetData($tHitTest, "Bottom"))
	DllStructSetData($tToolInfo, "hInst", DllStructGetData($tHitTest, "hInst"))
	DllStructSetData($tToolInfo, "Param", DllStructGetData($tHitTest, "Param"))

	Return _GUIToolTip_ToolToArray($hTool, $tToolInfo, $bResult = True)
EndFunc   ;==>_GUIToolTip_HitTest

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_NewToolRect($hTool, $hWnd, $iID, $iLeft, $iTop, $iRight, $iBottom)
	If __GUICtrl_CheckHandleOutProcess($hWnd) Then
		Return SetError(6, 0, "") ; Processes not in same AutoIt Mode
	EndIf

	Local $tToolInfo = DllStructCreate($tagTOOLINFO)
	Local $iToolInfo = DllStructGetSize($tToolInfo)
	DllStructSetData($tToolInfo, "Size", $iToolInfo)
	DllStructSetData($tToolInfo, "hwnd", $hWnd)
	DllStructSetData($tToolInfo, "ID", $iID)
	DllStructSetData($tToolInfo, "Left", $iLeft)
	DllStructSetData($tToolInfo, "Top", $iTop)
	DllStructSetData($tToolInfo, "Right", $iRight)
	DllStructSetData($tToolInfo, "Bottom", $iBottom)

;~ 	__GUICtrl_SendMsg($hTool, $TTM_NEWTOOLRECTW, 0, $tToolInfo)
	__GUICtrl_SendMsg($hTool, $TTM_NEWTOOLRECTW, 0, $tToolInfo, 0, True, 10, False, -1)

EndFunc   ;==>_GUIToolTip_NewToolRect

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_Pop($hTool)
	_SendMessage($hTool, $TTM_POP)
EndFunc   ;==>_GUIToolTip_Pop

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_PopUp($hTool)
	_SendMessage($hTool, $TTM_POPUP)
EndFunc   ;==>_GUIToolTip_PopUp

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_SetDelayTime($hTool, $iDuration, $iTime)
	_SendMessage($hTool, $TTM_SETDELAYTIME, $iDuration, $iTime)
EndFunc   ;==>_GUIToolTip_SetDelayTime

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_SetMargin($hTool, $iLeft, $iTop, $iRight, $iBottom)
	Local $tRECT = DllStructCreate($tagRECT)
	DllStructSetData($tRECT, "Left", $iLeft)
	DllStructSetData($tRECT, "Top", $iTop)
	DllStructSetData($tRECT, "Right", $iRight)
	DllStructSetData($tRECT, "Bottom", $iBottom)

	__GUICtrl_SendMsg($hTool, $TTM_SETMARGIN, 0, $tRECT)

EndFunc   ;==>_GUIToolTip_SetMargin

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_SetMaxTipWidth($hTool, $iWidth)
	Return _SendMessage($hTool, $TTM_SETMAXTIPWIDTH, 0, $iWidth)
EndFunc   ;==>_GUIToolTip_SetMaxTipWidth

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_SetTipBkColor($hTool, $iColor)
	_SendMessage($hTool, $TTM_SETTIPBKCOLOR, $iColor)
EndFunc   ;==>_GUIToolTip_SetTipBkColor

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_SetTipTextColor($hTool, $iColor)
	_SendMessage($hTool, $TTM_SETTIPTEXTCOLOR, $iColor)
EndFunc   ;==>_GUIToolTip_SetTipTextColor

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_SetTitle($hTool, $sTitle, $iIcon = 0)
	Local $tBuffer = $__g_tTTBuffer
	DllStructSetData($tBuffer, "Text", $sTitle)
	Local $iRet = __GUICtrl_SendMsg($hTool, $TTM_SETTITLEW, $iIcon, $tBuffer)

	Return $iRet <> 0
EndFunc   ;==>_GUIToolTip_SetTitle

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_SetToolInfo($hTool, $hWnd, $iID, $sText = "", $iFlags = Default, $iLeft = 0, $iTop = 0, $iRight = 0, $iBottom = 0, $hInst = 0, $iParam = 0)
	If __GUICtrl_CheckHandleOutProcess($hWnd) Then
		Return SetError(6, 0, "") ; Processes not in same AutoIt Mode
	EndIf

	If $iFlags = Default Then $iFlags = BitOR($TTF_SUBCLASS, $TTF_IDISHWND)
	Local $tBuffer = $__g_tTTBuffer
	DllStructSetData($tBuffer, "Text", $sText)

	Local $tToolInfo = DllStructCreate($tagTOOLINFO)
	Local $iToolInfo = DllStructGetSize($tToolInfo)
	DllStructSetData($tToolInfo, "Size", $iToolInfo)
	DllStructSetData($tToolInfo, "Flags", $iFlags)
	DllStructSetData($tToolInfo, "hWnd", $hWnd)
	DllStructSetData($tToolInfo, "ID", $iID)
	DllStructSetData($tToolInfo, "Left", $iLeft)
	DllStructSetData($tToolInfo, "Top", $iTop)
	DllStructSetData($tToolInfo, "Right", $iRight)
	DllStructSetData($tToolInfo, "Bottom", $iBottom)
	DllStructSetData($tToolInfo, "Text", DllStructGetPtr($tBuffer))
	DllStructSetData($tToolInfo, "hInst", $hInst)
	DllStructSetData($tToolInfo, "Param", $iParam)

	__GUICtrl_SendMsg($hTool, $TTM_SETTOOLINFOW, 0, $tToolInfo, $tBuffer, False, 10, False, -1)

EndFunc   ;==>_GUIToolTip_SetToolInfo

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_SetWindowTheme($hTool, $sStyle)
	Local $tBuffer = _WinAPI_MultiByteToWideChar($sStyle)

	__GUICtrl_SendMsg($hTool, $TTM_SETWINDOWTHEME, 0, $tBuffer)

EndFunc   ;==>_GUIToolTip_SetWindowTheme

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_ToolExists($hTool)
	Return _SendMessage($hTool, $TTM_GETCURRENTTOOL) <> 0
EndFunc   ;==>_GUIToolTip_ToolExists

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_ToolToArray($hTool, ByRef $tToolInfo, $iError)
	Local $aTool[10]

	$aTool[0] = DllStructGetData($tToolInfo, "Flags")
	$aTool[1] = DllStructGetData($tToolInfo, "hWnd")
	$aTool[2] = DllStructGetData($tToolInfo, "ID")
	$aTool[3] = DllStructGetData($tToolInfo, "Left")
	$aTool[4] = DllStructGetData($tToolInfo, "Top")
	$aTool[5] = DllStructGetData($tToolInfo, "Right")
	$aTool[6] = DllStructGetData($tToolInfo, "Bottom")
	$aTool[7] = DllStructGetData($tToolInfo, "hInst")
;~ 	$aTool[8] = DllStructGetData($tToolInfo, "Text")
	$aTool[8] = _GUIToolTip_GetText($hTool, $aTool[1], $aTool[2])
	$aTool[9] = DllStructGetData($tToolInfo, "Param")

	Return SetError($iError, 0, $aTool)
EndFunc   ;==>_GUIToolTip_ToolToArray

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_TrackActivate($hTool, $bActivate = True, $hWnd = 0, $iID = 0)
	Local $tToolInfo = DllStructCreate($tagTOOLINFO)
	Local $iToolInfo = DllStructGetSize($tToolInfo)

	DllStructSetData($tToolInfo, "Size", $iToolInfo)
	DllStructSetData($tToolInfo, "hWnd", $hWnd)
	DllStructSetData($tToolInfo, "ID", $iID)

	__GUICtrl_SendMsg($hTool, $TTM_TRACKACTIVATE, $bActivate, $tToolInfo)

EndFunc   ;==>_GUIToolTip_TrackActivate

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_TrackPosition($hTool, $iX, $iY)
	_SendMessage($hTool, $TTM_TRACKPOSITION, 0, _WinAPI_MakeLong($iX, $iY))
EndFunc   ;==>_GUIToolTip_TrackPosition

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_Update($hTool)
	_SendMessage($hTool, $TTM_UPDATE)
EndFunc   ;==>_GUIToolTip_Update

; #FUNCTION# ====================================================================================================================
; Author ........: Paul Campbell (PaulIA)
; Modified.......:
; ===============================================================================================================================
Func _GUIToolTip_UpdateTipText($hTool, $hWnd, $iID, $sText)
	If __GUICtrl_CheckHandleOutProcess($hWnd) Then
		Return SetError(6, 0, "") ; Processes not in same AutoIt Mode
	EndIf

	Local $tBuffer = $__g_tTTBuffer
	DllStructSetData($tBuffer, "Text", $sText)

	Local $tToolInfo = DllStructCreate($tagTOOLINFO)
	Local $iToolInfo = DllStructGetSize($tToolInfo)
	DllStructSetData($tToolInfo, "Size", $iToolInfo)
	DllStructSetData($tToolInfo, "hWnd", $hWnd)
	DllStructSetData($tToolInfo, "ID", $iID)

	__GUICtrl_SendMsg($hTool, $TTM_UPDATETIPTEXTW, 0, $tToolInfo, $tBuffer, False, 10, False, -1)

EndFunc   ;==>_GUIToolTip_UpdateTipText
