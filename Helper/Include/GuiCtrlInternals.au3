#include-once

#include "Memory.au3"               ; _MemFree()

#include "WinAPISysInternals.au3"   ; _WinAPI_GetParent()

; #INDEX# =======================================================================================================================
; Title .........: GUI Ctrl Extended UDF Library for AutoIt3
; AutoIt Version : 3.3.18.0
; Description ...: Functions that assist with _GUI control management.
; Author(s) .....: jpm
; ===============================================================================================================================

#Region Global Variables and Constants

; #CONSTANTS# ===================================================================================================================
Global Const $__GUICTRL_IDS_OFFSET = 2
Global Const $__GUICTRL_ID_MAX_WIN = 16
Global Const $__GUICTRL_STARTID = 10000
Global Const $__GUICTRL_ID_MAX_IDS = 65535 - $__GUICTRL_STARTID

Global Const $__GUICTRLCONSTANT_WS_TABSTOP = 0x00010000
Global Const $__GUICTRLCONSTANT_WS_VISIBLE = 0x10000000
Global Const $__GUICTRLCONSTANT_WS_CHILD = 0x40000000

Global Const $__GUICTRLCONSTANT_WS_EX_CLIENTEDGE = 0x00000200
; ===============================================================================================================================

; #VARIABLES# ===================================================================================================================
Global $__g_hGUICtrl_LastWnd
Global $__g_aGUICtrl_IDs_Used[$__GUICTRL_ID_MAX_WIN][$__GUICTRL_ID_MAX_IDS + $__GUICTRL_IDS_OFFSET + 1] ; [index][0] = HWND, [index][1] = NEXT ID
; ===============================================================================================================================

#EndRegion Global Variables and Constants

#Region Functions list

; #CURRENT# =====================================================================================================================
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
; __GUICtrl_SendMsg
; __GUICtrl_SendMsg_Init
; __GUICtrl_SendMsg_InProcess
; __GUICtrl_SendMsg_OutProcess
; __GUICtrl_SendMsg_Internal
;
; __GUICtrl_TagOutProcess
; __GUICtrl_CheckHandleOutProcess
; __GUICtrl_CheckProcessSameMode
; __GUICtrl_FreeGlobalID
; __GUICtrl_GetNextGlobalID
; __GUICtrl_GetVersion
; __GUICtrl_IsWow64Process
; ===============================================================================================================================

#EndRegion Functions list

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __GUICtrl_SendMsg
; Description ...: _SendMessage() wrapper  for handling In or out process
; Syntax.........: __GUICtrl_SendMsg($hWnd, $iMsg, $iIndex, ByRef $tItem, $tBuffer, $bRetItem, $iElement, $bRetBuffer, $iElementMAX)
; Parameters ....: $hWnd        - Handle of the control
;                  $iMsg        - SendMessage Msg value
;                  $iIndex      - index of the Item
;                  $tItem       - Struct of the Item
;                  $tBuffer     - Struct to contain return chars
;                  $bRetItem    - tItem must be return
;                  $iElement    - index of the element
;                  $bRetBuffer  - tBuffer must be return
;                  $iElementMax - index of the MAX element
; Return values .: result of the SendMessage
; Author ........: Jpm
; Modified ......:
; Remarks .......:
; Related .......: _GUICtrlListView_GetItemEx, _GUICtrlListView_GetItemText, __GUICtrlListView_Sort
; ===============================================================================================================================
Func __GUICtrl_SendMsg($hWnd, $iMsg, $iIndex, ByRef $tItem, $tBuffer = 0, $bRetItem = False, $iElement = -1, $bRetBuffer = False, $iElementMax = $iElement)
	If $iElement > 0 Then
		DllStructSetData($tItem, $iElement, DllStructGetPtr($tBuffer))
		If $iElement = $iElementMax Then DllStructSetData($tItem, $iElement + 1, DllStructGetSize($tBuffer))
	EndIf

	Local $iRet
	If IsHWnd($hWnd) Then
		If ($hWnd = $__g_hGUICtrl_LastWnd) Or (DllCall("user32.dll", "dword", "GetWindowThreadProcessId", "hwnd", $hWnd, "dword*", 0)[2] = @AutoItPID) Then ; _WinAPI_InProcess
			$__g_hGUICtrl_LastWnd = $hWnd
			$iRet = DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $hWnd, "uint", $iMsg, "wparam", $iIndex, "struct*", $tItem)[0]
		Else
			Local $iItem = Ceiling(DllStructGetSize($tItem) / 16) * 16 ; to allocate buffer on a 16 byte boundary
			Local $tMemMap, $pText
			Local $iBuffer = 0
			If ($iElement > 0) Or ($iElementMax = 0) Then $iBuffer = DllStructGetSize($tBuffer)
			Local $pMemory = _MemInit($hWnd, $iItem + $iBuffer, $tMemMap)
			If $iBuffer Then
				$pText = $pMemory + $iItem
				If $iElementMax Then
					DllStructSetData($tItem, $iElement, $pText)
				Else
					$iIndex = $pText
				EndIf
				_MemWrite($tMemMap, $tBuffer, $pText, $iBuffer)
			EndIf
			_MemWrite($tMemMap, $tItem, $pMemory, $iItem)
			$iRet = DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $hWnd, "uint", $iMsg, "wparam", $iIndex, "ptr", $pMemory)[0]
			If $iBuffer And $bRetBuffer Then
				_MemRead($tMemMap, $pText, $tBuffer, $iBuffer)
			EndIf
			If $bRetItem Then _MemRead($tMemMap, $pMemory, $tItem, $iItem)
			_MemFree($tMemMap)
		EndIf
	Else
		$iRet = GUICtrlSendMsg($hWnd, $iMsg, $iIndex, DllStructGetPtr($tItem))
	EndIf

	Return $iRet
EndFunc   ;==>__GUICtrl_SendMsg

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __GUICtrl_SendMsg_Init
; Description ...: Split of __GUICtrl_SendMsg() to be use if looping is required
; Syntax.........: __GUICtrl_SendMsg($hWnd, $iMsg, $iIndex, ByRef $tItem, $tBuffer, $bRetItem, $iElement, $bRetBuffer, $iElementMAX)
; Parameters ....: $hWnd        - Handle of the control
;                  $iMsg        - SendMessage Msg value
;                  $iIndex      - index of the Item
;                  $tItem       - Struct of the Item
;                  $tBuffer     - Struct to contain return chars
;                  $bRetItem    - tItem must be return
;                  $iElement    - index of the element
;                  $bRetBuffer  - tBuffer must be return
;                  $iElementMax - index of the MAX element
; Return values .: ptr to function __GUICtrl_SendMsg_* according to $hWnd type
; Author ........: Jpm
; Modified ......:
; Remarks .......:
; Related .......: __GUICtrl_SendMsg_InProcess, __GUICtrl_SendMsg_OutProcess, __GUICtrl_SendMsg_Internal
; ===============================================================================================================================
Func __GUICtrl_SendMsg_Init($hWnd, $iMsg, $iIndex, ByRef $tItem, $tBuffer = 0, $bRetItem = False, $iElement = -1, $bRetBuffer = False, $iElementMax = $iElement)
	#forceref $iMsg, $iIndex, $bRetItem, $bRetBuffer
	DllStructSetData($tItem, $iElement, DllStructGetPtr($tBuffer))
	If $iElement = $iElementMax Then DllStructSetData($tItem, $iElement + 1, DllStructGetSize($tBuffer))

	Local $pFunc
	If IsHWnd($hWnd) Then
		If ($hWnd = $__g_hGUICtrl_LastWnd) Or DllCall("user32.dll", "dword", "GetWindowThreadProcessId", "hwnd", $hWnd, "dword*", 0)[2] = @AutoItPID Then ; _WinAPI_InProcess
			$__g_hGUICtrl_LastWnd = $hWnd
			$pFunc = __GUICtrl_SendMsg_InProcess
			SetExtended(1)
		Else
			$pFunc = __GUICtrl_SendMsg_OutProcess
			SetExtended(2)
		EndIf
	Else
		$pFunc = __GUICtrl_SendMsg_Internal
		SetExtended(3)
	EndIf

	Return $pFunc
EndFunc   ;==>__GUICtrl_SendMsg_Init

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __GUICtrl_SendMsg_InProcess
; Description ...: Split of __GUICtrl_SendMsg() to be use for InProcess
; Syntax.........: __GUICtrl_SendMsg_InProcess($hWnd, $iMsg, $iIndex, ByRef $tItem, $tBuffer, $bRetItem, $iElement, $bRetBuffer, $iElementMAX)
; Parameters ....: $hWnd        - Handle of the control
;                  $iMsg        - SendMessage Msg value
;                  $iIndex      - index of the Item
;                  $tItem       - Struct of the Item
;                  $tBuffer     - Struct to contain return chars
;                  $bRetItem    - tItem must be return
;                  $iElement    - index of the element
;                  $bRetBuffer  - tBuffer must be return
;                  $iElementMax - index of the MAX element
; Return values .: result of the SendMessage
; Author ........: Jpm
; Modified ......:
; Remarks .......:
; Related .......: __GUICtrl_SendMsg_Init
; ===============================================================================================================================
Func __GUICtrl_SendMsg_InProcess($hWnd, $iMsg, $iIndex, ByRef $tItem, $tBuffer = 0, $bRetItem = False, $iElement = -1, $bRetBuffer = False, $iElementMax = $iElement)
	#forceref $tBuffer, $bRetItem, $bRetBuffer, $iElementMax
	Return DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $hWnd, "uint", $iMsg, "wparam", $iIndex, "struct*", $tItem)[0]
EndFunc   ;==>__GUICtrl_SendMsg_InProcess

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __GUICtrl_SendMsg_OutProcess
; Description ...: Split of __GUICtrl_SendMsg() to be use for OutProcess
; Syntax.........: __GUICtrl_SendMsg_InProcess($hWnd, $iMsg, $iIndex, ByRef $tItem, $tBuffer, $bRetItem, $iElement, $bRetBuffer, $iElementMAX)
; Parameters ....: $hWnd        - Handle of the control
;                  $iMsg        - SendMessage Msg value
;                  $iIndex      - index of the Item
;                  $tItem       - Struct of the Item
;                  $tBuffer     - Struct to contain return chars
;                  $bRetItem    - tItem must be return
;                  $iElement    - index of the element
;                  $bRetBuffer  - tBuffer must be return
;                  $iElementMax - index of the MAX element
; Return values .: result of the SendMessage
; Author ........: Jpm
; Modified ......:
; Remarks .......:
; Related .......: __GUICtrl_SendMsg_Init
; ===============================================================================================================================
Func __GUICtrl_SendMsg_OutProcess($hWnd, $iMsg, $iIndex, ByRef $tItem, $tBuffer = 0, $bRetItem = False, $iElement = -1, $bRetBuffer = False, $iElementMax = $iElement)
	Local $iItem = DllStructGetSize($tItem)
	Local $tMemMap, $pText
	Local $iBuffer = 0
	If ($iElement > 0) Or ($iElementMax = 0) Then $iBuffer = DllStructGetSize($tBuffer)
	Local $pMemory = _MemInit($hWnd, $iItem + $iBuffer, $tMemMap)
	If $iBuffer Then
		$pText = $pMemory + $iItem
		If $iElementMax Then
			DllStructSetData($tItem, $iElement, $pText)
		Else
			$iIndex = $pText
		EndIf
		_MemWrite($tMemMap, $tBuffer, $pText, $iBuffer)
	EndIf
	_MemWrite($tMemMap, $tItem, $pMemory, $iItem)
	Local $iRet = DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $hWnd, "uint", $iMsg, "wparam", $iIndex, "ptr", $pMemory)[0]
	If $iBuffer And $bRetBuffer Then _MemRead($tMemMap, $pText, $tBuffer, $iBuffer)
	If $bRetItem Then _MemRead($tMemMap, $pMemory, $tItem, $iItem)
	_MemFree($tMemMap)

	Return $iRet
EndFunc   ;==>__GUICtrl_SendMsg_OutProcess

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __GUICtrl_SendMsg_Internal
; Description ...: Split of __GUICtrl_SendMsg() to be use for id Control
; Syntax.........: __GUICtrl_SendMsg_Internal($hWnd, $iMsg, $iIndex, ByRef $tItem, $tBuffer, $bRetItem, $iElement, $bRetBuffer, $iElementMAX)
; Parameters ....: $hWnd        - Handle of the control
;                  $iMsg        - SendMessage Msg value
;                  $iIndex      - index of the Item
;                  $tItem       - Struct of the Item
;                  $tBuffer     - Struct to contain return chars
;                  $bRetItem    - tItem must be return
;                  $iElement    - index of the element
;                  $bRetBuffer  - tBuffer must be return
;                  $iElementMax - index of the MAX element
; Return values .: result of the SendMessage
; Author ........: Jpm
; Modified ......:
; Remarks .......:
; Related .......: __GUICtrl_SendMsg_Init
; ===============================================================================================================================
Func __GUICtrl_SendMsg_Internal($hWnd, $iMsg, $iIndex, ByRef $tItem, $tBuffer = 0, $bRetItem = False, $iElement = -1, $bRetBuffer = False, $iElementMax = $iElement)
	#forceref $tBuffer, $bRetItem, $bRetBuffer, $iElementMax
	Return GUICtrlSendMsg($hWnd, $iMsg, $iIndex, DllStructGetPtr($tItem))
EndFunc   ;==>__GUICtrl_SendMsg_Internal

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __GUICtrl_Destroy
; Description ...: Convert ptr when use with an OutProcess control
; Syntax.........: __GUICtrl_Destroy($hWnd, $sClassName)
; Parameters ....: $hWnd        - Handle of the control
;                  $sClassNameb - Claasname of the Control
; Return values .: True/False @error
; Author ........: Jpm
; Modified ......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func __GUICtrl_Destroy($hWnd, $sClassName)
	If Not _WinAPI_IsClassName($hWnd, $sClassName) Then Return SetError(2, 0, False)

	Local $iDestroyed = 0
	If IsHWnd($hWnd) Then
		If DllCall("user32.dll", "dword", "GetWindowThreadProcessId", "hwnd", $hWnd, "dword*", 0)[2] = @AutoItPID Then ; _WinAPI_InProcess()
			Local $nCtrlID = DllCall("user32.dll", "int", "GetDlgCtrlID", "hwnd", $hWnd)[0] ; _WinAPI_GetDlgCtrlID()
			Local $hParent = _WinAPI_GetParent($hWnd)
			$iDestroyed = _WinAPI_DestroyWindow($hWnd)
			Local $iRet = __GUICtrl_FreeGlobalID($hParent, $nCtrlID)
			If Not $iRet Then
				; can check for errors here if needed, for debug
			EndIf
		Else
			; Not Allowed to Destroy Other Applications Control(s)
			Return SetError(1, 0, False)
		EndIf
	Else
		$iDestroyed = GUICtrlDelete($hWnd)
	EndIf
	If $iDestroyed Then $hWnd = 0

	Return $iDestroyed <> 0
EndFunc   ;==>__GUICtrl_Destroy

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __GUICtrl_TagOutProcess
; Description ...: Convert ptr when use with an OutProcess control
; Syntax.........: __GUICtrl_TagOutProcess($hWnd, ByRef $sTag)
; Parameters ....: $hWnd        - Handle of the control
;                  $sTag        - £tag to be updated
; Return values .: None
; Author ........: Jpm
; Modified ......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func __GUICtrl_TagOutProcess($hWnd, ByRef $sTag)
	Local $bIsWow64 = __GUICtrl_IsWow64Process($hWnd)
	;x86 read remote x64
	If Not (@AutoItX64 Or $bIsWow64) Then
		$sTag = StringRegExpReplace($sTag, "(dword_ptr)|(uint_ptr)|(int_ptr)|(ptr)|(lparam)|(wparam)|(hwnd)|(handle)", "UINT64")
	EndIf

	;x64 read remote x86
	If @AutoItX64 And $bIsWow64 Then
		$sTag = StringRegExpReplace($sTag, "(dword_ptr)|(uint_ptr)|(int_ptr)|(ptr)|(lparam)|(wparam)|(hwnd)|(handle)", "UINT")
	EndIf
EndFunc   ;==>__GUICtrl_TagOutProcess

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __GUICtrl_IsWow64Process
; Description ...: Determines whether the specified control has been created by a process running under WOW64
; Syntax.........: __GUICtrl_IsWow64Process($hWnd)
; Parameters ....: $hWnd        - Handle of the control
; Return values .: True : The process is running under WOW64
; Author ........: Jpm
; Modified ......:
; Remarks .......:  wrapper of _WinAPI_IsWow64Process() for specific $hWnd
; Related .......:
; ===============================================================================================================================
Func __GUICtrl_IsWow64Process($hWnd)
	Local $iPID = DllCall("user32.dll", "dword", "GetWindowThreadProcessId", "hwnd", $hWnd, "dword*", 0)[2]
	If Not $iPID Then $iPID = @AutoItPID

	Local $hProcess = DllCall('kernel32.dll', 'handle', 'OpenProcess', 'dword', (__GUICtrl_GetVersion() < 6.0 ? 0x00000400 : 0x00001000), _ ; PROCESS_QUERY_INFORMATION : PROCESS_QUERY_LIMITED_INFORMATION
			'bool', 0, 'dword', $iPID)
	If @error Or Not $hProcess[0] Then Return SetError(@error + 20, @extended, False)

	Local $aCall = DllCall('kernel32.dll', 'bool', 'IsWow64Process', 'handle', $hProcess[0], 'bool*', 0)

	Return $aCall[2]
EndFunc   ;==>__GUICtrl_IsWow64Process

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __GUICtrl_CheckHandleOutProcess
; Description ...: Check Running Mode with an OutProcess
; Syntax.........: __GUICtrl_CheckHandleOutProcess($hWnd)
; Parameters ....: $hWnd  - Handle to the target window
; Return values .: Set @error if different process not having same running mode
; Author ........: Jpm
; Modified ......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func __GUICtrl_CheckHandleOutProcess($hWnd)
	If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
	Local $aCall = DllCall("user32.dll", "dword", "GetWindowThreadProcessId", "hwnd", $hWnd, "dword*", 0)
	Local $iPID = $aCall[2]
	If $iPID <> @AutoItPID Then ; Not _WinAPI_InProcess
		If Not __GUICtrl_CheckProcessSameMode($iPID) Then
			Return SetError(3, 0, 1) ; Not same AutoIt Mode
		EndIf
	Else
		If $hWnd = 0 Then Return 0
	EndIf
	Return 0
EndFunc   ;==>__GUICtrl_CheckHandleOutProcess

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __GUICtrl_CheckProcessSameMode
; Description ...: Check Running Mode with an OutProcess
; Syntax.........: __GUICtrl_CheckProcessSameMode($iPid)
; Parameters ....: $iPid  - Process identification
; Return values .: Set @error if different process not having same running mode
; Author ........: Jpm
; Modified ......:
; Remarks .......:
; Related .......:
; ===============================================================================================================================
Func __GUICtrl_CheckProcessSameMode($iPID)
	Local $iCurMode = 0  ; $SCS_32BIT_BINARY
	If @AutoItX64 Then $iCurMode = 6 ; $SCS_64BIT_BINARY

;~ 	Local $sFilePath = _WinAPI_GetProcessFileName($iPID)
	Local $hProcess = DllCall('kernel32.dll', 'handle', 'OpenProcess', 'dword', ((__GUICtrl_GetVersion() < 6.0) ? 0x00000410 : 0x00001010), _  ; PROCESS_QUERY_INFORMATION : PROCESS_QUERY_LIMITED_INFORMATION + PROCESS_VM_READ
			'bool', 0, 'dword', $iPID)
	Local $sFilePath = DllCall(@SystemDir & '\psapi.dll', 'dword', 'GetModuleFileNameExW', 'handle', $hProcess[0], 'handle', 0, _
			'wstr', '', 'int', 4096)[3]
	DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hProcess[0])

;~ 	Return True
;~ 	_WinAPI_GetBinaryType($sFilePath)
	Local $aCall = DllCall('kernel32.dll', 'int', 'GetBinaryTypeW', 'wstr', $sFilePath, 'dword*', 0)

;~ 	Local $bIsSameMode = (@extended = $iCurMode)
	Return ($aCall[2] = $iCurMode)
EndFunc   ;==>__GUICtrl_CheckProcessSameMode

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __GUICtrl_GetVersion
; Description ...: Retrieves version of the current operating system
; Syntax.........: __GUICtrl_GetVersion()
; Parameters ....: none
; Return values .: The string containing the current OS version
; Author ........: Jpm
; Modified ......:
; Remarks .......: duplicate _WinAPI_GetVersion()
; Related .......:
; ===============================================================================================================================
Func __GUICtrl_GetVersion()
	Local Static $tagOSVERSIONINFO = 'struct;dword OSVersionInfoSize;dword MajorVersion;dword MinorVersion;dword BuildNumber;dword PlatformId;wchar CSDVersion[128];endstruct'
	Local $tOSVI = DllStructCreate($tagOSVERSIONINFO)
	DllStructSetData($tOSVI, 1, DllStructGetSize($tOSVI))

	DllCall('kernel32.dll', 'bool', 'GetVersionExW', 'struct*', $tOSVI)

	Return Number(DllStructGetData($tOSVI, 2) & "." & DllStructGetData($tOSVI, 3))
EndFunc   ;==>__GUICtrl_GetVersion

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __GUICtrl_GetNextGlobalID
; Description ...: Used for setting controlID to GUI controls UDF
; Syntax.........: __GUICtrl_GetNextGlobalID($hWnd)
; Parameters ....: $hWnd      - handle to Main Window
; Return values .: Success - Control ID
;                  Failure - 0 and @error is set, @extended may be set
; Author ........: Gary Frost
; Modified.......:
; Remarks .......: For Internal Use Only
; Related .......:
; ===============================================================================================================================
Func __GUICtrl_GetNextGlobalID($hWnd)
	If DllCall("user32.dll", "dword", "GetWindowThreadProcessId", "hwnd", $hWnd, "dword*", 0)[2] <> @AutoItPID Then ; Not _WinAPI_InProcess
		Return SetError(4, 0, 0)
	EndIf

	Local $nCtrlID, $iUsedIndex = -1, $bAllUsed = True

	; check if window still exists
	If Not WinExists($hWnd) Then Return SetError(2, -1, 0)

	; check that all slots still hold valid window handles
	For $iIndex = 0 To $__GUICTRL_ID_MAX_WIN - 1
		If $__g_aGUICtrl_IDs_Used[$iIndex][0] <> 0 Then
			; window no longer exist, free up the slot and reset the control id counter
			If Not WinExists($__g_aGUICtrl_IDs_Used[$iIndex][0]) Then
				For $x = 0 To UBound($__g_aGUICtrl_IDs_Used, 2) - 1 ; $UBOUND_COLUMNS = 2
					$__g_aGUICtrl_IDs_Used[$iIndex][$x] = 0
				Next
				$__g_aGUICtrl_IDs_Used[$iIndex][1] = $__GUICTRL_STARTID
				$bAllUsed = False
			EndIf
		EndIf
	Next

	; check if window has been used before with this function
	For $iIndex = 0 To $__GUICTRL_ID_MAX_WIN - 1
		If $__g_aGUICtrl_IDs_Used[$iIndex][0] = $hWnd Then
			$iUsedIndex = $iIndex
			ExitLoop ; $hWnd has been used before
		EndIf
	Next

	; window hasn't been used before, get 1st un-used index
	If $iUsedIndex = -1 Then
		For $iIndex = 0 To $__GUICTRL_ID_MAX_WIN - 1
			If $__g_aGUICtrl_IDs_Used[$iIndex][0] = 0 Then
				$__g_aGUICtrl_IDs_Used[$iIndex][0] = $hWnd
				$__g_aGUICtrl_IDs_Used[$iIndex][1] = $__GUICTRL_STARTID
				$bAllUsed = False
				$iUsedIndex = $iIndex
				ExitLoop
			EndIf
		Next
	EndIf

	If $iUsedIndex = -1 And $bAllUsed Then Return SetError(16, 0, 0) ; used up all 16 window slots

	; used all control ids
	If $__g_aGUICtrl_IDs_Used[$iUsedIndex][1] = ($__GUICTRL_STARTID + $__GUICTRL_ID_MAX_IDS) Then
		; check if control has been deleted, if so use that index in array
		For $iIDIndex = $__GUICTRL_IDS_OFFSET To UBound($__g_aGUICtrl_IDs_Used, 2) - 1 ; $UBOUND_COLUMNS = 2
			If $__g_aGUICtrl_IDs_Used[$iUsedIndex][$iIDIndex] = 0 Then
				$nCtrlID = ($iIDIndex - $__GUICTRL_IDS_OFFSET) + $__GUICTRL_STARTID
				$__g_aGUICtrl_IDs_Used[$iUsedIndex][$iIDIndex] = $nCtrlID
				Return $nCtrlID
			EndIf
		Next
		Return SetError(8, $__GUICTRL_ID_MAX_IDS, 0) ; we have used up all available control ids
	EndIf

	; new control id
	$nCtrlID = $__g_aGUICtrl_IDs_Used[$iUsedIndex][1]
	$__g_aGUICtrl_IDs_Used[$iUsedIndex][1] += 1
	$__g_aGUICtrl_IDs_Used[$iUsedIndex][($nCtrlID - $__GUICTRL_STARTID) + $__GUICTRL_IDS_OFFSET] = $nCtrlID
	Return $nCtrlID
EndFunc   ;==>__GUICtrl_GetNextGlobalID

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __GUICtrl_FreeGlobalID
; Description ...: Used for freeing controlID used for GUI controls UDF
; Syntax.........: __GUICtrl_FreeGlobalID($hWnd, $iGlobalID)
; Parameters ....: $hWnd      - handle to Main Window
;                  $iGlobalID - Control ID to free up for re-use if needed
; Return values .: None
; Author ........: Gary Frost
; Modified.......:
; Remarks .......: For Internal Use Only
; Related .......:
; ===============================================================================================================================
Func __GUICtrl_FreeGlobalID($hWnd, $iGlobalID)
	; invalid udf global id passed in
	If ($iGlobalID - $__GUICTRL_STARTID) < 0 Or (($iGlobalID - $__GUICTRL_STARTID) > $__GUICTRL_ID_MAX_IDS) Then Return SetError(-1, 0, False)

	For $iIndex = 0 To $__GUICTRL_ID_MAX_WIN - 1
		If $__g_aGUICtrl_IDs_Used[$iIndex][0] = $hWnd Then
			For $x = $__GUICTRL_IDS_OFFSET To UBound($__g_aGUICtrl_IDs_Used, 2) - 1 ; $UBOUND_COLUMNS = 2
				If $__g_aGUICtrl_IDs_Used[$iIndex][$x] = $iGlobalID Then
					; free up control id
					$__g_aGUICtrl_IDs_Used[$iIndex][$x] = 0
					Return True
				EndIf
			Next
			; $iGlobalID wasn't found in the used list
			Return SetError(-3, 0, False)
		EndIf
	Next
	; $hWnd wasn't found in the used list
	Return SetError(-2, 0, False)
EndFunc   ;==>__GUICtrl_FreeGlobalID
