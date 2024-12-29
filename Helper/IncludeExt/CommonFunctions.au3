#include-once
;===============================================================================
; Common functions used in my scripts
; https://github.com/jmclaren7/autoit-scripts/blob/master/CommonFunctions.au3
;===============================================================================
; If these files have already been included using a custom path, you may need to remove them here
#include <Array.au3>
#include <AutoItConstants.au3>
#include <Date.au3>
#include <File.au3>
#include <GUIConstantsEx.au3>
#include <GuiEdit.au3>
#include <GuiListBox.au3>
#include <String.au3>
#include <WinAPIFiles.au3>
#include <WinAPIProc.au3>
#include <WinAPIShPath.au3>
#include <WinAPISysWin.au3>
#include <WindowsConstants.au3>
;===============================================================================

; #FUNCTION# ====================================================================================================================
; Name ..........: _RandomString
; Description ...: Generates a random string of numbers and characters
; Syntax ........: _RandomString($Min, $Max[, $Chars = Default])
; Parameters ....: $Min                 - Minimum string length.
;                  $Max                 - Maximum string length.
;                  $Chars               - [optional] String of characters to choose from. Default is all numbers and letters (upper and lower).
; Return values .: A random string string
; Author ........: JohnMC - JohnsCS.com
; Modified ......: 05/11/2024
; ===============================================================================================================================
Func _RandomString($iMin, $iMax, $sChars = Default)
	If $sChars = Default Then $sChars = '0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'
	Local $sReturn, $aChars = StringSplit($sChars, '')

	For $i = 1 To Random($iMin, $iMax, 1)
		$sReturn &= $aChars[Random(1, $aChars[0], 1)]
	Next

	Return $sReturn
EndFunc   ;==>_RandomString

; #FUNCTION# ====================================================================================================================
; Name ..........: _ListSelect
; Description ...: Creates a list select box
; Syntax ........: _ListSelect()
; Parameters ....:
; Return values .: The text of the selected item, @extended contains the index of the slected item
; Author ........: JohnMC - JohnsCS.com
; Modified ......: 04/24/2024
; ===============================================================================================================================
Func _ListSelect($aList, $sTitle = "", $sMessage = "", $iDefaultIndex = 1, $sIcon = "", $iWidth = 400, $iHieght = 175)
	Local $IconFile = StringLeft($sIcon, StringInStr($sIcon, ",", 0, -1) - 1)
	Local $IconID = Number(StringTrimLeft($sIcon, StringInStr($sIcon, ",", 0, -1)))

	Local $ListSelectGUI = GUICreate($sTitle, $iWidth, $iHieght, -1, -1)
	Local $ListSelectList1 = GUICtrlCreateList("", 70, 48, $iWidth - 90, $iHieght - 90, -1, 0)
	Local $ListSelectIcon1 = GUICtrlCreateIcon($IconFile, $IconID, 20, 64, 32, 32)
	Local $ListSelectLabel1 = GUICtrlCreateLabel($sMessage, 20, 10, $iWidth - 50, 30)
	Local $ListSelectCancel = GUICtrlCreateButton("Cancel", 304, 140, 75, 25)
	Local $ListSelectOK = GUICtrlCreateButton("OK", 216, 140, 75, 25)
	GUISetIcon($IconFile, $IconID)
	GUISetState(@SW_SHOW)

	For $i = UBound($aList) - 1 To 1 Step -1
		GUICtrlSetData($ListSelectList1, $aList[$i])
	Next
	_GUICtrlListBox_SelectString($ListSelectList1, $aList[$iDefaultIndex])

	Local $ListSelectGUIMsg
	While 1
		$ListSelectGUIMsg = GUIGetMsg()
		Switch $ListSelectGUIMsg
			Case $GUI_EVENT_CLOSE, $ListSelectCancel
				GUIDelete($ListSelectGUI)
				Return SetError(1, 0, 0)

			Case $ListSelectOK
				Local $ReturnIndex = _GUICtrlListBox_GetCurSel($ListSelectList1)
				Local $ReturnText = _GUICtrlListBox_GetText($ListSelectList1, $ReturnIndex)
				GUIDelete($ListSelectGUI)

				Return SetError(0, $ReturnIndex, $ReturnText)
		EndSwitch
		Sleep(10)
	WEnd
EndFunc   ;==>_ListSelect

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

; #FUNCTION# ====================================================================================================================
; Name ..........: _GetVisibleWindows
; Description ...: Get an array of information for visible windows
; Syntax ........: _GetVisibleWindows([$Options = 1])
; Parameters ....: $GetText             - [optional] True/False - Get window text (slow)
; Return values .: 2D Array of windows and window information
;				   [0][0] contains the number of windows
; Author ........: JohnMC - JohnsCS.com, based on AdamUL's _GetVisibleWindows
; Modified ......: 03/29/2024  --  v2.1
; ===============================================================================================================================
Func _GetVisibleWindows($GetText = False)
	Local $NewCol, $TotalTime, $Timer

	Local $TotalTime = TimerInit()

	; 0,1 Retrieve a list of windows
	Local $aWinList = WinList()
	If Not IsArray($aWinList) Then Return SetError(0, 0, 0)
	$aWinList[0][1] = "WinHandle"

	; Resize Array
	ReDim $aWinList[UBound($aWinList) + 1][14 + 1]

	; 2 Add window states
	$NewCol = 2
	$aWinList[0][$NewCol] = "State"
	For $i = 1 To $aWinList[0][0]
		$aWinList[$i][$NewCol] = WinGetState($aWinList[$i][1])
	Next

	; 3 Add state descriptions
	Local $WindowStateIndex = 2
	$NewCol = 3
	$aWinList[0][$NewCol] = "StateDesc"
	For $i = 1 To $aWinList[0][0]
		Local $Desc = ""
		If BitAND($aWinList[$i][$WindowStateIndex], $WIN_STATE_EXISTS) Then $Desc &= "Exists"
		If BitAND($aWinList[$i][$WindowStateIndex], $WIN_STATE_VISIBLE) Then $Desc &= "Visible"
		If BitAND($aWinList[$i][$WindowStateIndex], $WIN_STATE_ENABLED) Then $Desc &= "Enabled"
		If BitAND($aWinList[$i][$WindowStateIndex], $WIN_STATE_ACTIVE) Then $Desc &= "Active"
		If BitAND($aWinList[$i][$WindowStateIndex], $WIN_STATE_MINIMIZED) Then $Desc &= "Minimized"
		If BitAND($aWinList[$i][$WindowStateIndex], $WIN_STATE_MAXIMIZED) Then $Desc &= "Maximized"

		$aWinList[$i][$NewCol] = $Desc
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

	; 4 Get Process ID (PID) and add to the array.
	$NewCol = 4
	$aWinList[0][$NewCol] = "PID"
	For $i = 1 To $aWinList[0][0]
		$aWinList[$i][$NewCol] = WinGetProcess($aWinList[$i][1])
	Next

	; 5 Add process name
	Local $PIDIndex = 4
	$NewCol = 5
	$aWinList[0][$NewCol] = "Name"
	Local $aProcessList = ProcessList()
	For $i = 1 To $aWinList[0][0]
		For $b = 1 To UBound($aProcessList) - 1
			If $aProcessList[$b][1] = $aWinList[$i][$PIDIndex] Then
				$aWinList[$i][$NewCol] = $aProcessList[$b][0]
			EndIf
		Next
	Next

	; 6 Add path using winapi method
	$NewCol = 6
	$aWinList[0][$NewCol] = "Path"
	For $i = 1 To $aWinList[0][0]
		Local $Path = _WinAPI_GetProcessFileName($aWinList[$i][4])
		; No path might mean the process is elevated so let's try some other things...
		If $Path = "" Then
			; Might only help if the stars align
			Local $aEnum = _WinAPI_EnumProcessModules($aWinList[$i][4])
			If Not @error And UBound($aEnum) >= 2 Then
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

	; 7 Add command line string
	$NewCol = 7
	$aWinList[0][$NewCol] = "Command"
	For $i = 1 To $aWinList[0][0]
		$aWinList[$i][$NewCol] = _WinAPI_GetProcessCommandLine($aWinList[$i][4])
	Next

	; 8 Add window position and size
	;   -3200,-3200 is minimized window
	;   -8,-8 is maximized window on 1st display, and x,-8 is maximized window on the nth display were x is the nth display width plus -8 (W + -8).
	$NewCol = 8
	$aWinList[0][$NewCol] = "Position"
	For $i = 1 To $aWinList[0][0]
		Local $aWinPosSize = WinGetPos($aWinList[$i][1])
		If Not @error Then
			$aWinList[$i][$NewCol] = $aWinPosSize[0] & "," & $aWinPosSize[1] & "," & $aWinPosSize[2] & "," & $aWinPosSize[3]
		EndIf
	Next

	; 9,10 Add window style and exstyle
	$NewCol = 9
	$aWinList[0][$NewCol] = "Style"
	$NewCol2 = 10
	$aWinList[0][$NewCol2] = "ExStyle"

	For $i = 1 To $aWinList[0][0]
		Local $tWINDOWINFO = _WinAPI_GetWindowInfo($aWinList[$i][1])
		$aWinList[$i][$NewCol] = DllStructGetData($tWINDOWINFO, 'Style', 1)
		$aWinList[$i][$NewCol2] = DllStructGetData($tWINDOWINFO, 'ExStyle', 1)
	Next

	; 11 Add style descriptions
	Local $StyleIndex = 9
	$NewCol = 11
	$aWinList[0][$NewCol] = "StyleDesc"
	For $i = 1 To $aWinList[0][0]
		Local $Desc = ""
		If BitAND($aWinList[$i][$StyleIndex], $WS_BORDER) Then $Desc &= "Border"
		If BitAND($aWinList[$i][$StyleIndex], $WS_POPUP) Then $Desc &= "Popup"
		If BitAND($aWinList[$i][$StyleIndex], $WS_SYSMENU) Then $Desc &= "Sysmenu"
		If BitAND($aWinList[$i][$StyleIndex], $WS_GROUP) Then $Desc &= "Group"
		If BitAND($aWinList[$i][$StyleIndex], $WS_SIZEBOX) Then $Desc &= "Sizebox"
		If BitAND($aWinList[$i][$StyleIndex], $WS_CHILD) Then $Desc &= "Child"
		If BitAND($aWinList[$i][$StyleIndex], $WS_DLGFRAME) Then $Desc &= "Dialog"
		If BitAND($aWinList[$i][$StyleIndex], $WS_MINIMIZEBOX) Then $Desc &= "Minbox"
		If BitAND($aWinList[$i][$StyleIndex], $WS_MAXIMIZEBOX) Then $Desc &= "Maxbox"
		If BitAND($aWinList[$i][$StyleIndex], $WS_CAPTION) Then $Desc &= "Caption"
		If BitAND($aWinList[$i][$StyleIndex], $WS_CLIPCHILDREN) Then $Desc &= "Clipchildren"
		If BitAND($aWinList[$i][$StyleIndex], $WS_CLIPSIBLINGS) Then $Desc &= "Clipsiblings"
		If BitAND($aWinList[$i][$StyleIndex], $WS_DISABLED) Then $Desc &= "Disabled"
		If BitAND($aWinList[$i][$StyleIndex], $WS_HSCROLL) Then $Desc &= "Hscroll"
		If BitAND($aWinList[$i][$StyleIndex], $WS_MAXIMIZE) Then $Desc &= "Maximize"
		If BitAND($aWinList[$i][$StyleIndex], $WS_OVERLAPPEDWINDOW) Then $Desc &= "Overlappedwindow"
		If BitAND($aWinList[$i][$StyleIndex], $WS_POPUPWINDOW) Then $Desc &= "Popupwindow"
		If BitAND($aWinList[$i][$StyleIndex], $DS_MODALFRAME) Then $Desc &= "Modalframe"
		If BitAND($aWinList[$i][$StyleIndex], $DS_SETFOREGROUND) Then $Desc &= "Setforeground"
		If BitAND($aWinList[$i][$StyleIndex], $WS_TABSTOP) Then $Desc &= "Tabstop"
		$aWinList[$i][$NewCol] = $Desc
	Next

	; 12 Add ExStyle descriptions
	Local $ExStyleIndex = 10
	$NewCol = 12
	$aWinList[0][$NewCol] = "ExStyleDesc"
	For $i = 1 To $aWinList[0][0]
		Local $Desc = ""
		If BitAND($aWinList[$i][$ExStyleIndex], $WS_EX_TOOLWINDOW) Then $Desc &= "Tool"
		If BitAND($aWinList[$i][$ExStyleIndex], $WS_EX_TOPMOST) Then $Desc &= "Top"
		If BitAND($aWinList[$i][$ExStyleIndex], $WS_EX_CONTROLPARENT) Then $Desc &= "Parent"
		If BitAND($aWinList[$i][$ExStyleIndex], $WS_EX_APPWINDOW) Then $Desc &= "App"
		If BitAND($aWinList[$i][$ExStyleIndex], $WS_EX_MDICHILD) Then $Desc &= "Child"
		If BitAND($aWinList[$i][$ExStyleIndex], $WS_EX_DLGMODALFRAME) Then $Desc &= "Dialogframe"
		If BitAND($aWinList[$i][$ExStyleIndex], $WS_EX_ACCEPTFILES) Then $Desc &= "Acceptfiles"
		If BitAND($aWinList[$i][$ExStyleIndex], $WS_EX_CLIENTEDGE) Then $Desc &= "Clientedge"
		If BitAND($aWinList[$i][$ExStyleIndex], $WS_EX_COMPOSITED) Then $Desc &= "Composited"
		If BitAND($aWinList[$i][$ExStyleIndex], $WS_EX_LAYERED) Then $Desc &= "Layered"
		If BitAND($aWinList[$i][$ExStyleIndex], $WS_EX_WINDOWEDGE) Then $Desc &= "Windowedge"
		If BitAND($aWinList[$i][$ExStyleIndex], $WS_EX_STATICEDGE) Then $Desc &= "Staticedge"
		If BitAND($aWinList[$i][$ExStyleIndex], $WS_EX_PALETTEWINDOW) Then $Desc &= "Palettewindow"
		If BitAND($aWinList[$i][$ExStyleIndex], $WS_EX_NOACTIVATE) Then $Desc &= "Noactivate"
		$aWinList[$i][$NewCol] = $Desc
	Next

	; 13 Get Arch
	$NewCol = 13
	$aWinList[0][$NewCol] = "Arch"
	For $i = 1 To $aWinList[0][0]
		Local $sArch
		If _WinAPI_GetBinaryType($aWinList[$i][6]) = 1 Then
			Switch @extended
				Case $SCS_32BIT_BINARY
					$sArch = "32-bit"
				Case $SCS_64BIT_BINARY
					$sArch = "64-bit"
				Case $SCS_DOS_BINARY
					$sArch = "DOS"
				Case $SCS_WOW_BINARY
					$sArch = "16-bit"
			EndSwitch
		EndIf
		$aWinList[$i][$NewCol] = $sArch
	Next

	; 14 Get Window's text and add to the array.
	$NewCol = 14
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

; #FUNCTION# ====================================================================================================================
; Name ..........: IsActivated
; Description ...: Check Windows activation status
; Syntax ........: IsActivated()
; Parameters ....: None
; Return values .: A string with the activation status
;					@error will be true if activation status could not be retrieved
;					@extended will be 0 for activated and >0 for not activated
; Author ........: AutoIT Forum, modified by JohnMC - JohnsCS.com
; Date/Version ..: 03/02/2024  --  v1.1
; ===============================================================================================================================
Func IsActivated()
	$oWMIService = ObjGet("winmgmts:\\.\root\cimv2")
	If Not IsObj($oWMIService) Then Return SetError(1, 0, "WMI Object Error")

	$oCollection = $oWMIService.ExecQuery("SELECT Description, LicenseStatus, GracePeriodRemaining FROM SoftwareLicensingProduct WHERE PartialProductKey <> null")
	If Not IsObj($oCollection) Then Return SetError(2, 0, "WMI Query Error")

	For $oItem In $oCollection
		Switch $oItem.LicenseStatus
			Case 0
				Return SetError(0, 1, "Unlicensed")

			Case 1
				If $oItem.GracePeriodRemaining Then
					If StringInStr($oItem.Description, "TIMEBASED_") Then
						Return SetError(0, 0, "Timebased activation will expire in " & Round($oItem.GracePeriodRemaining / 60 / 24, 1) & " days")

					Else
						Return SetError(0, 0, "Volume activation will expire in " & Round($oItem.GracePeriodRemaining / 60 / 24, 1) & " days")

					EndIf
				Else
					Return SetError(0, 0, "The machine is permanently activated.")

				EndIf

			Case 2
				Return SetError(0, 2, "Initial grace period ends in " & Round($oItem.GracePeriodRemaining / 60 / 24, 1) & " days")

			Case 3
				Return SetError(0, 3, "Additional grace period ends in " & Round($oItem.GracePeriodRemaining / 60 / 24, 1) & " days")

			Case 4
				Return SetError(0, 4, "Non-genuine grace period ends in " & Round($oItem.GracePeriodRemaining / 60 / 24, 1) & " days")

			Case 5
				Return SetError(0, 5, "Windows is in Notification mode")

			Case 6
				Return SetError(0, 6, "Extended grace period ends in " & Round($oItem.GracePeriodRemaining / 60 / 24, 1) & " days")

			Case Else
				Return SetError(4, 7, "Unknown Status Code")

		EndSwitch
	Next

	Return SetError(3, 0, "Unknown Error")
EndFunc   ;==>IsActivated

; #FUNCTION# ====================================================================================================================
; Name ..........: _FileModifiedAge
; Description ...:
; Syntax ........: _FileModifiedAge($sFile)
; Parameters ....: $sFile - Path to a file
; Return values .: File age in miliseconds
; Author ........: AutoIT Forum, modified by JohnMC - JohnsCS.com
; Date/Version ..: 11/15/2023  --  v1.1
; ===============================================================================================================================
Func _FileModifiedAge($sFile)
	$hFile = _WinAPI_CreateFile($sFile, 2, 2)
	If $hFile = 0 Then
		Return SetError(1, 0, -1)
	Else
		$tFileTime = _Date_Time_GetFileTime($hFile)
		$aFileTime = _Date_Time_FileTimeToArray($tFileTime[2])
		$sFileTime = _Date_Time_FileTimeToStr($tFileTime[2], 1)

		_WinAPI_CloseHandle($hFile)

		$tSystemTime = _Date_Time_GetSystemTime()
		$sSystemTime = _Date_Time_SystemTimeToDateTimeStr($tSystemTime, 1)
		$aSystemTime = _Date_Time_SystemTimeToArray($tSystemTime)

		$iFileAge = _DateDiff('s', $sFileTime, $sSystemTime) * 1000 + $aSystemTime[6] - $aFileTime[6]

		Return $iFileAge
	EndIf
EndFunc   ;==>_FileModifiedAge

; #FUNCTION# ====================================================================================================================
; Name ..........: _WMI
; Description ...: Run a WMI query with th eoption of returning the first object (default) or all objects
; Syntax ........: _WMI($Query[, $Single = True])
; Parameters ....: $Query               - an unknown value.
;                  $Single              - [optional] Return the first object, Default is True.
; Return values .: Object(s)
; Author ........: JohnMC - JohnsCS.com
; Date/Version ..: 11/15/2023  --  v1.1
; ===============================================================================================================================
Func _WMI($Query, $Single = True)
	If Not IsDeclared("_objWMIService") Then
		Local $Object = ObjGet("winmgmts:\\.\root\CIMV2")
		If @error Or Not IsObj($Object) Then Return SetError(1, 0, 0)
		Global $_objWMIService = $Object
	EndIf

	Local $colItems = $_objWMIService.ExecQuery($Query)
	If @error Or Not IsObj($colItems) Then Return SetError(2, 0, 0)

	If $Single Then
		Local $objItem

		For $objItem In $colItems
			Return $objItem
		Next

	Else
		Return $colItems

	EndIf

EndFunc   ;==>_WMI

; #FUNCTION# ====================================================================================================================
; Name ..........: _INetSmtpMailCom
; Description ...: Send an email using a Windows API with authentication and encryption which isn't available in the AutoIt UDF _INetSmtpMail
; Syntax ........: _INetSmtpMailCom($sSMTPServer, $sFromName, $sFromAddress, $sToAddress[, $sSubject = ""[, $sBody = ""[,
;                  $sUsername = ""[, $sPassword = ""[, $sCCAddress = ""[, $sBCCAddress = ""[, $iPort = 587[, $bSSL = False[,
;                  $bTLS = True]]]]]]]]])
; Parameters ....: $sSMTPServer         - a string value.
;                  $sFromName           - a string value.
;                  $sFromAddress        - a string value.
;                  $sToAddress          - a string value.
;                  $sSubject            - [optional] a string value. Default is "".
;                  $sBody               - [optional] a string value. Default is "".
;                  $sUsername           - [optional] a string value. Default is "".
;                  $sPassword           - [optional] a string value. Default is "".
;                  $sCCAddress          - [optional] a string value. Default is "".
;                  $sBCCAddress         - [optional] a string value. Default is "".
;                  $iPort               - [optional] an integer value. Default is 587.
;                  $bSSL                - [optional] a boolean value. Default is False.
;                  $bTLS                - [optional] a boolean value. Default is True.
; Return values .: None
; Author ........: AutoIT Forum, modified by JohnMC - JohnsCS.com
; Date/Version ..: 11/15/2023  --  v1.1
; ===============================================================================================================================
Func _INetSmtpMailCom($sSMTPServer, $sFromName, $sFromAddress, $sToAddress, $sSubject = "", $sBody = "", $sUsername = "", $sPassword = "", $sCCAddress = "", $sBCCAddress = "", $iPort = 587, $bSSL = False, $bTLS = True)
	Local $oMail = ObjCreate("CDO.Message")

	$oMail.From = '"' & $sFromName & '" <' & $sFromAddress & '>'
	$oMail.To = $sToAddress
	$oMail.Subject = $sSubject

	If $sCCAddress Then $oMail.Cc = $sCCAddress
	If $sBCCAddress Then $oMail.Bcc = $sBCCAddress

	If StringInStr($sBody, "<") And StringInStr($sBody, ">") Then
		$oMail.HTMLBody = $sBody
	Else
		$oMail.Textbody = $sBody & @CRLF
	EndIf

	$oMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	$oMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $sSMTPServer
	$oMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $iPort

	; Authenticated SMTP
	If $sUsername <> "" Then
		$oMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		$oMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = $sUsername
		$oMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $sPassword
	EndIf

	; Set security parameters
	If $bSSL Then $oMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
	If $bTLS Then $oMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendtls") = True

	; Update settings
	$oMail.Configuration.Fields.Update
	$oMail.Fields.Update

	; Send the Message
	$oMail.Send
	If @error Then Return SetError(2, 0, 0)

	$oMail = ""

EndFunc   ;==>_INetSmtpMailCom

; #FUNCTION# ====================================================================================================================
; Name ..........: _StringExtract
; Description ...:
; Syntax ........: _StringExtract($sString, $sStartSearch, $sEndSearch[, $iStartTrim = 0[, $iEndTrim = 0]])
; Parameters ....: $sString             - a string value.
;                  $sStartSearch        - a string value.
;                  $sEndSearch          - a string value.
;                  $iStartTrim          - [optional] an integer value. Default is 0.
;                  $iEndTrim            - [optional] an integer value. Default is 0.
; Return values .: None
; Author ........: Your Name
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _StringExtract($sString, $sStartSearch, $sEndSearch, $iStartTrim = 0, $iEndTrim = 0)
	$iStartPos = StringInStr($sString, $sStartSearch)
	If Not $iStartPos Then Return SetError(1, 0, 0)
	$iStartPos = $iStartPos + StringLen($sStartSearch)

	$iCount = StringInStr($sString, $sEndSearch, 0, 1, $iStartPos)
	If Not $iCount Then Return SetError(2, 0, 0)
	$iCount = $iCount - $iStartPos

	$sNewString = StringMid($sString, $iStartPos + $iStartTrim, $iCount + $iEndTrim - $iStartTrim)

	Return $sNewString

EndFunc   ;==>_StringExtract

;===============================================================================
;
; Description:      GetUnixTimeStamp - Get time as Unix timestamp value for a specified date
;                   to get the current time stamp call GetUnixTimeStamp with no parameters
;                   beware the current time stamp has system UTC included to get timestamp with UTC + 0
;                   substract your UTC , exemple your UTC is +2 use GetUnixTimeStamp() - 2*3600
; Parameter(s):     Requierd : None
;                   Optional :
;                               - $year => Year ex : 1970 to 2038
;                               - $mon  => Month ex : 1 to 12
;                               - $days => Day ex : 1 to Max Day OF Month
;                               - $hour => Hour ex : 0 to 23
;                               - $min  => Minutes ex : 1 to 60
;                               - $sec  => Seconds ex : 1 to 60
; Return Value(s):  On Success - Returns Unix timestamp
;                   On Failure - No Failure if valid parameters are valid
; Author(s):        azrael-sub7
; User Calltip:     GetUnixTimeStamp() (required: <_AzUnixTime.au3>)
;
;===============================================================================
Func _GetUnixTimeStamp($year = 0, $mon = 0, $days = 0, $hour = 0, $Min = 0, $sec = 0)
	If $year = 0 Then $year = Number(@YEAR)
	If $mon = 0 Then $mon = Number(@MON)
	If $days = 0 Then $days = Number(@MDAY)
	If $hour = 0 Then $hour = Number(@HOUR)
	If $Min = 0 Then $Min = Number(@MIN)
	If $sec = 0 Then $sec = Number(@SEC)
	Local $NormalYears = 0
	Local $LeepYears = 0
	For $i = 1970 To $year - 1 Step +1
		If _BoolLeapYear($i) = True Then
			$LeepYears = $LeepYears + 1
		Else
			$NormalYears = $NormalYears + 1
		EndIf
	Next
	Local $yearNum = (366 * $LeepYears * 24 * 3600) + (365 * $NormalYears * 24 * 3600)
	Local $MonNum = 0
	For $i = 1 To $mon - 1 Step +1
		$MonNum = $MonNum + _LastDayInMonth($year, $i)
	Next
	Return $yearNum + ($MonNum * 24 * 3600) + (($days - 1) * 24 * 3600) + $hour * 3600 + $Min * 60 + $sec
EndFunc   ;==>_GetUnixTimeStamp

;===============================================================================
;
; Description:      UnixTimeStampToTime - Converts UnixTime to Date
; Parameter(s):     Requierd : $UnixTimeStamp => UnixTime ex : 1102141493
;                   Optional : None
; Return Value(s):  On Success - Returns Array
;                               - $Array[0] => Year ex : 1970 to 2038
;                               - $Array[1] => Month ex : 1 to 12
;                               - $Array[2] => Day ex : 1 to Max Day OF Month
;                               - $Array[3] => Hour ex : 0 to 23
;                               - $Array[4] => Minutes ex : 1 to 60
;                               - $Array[5] => Seconds ex : 1 to 60
;                   On Failure  - No Failure if valid parameter is a valid UnixTimeStamp
; Author(s):        azrael-sub7
; User Calltip:     UnixTimeStampToTime() (required: <_AzUnixTime.au3>)
;
;===============================================================================
Func _UnixTimeStampToTime($UnixTimeStamp)
	Dim $pTime[6]
	$pTime[0] = Floor($UnixTimeStamp / 31436000) + 1970 ; pTYear

	Local $pLeap = Floor(($pTime[0] - 1969) / 4)
	Local $pDays = Floor($UnixTimeStamp / 86400)
	$pDays = $pDays - $pLeap
	$pDaysSnEp = Mod($pDays, 365)

	$pTime[1] = 1 ;$pTMon
	$pTime[2] = $pDaysSnEp ;$pTDays

	If $pTime[2] > 59 And _BoolLeapYear($pTime[0]) = True Then $pTime[2] += 1

	While 1
		If ($pTime[2] > 31) Then
			$pTime[2] = $pTime[2] - _LastDayInMonth($pTime[1])
			$pTime[1] = $pTime[1] + 1
		Else
			ExitLoop
		EndIf
	WEnd

	Local $pSec = $UnixTimeStamp - ($pDays + $pLeap) * 86400

	$pTime[3] = Floor($pSec / 3600) ; $pTHour
	$pTime[4] = Floor(($pSec - ($pTime[3] * 3600)) / 60) ;$pTmin
	$pTime[5] = ($pSec - ($pTime[3] * 3600)) - ($pTime[4] * 60) ; $pTSec

	Return $pTime
EndFunc   ;==>_UnixTimeStampToTime

;===============================================================================
;
; Description:      BoolLeapYear - Check if Year is Leap Year
; Parameter(s):     Requierd : $year => Year to check ex : 2011
;                   Optional : None
; Return Value(s):  True if $year is Leap Year else False
; Author(s):        azrael-sub7
; User Calltip:     BoolLeapYear() (required: <_AzUnixTime.au3>)
; Credits :         Wikipedia Leap Year
;===============================================================================
Func _BoolLeapYear($year)
	If Mod($year, 400) = 0 Then
		Return True ;is_leap_year
	ElseIf Mod($year, 100) = 0 Then
		Return False ;is_not_leap_y
	ElseIf Mod($year, 4) = 0 Then
		Return True ;is_leap_year
	Else
		Return False ;is_not_leap_y
	EndIf
EndFunc   ;==>_BoolLeapYear

;===============================================================================
;
; Description:      _LastDayInMonth
;                   if the function is called with no parameters it returns maximum days for current system set month
;                   else it returns maximum days for the specified month in specified year
; Parameter(s):     Requierd : None
;                   Optional :
;                               - $year : year : 1970 to 2038
;                               - $mon : month : 1 to 12
; Return Value(s):
; Author(s):        azrael-sub7
; User Calltip:
;===============================================================================
Func _LastDayInMonth($year = @YEAR, $mon = @MON)
	If Number($mon) = 2 Then
		If _BoolLeapYear($year) = True Then
			Return 29 ;is_leap_year
		Else
			Return 28 ;is_not_leap_y
		EndIf
	Else
		If $mon < 8 Then
			If Mod($mon, 2) = 0 Then
				Return 30
			Else
				Return 31
			EndIf
		Else
			If Mod($mon, 2) = 1 Then
				Return 30
			Else
				Return 31
			EndIf
		EndIf
	EndIf
EndFunc   ;==>_LastDayInMonth

;===============================================================================
; Function Name:    __StringProper
; Description:		Improved version of _StringProper, wont capitalize after apostrophes
; Call With:		__StringProper($s_String)
; Parameter(s):
; Return Value(s):  On Success -
; 					On Failure -
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		10/15/2014  --  v1.0
;===============================================================================
Func __StringProper($s_String)
	Local $iX = 0
	Local $CapNext = 1
	Local $s_nStr = ""
	Local $s_CurChar
	For $iX = 1 To StringLen($s_String)
		$s_CurChar = StringMid($s_String, $iX, 1)
		Select
			Case $CapNext = 1
				If StringRegExp($s_CurChar, '[a-zA-ZÀ-ÿšœžŸ]') Then
					$s_CurChar = StringUpper($s_CurChar)
					$CapNext = 0
				EndIf
			Case Not StringRegExp($s_CurChar, '[a-zA-ZÀ-ÿšœžŸ]') And $s_CurChar <> "'"
				$CapNext = 1
			Case Else
				$s_CurChar = StringLower($s_CurChar)
		EndSelect
		$s_nStr &= $s_CurChar
	Next
	Return $s_nStr
EndFunc   ;==>__StringProper

;===============================================================================
; Function Name:    _FileInUse()
; Description:      Checks if file is in use
; Call With:        _FileInUse($sFilename, $iAccess = 0)
; Parameter(s):     $sFilename = File name
;                   $iAccess = 0 = GENERIC_READ - other apps can have file open in readonly mode
;                   $iAccess = 1 = GENERIC_READ|GENERIC_WRITE - exclusive access to file,
;                   fails if file open in readonly mode by app
; Return Value(s):  1 - file in use (@error contains system error code)
;                   0 - file not in use
;                   -1 dllcall error (@error contains dllcall error code)
; Author(s):        Siao, rover
; Date/Version:		10/15/2014  --  v1.0
;===============================================================================
Func _FileInUse($sFilename, $iAccess = 0)
	Local $aRet, $hFile, $iError, $iDA
	Local Const $GENERIC_WRITE = 0x40000000
	Local Const $GENERIC_READ = 0x80000000
	Local Const $FILE_ATTRIBUTE_NORMAL = 0x80
	Local Const $OPEN_EXISTING = 3
	$iDA = $GENERIC_READ
	If BitAND($iAccess, 1) <> 0 Then $iDA = BitOR($GENERIC_READ, $GENERIC_WRITE)
	$aRet = DllCall("Kernel32.dll", "hwnd", "CreateFile", _
			"str", $sFilename, _ ;lpFileName
			"dword", $iDA, _ ;dwDesiredAccess
			"dword", 0x00000000, _ ;dwShareMode = DO NOT SHARE
			"dword", 0x00000000, _ ;lpSecurityAttributes = NULL
			"dword", $OPEN_EXISTING, _ ;dwCreationDisposition = OPEN_EXISTING
			"dword", $FILE_ATTRIBUTE_NORMAL, _ ;dwFlagsAndAttributes = FILE_ATTRIBUTE_NORMAL
			"hwnd", 0) ;hTemplateFile = NULL
	$iError = @error
	If @error Or IsArray($aRet) = 0 Then Return SetError($iError, 0, -1)
	$hFile = $aRet[0]
	If $hFile = -1 Then ;INVALID_HANDLE_VALUE = -1
		$aRet = DllCall("Kernel32.dll", "int", "GetLastError")
		;ERROR_SHARING_VIOLATION = 32 0x20
		;The process cannot access the file because it is being used by another process.
		If @error Or IsArray($aRet) = 0 Then Return SetError($iError, 0, 1)
		Return SetError($aRet[0], 0, 1)
	Else
		;close file handle
		DllCall("Kernel32.dll", "int", "CloseHandle", "hwnd", $hFile)
		Return SetError(@error, 0, 0)
	EndIf
EndFunc   ;==>_FileInUse

;===============================================================================
; Function Name:    _FileInUseWait
; Description:		Checkes to see if a file has open handles
; Call With:		_FileInUse($sFilePath, $iAccess = 0)
; Parameter(s):
; Return Value(s):  On Success -
; 					On Failure -
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		10/15/2014  --  v1.0
;===============================================================================
Func _FileInUseWait($sFilePath, $Timeout = 0, $Sleep = 2000)
	$Timeout = $Timeout * 1000
	$Time = TimerInit()
	While 1
		If _FileInUse($sFilePath) Then
			_Log("  File locked")
		Else
			Return 1
		EndIf
		If $Timeout > 0 And $Timeout < TimerDiff($Time) Then
			_Log("  Timeout, file locked")
			Return 0
		EndIf
		Sleep($Sleep)
	WEnd
EndFunc   ;==>_FileInUseWait

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
Func _RunWait($sProgram, $Working = "", $Show = Default, $Opt = Default, $Live = False, $Diag = False)
	Local $sData, $iPid

	If $Show = Default Then $Show = @SW_HIDE
	If $Opt = Default Then $Opt = $STDERR_MERGED

	$iPid = Run($sProgram, $Working, $Show, $Opt)
	If @error Then
		_Log("_RunWait: Couldn't Run " & $sProgram)
		Return SetError(1, 0, 0)
	EndIf

	$sData = _ProcessWaitClose($iPid, $Live, $Diag)

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
Func _ProcessWaitClose($iPid, $Live = Default, $Diag = Default)
	Local $sData, $sStdRead

	If $Live = Default Then $Live = False
	If $Diag = Default Then $Diag = False

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
;===============================================================================
; Function Name:    _TreeList()
; Description:
; Call With:		_TreeList()
; Parameter(s):
; Return Value(s):  On Success -
; 					On Failure -
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		06/02/2011  --  v1.0
;===============================================================================
Func _TreeList($Path, $mode = 1)
	Local $FileList_Original = _FileListToArray($Path, "*", 0)
	Local $FileList[1]

	For $i = 1 To UBound($FileList_Original) - 1
		Local $file_path = $Path & "\" & $FileList_Original[$i]
		If StringInStr(FileGetAttrib($file_path), "D") Then
			$new_array = _TreeList($file_path, $mode)
			_ArrayConcatenate($FileList, $new_array, 1)
		Else
			ReDim $FileList[UBound($FileList) + 1]
			$FileList[UBound($FileList) - 1] = $file_path
		EndIf
	Next

	Return $FileList
EndFunc   ;==>_TreeList
;===============================================================================
; Function Name:    _StringStripWS()
; Description:		Strips all white chars, excluing char(32) the regular space
; Call With:		_StringStripWS($String)
; Parameter(s): 	$String - String To Strip
; Return Value(s):  On Success - Striped String
; 					On Failure - Full String
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		01/29/2010  --  v1.0
;===============================================================================
Func _StringStripWS($String)
	Return StringRegExpReplace($String, "[" & Chr(0) & Chr(9) & Chr(10) & Chr(11) & Chr(12) & Chr(13) & "]", "")
EndFunc   ;==>_StringStripWS
;===============================================================================
; Function Name:    _mousecheck()
; Description:		Checks for mouse movement
; Call With:		_mousecheck($Sleep)
; Parameter(s): 	$Sleep - Miliseconds between mouse checks, 0=Compare At Next Call
; Return Value(s):  On Success - 1 (Mouse Moved)
; 					On Failure - 0 (Mouse Didnt Move)
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		01/29/2010  --  v1.0
;===============================================================================
Func _MouseCheck($Sleep = 300)
	Local $MOUSECHECK_POS1, $MOUSECHECK_POS2

	If $Sleep = 0 Then Global $MOUSECHECK_POS1
	If IsArray($MOUSECHECK_POS1) = 0 And $Sleep = 0 Then $MOUSECHECK_POS1 = MouseGetPos()
	Sleep($Sleep)
	$MOUSECHECK_POS2 = MouseGetPos()
	If Abs($MOUSECHECK_POS1[0] - $MOUSECHECK_POS2[0]) > 2 Or Abs($MOUSECHECK_POS1[1] - $MOUSECHECK_POS2[1]) > 2 Then
		If $Sleep = 0 Then $MOUSECHECK_POS1 = $MOUSECHECK_POS2
		Return 1
	EndIf

	Return 0
EndFunc   ;==>_MouseCheck
;===============================================================================
; Function Name:    _KeyValue()
; Description:		Work with 2d arrays treated as key value pairs such as the ones produced by INIReadSection()
; Call With:		_KeyValue(ByRef $Array, $Key[, $Value[, $Extended]])
; Parameter(s): 	$Array - A previously declared array, if not array, it will be made as one
;					$Key - The value to look for in the first column/dimention or the "Key" in an INI section
;		(Optional)	$Value - The value to write to the array
;		(Optional)	$Delete - If True, delete the specified key
;
; Return Value(s):  On Success - The value found or set or true if a value was deleted
; 					On Failure - "" and sets @error to 1
;
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		10/05/2023  --  v1.0
; Notes:            $Array[0][0] Contains the number of stored parameters
; Example:			_KeyValue($Settings, "trayicon", "1")
;===============================================================================
Func _KeyValue(ByRef $aArray, $Key, $Value = Default, $Delete = Default)
	Local $i

	If $Delete = Default Then $Delete = False

	; Make $Array an array if not already
	If Not IsArray($aArray) Then Dim $aArray[1][2]

	; Loop through array to check for existing key
	For $i = 1 To UBound($aArray) - 1
		If $aArray[$i][0] = $Key Then
			; Read existing value
			If $Value = Default Then
				Return $aArray[$i][1]

				; Update existing value
			Else
				$aArray[$i][1] = $Value
				$aArray[0][0] = UBound($aArray) - 1
				Return $Value
			EndIf

			; Delete existing value
			If $Delete Then
				Local $aNewArray[]
				; Loop through array and copy all keys/values not matching the specified key
				For $i = 1 To UBound($aArray) - 1
					; Skip the key to be deleted
					If $aArray[$i][0] = $Key Then ContinueLoop

					; Resize array and add new key/value
					ReDim $aArray[UBound($aNewArray) + 1][2]
					$aArray[UBound($aNewArray)][0] = $aArray[$i][0]
					$aArray[UBound($aNewArray)][1] = $aArray[$i][1]
				Next

				$aNewArray[0][0] = UBound($aArray) - 1

				; Return array with key/value removed
				$aArray = $aNewArray
				Return True
			EndIf
		EndIf
	Next

	; Add new key/value if it's been specified
	If $Value <> Default Then
		ReDim $aArray[UBound($aArray) + 1][2]
		$aArray[UBound($aArray) - 1][0] = $Key
		$aArray[UBound($aArray) - 1][1] = $Value
		$aArray[0][0] = UBound($aArray) - 1

		Return $Value
	EndIf

	; Return error because a key doesn't exist and nothing else to do
	SetError(1)
	Return ""
EndFunc   ;==>_KeyValue

;===============================================================================
; Function Name:   	_Log()
; Description:		Console log, file log, custom GUI log
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
;					7/21/2024 -- Added script line number
;===============================================================================
Func _Log($sMessage, $iLevel = Default, $bOverWriteLast = Default, $iCallingLine = @ScriptLineNumber)
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
	Local $sTime = @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC

	If @Compiled Then
		Local $sLogLine = $sTime & "> " & $sMessage
	Else
		Local $sLogLine = $sTime & " " & $iCallingLine & "> " & $sMessage
	EndIf

	Local $Minimize = False

	; Do not log this message if $iLevel is greater than global $LogLevel
	If $iLevel > $LogLevel Then Return ""

	; Send to console
	Local $ExternalConsoleFunction = "_Console_Write" ; From Console.au3

	If $bOverWriteLast And Not @Compiled Then
		; Do Nothing
	ElseIf $bOverWriteLast Then
		ConsoleWrite(@CR & $sLogLine)
		Call($ExternalConsoleFunction, @CR & $sLogLine)
	Else
		ConsoleWrite(@CRLF & $sLogLine)
		Call($ExternalConsoleFunction, @CRLF & $sLogLine)
	EndIf

	; Append message to custom GUI if $LogTitle is set
	If $LogTitle <> "" Then
		Local $sLogLineGUI = StringReplace($sLogLine, @LF, @CRLF)

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
			_GUICtrlEdit_AppendText($_hLogEdit, $sLogLineGUI)
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
			_GUICtrlEdit_AppendText($_hLogEdit, @CRLF & $sLogLineGUI)
			_GUICtrlEdit_LineScroll($_hLogEdit, -StringLen($sLogLineGUI), _GUICtrlEdit_GetLineCount($_hLogEdit))
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
;===============================================================================
; Function Name:    _TimeStamp()
; Description:		Returns time since 0 unlike the unknown timestamp behavior of timer_init
; Call With:		_TimeStamp([$Flag])
; Parameter(s): 	$Flag - (Optional) Default is 0 (Miliseconds)
;						1 = Return Total Seconds
;						2 = Return Total Minutes
; Return Value(s):  On Success - Time
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		01/29/2010  --  v1.0
;===============================================================================
Func _TimeStamp($Flag = 0)
	Local $Time

	$Time = @YEAR * 365 * 24 * 60 * 60 * 1000
	$Time = $Time + @YDAY * 24 * 60 * 60 * 1000
	$Time = $Time + @HOUR * 60 * 60 * 1000
	$Time = $Time + @MIN * 60 * 1000
	$Time = $Time + @SEC * 1000
	$Time = $Time + @MSEC

	If $Flag = 1 Then Return Int($Time / 1000) ;Return Seconds
	If $Flag = 2 Then Return Int($Time / 1000 / 60) ;Return Minutes
	Return Int($Time) ;Return Miliseconds
EndFunc   ;==>_TimeStamp
;===============================================================================
; Function Name:    _ProcessWaitNew()
; Description:		Wait for a new proccess to be created before continuing
; Call With:		_ProcessWaitNew($proc,$timeout=0)
; Parameter(s): 	$Process - PID or proccess name
;					$Timeout - (Optional) Miliseconds Before Giving Up
; Return Value(s):  On Success - 1
; 					On Failure - 0 (Timeout)
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		01/29/2010  --  v1.0
;===============================================================================
Func _ProcessWaitNew($Process, $Timeout = 0)
	Local $Count1 = _ProcessCount(), $Count2
	Local $StartTime = _TimeStamp()

	While 1
		$Count2 = _ProcessCount()
		If $Count2 > $Count1 Then Return 1
		If $Count2 < $Count1 Then $Count1 = $Count2

		If $Timeout > 0 And $StartTime < _TimeStamp() - $Timeout Then ExitLoop
		Sleep(100)
	WEnd

	Return 0
EndFunc   ;==>_ProcessWaitNew
;===============================================================================
; Function Name:    _ProcessCount()
; Description:		Returns the number of processes with the same name
; Call With:		_ProcessCount([$Process[,$OnlyUser]])
; Parameter(s): 	$Process - PID or process name
;					$OnlyUser - Only evaluate processes from this user
; Return Value(s):  On Success - Count
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		01/29/2010  --  v1.0
;===============================================================================
Func _ProcessCount($Process = @AutoItPID, $OnlyUser = "")
	Local $Count = 0, $Array = ProcessList($Process)

	For $i = 1 To $Array[0][0]
		If $Array[$i][1] = $Process Then
			If $OnlyUser <> "" And $OnlyUser <> _ProcessOwner($Array[$i][1]) Then ContinueLoop
			$Count = $Count + 1
		EndIf
	Next

	Return $Count
EndFunc   ;==>_ProcessCount

;===============================================================================
; Function Name:    _ProcessOwner()
; Description:		Gets username of the owner of a PID
; Call With:		_ProcessOwner($PID[,$Hostname])
; Parameter(s): 	$PID - PID of proccess
;					$Hostname - (Optional) The computers name to check on
; Return Value(s):  On Success - Username of owner
; 					On Failure - 0
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		01/29/2010  --  v1.0
;===============================================================================
Func _ProcessOwner($PID, $Hostname = ".")
	Local $User, $objWMIService, $colProcess, $objProcess

	$objWMIService = ObjGet("winmgmts://" & $Hostname & "/root/cimv2")
	$colProcess = $objWMIService.ExecQuery("Select * from Win32_Process Where ProcessID ='" & $PID & "'")

	For $objProcess In $colProcess
		If $objProcess.ProcessID = $PID Then
			$User = 0
			$objProcess.GetOwner($User)
			Return $User
		EndIf
	Next
EndFunc   ;==>_ProcessOwner

;===============================================================================
; Function Name:    _ProcessCloseOthers()
; Description:		Closes other proccess of the same name
; Call With:		_ProcessCloseOthers([$Process[,$ExcludingUser[,$OnlyUser]]])
; Parameter(s): 	$Process - (Optional) Name or PID
;					$ExcludingUser - (Optional) Username of owner to exclude
;					$OnlyUser - (Optional) Username of proccesses owner to close
; Return Value(s):  On Success - 1
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		01/29/2010  --  v1.0
;===============================================================================
Func _ProcessCloseOthers($Process = @ScriptName, $ExcludingUser = "", $OnlyUser = "")
	Local $Array = ProcessList($Process)

	For $i = 1 To $Array[0][0]
		If ($Array[$i][1] <> @AutoItPID) Then
			If $ExcludingUser <> "" And _ProcessOwner($Array[$i][1]) <> $ExcludingUser Then
				ProcessClose($Array[$i][1])
			ElseIf $OnlyUser <> "" And _ProcessOwner($Array[$i][1]) = $OnlyUser Then
				ProcessClose($Array[$i][1])
			ElseIf $OnlyUser = "" And $ExcludingUser = "" Then
				ProcessClose($Array[$i][1])
			EndIf
		EndIf
	Next
EndFunc   ;==>_ProcessCloseOthers

;===============================================================================
; Function Name:    _OnlyInstance()
; Description:		Checks to see if we are the only instance running
; Call With:		_OnlyInstance($iFlag)
; Parameter(s): 	$iFlag
;						0 = Continue Anyway
;						1 = Exit Without Notification
;						2 = Exit After Notifying
;						3 = Prompt What To Do
;						4 = Close Other Proccesses
; Return Value(s):  On Success - 1 (Found Another Instance)
; 					On Failure - 0 (Didnt Find Another Instance)
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		10/15/2014  --  v1.1
;===============================================================================
Func _OnlyInstance($iFlag)
	Local $ERROR_ALREADY_EXISTS = 183, $Handle, $LastError, $Message

	If @Compiled = 0 Then Return 0

	$Handle = DllCall("kernel32.dll", "int", "CreateMutex", "int", 0, "long", 1, "str", @ScriptName)
	$LastError = DllCall("kernel32.dll", "int", "GetLastError")
	If $LastError[0] = $ERROR_ALREADY_EXISTS Then
		SetError($LastError[0], $LastError[0], 0)
		Switch $iFlag
			Case 0
				Return 1
			Case 1
				ProcessClose(@AutoItPID)
			Case 2
				MsgBox(262144 + 48, @ScriptName, "The Program Is Already Running")
				ProcessClose(@AutoItPID)
			Case 3
				If MsgBox(262144 + 256 + 48 + 4, @ScriptName, "The Program (" & @ScriptName & ") Is Already Running, Continue Anyway?") = 7 Then ProcessClose(@AutoItPID)
			Case 4
				_ProcessCloseOthers()
		EndSwitch
		Return 1
	EndIf
	Return 0
EndFunc   ;==>_OnlyInstance

;===============================================================================
; Function Name:    _MsgBox()
; Description:		Displays a msgbox without haulting script by using /AutoIt3ExecuteLine
; Call With:		_MsgBox($Flag,$Title,$Text,$Timeout=0)
; Parameter(s): 	All the same options as standard message box
; Return Value(s):  On Success - PID of new proccess
; 					On Failure - 0
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		01/29/2010  --  v1.1
;===============================================================================
Func _MsgBox($Flag, $Title, $Text, $Timeout = 0)
	If $Title = "" Then $Title = @ScriptName
	If $Flag = "" Or IsInt($Flag) = 0 Then $Flag = 0
	Return Run('"' & @AutoItExe & '"' & ' /AutoIt3ExecuteLine "msgbox(' & $Flag & ',''' & $Title & ''',''' & $Text & ''',''' & $Timeout & ''')"')
EndFunc   ;==>_MsgBox

;===============================================================================
; Function Name:    _GetDriveFromSerial()
; Description:		Find a drives letter based on the drives serial
; Call With:		_GetDriveFromSerial($Serial)
; Parameter(s): 	$Serial - Serial of the drive
; Return Value(s):  On Success - Drive letter with ":"
; 					On Failure - 0
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		01/29/2010  --  v1.0
;===============================================================================
Func _GetDriveFromSerial($Serial)
	Local $Drivelist
	$Drivelist = StringSplit("c:,d:,e:,f:,g:,h:,i:,j:,k:,l:,m:,n:,o:,p:,q:,r:,s:,t:,u:,v:,w:,x:,y:,z:", ",")
	For $i = 1 To $Drivelist[0]
		If (DriveGetSerial($Drivelist[$i]) = $Serial And DriveStatus($Drivelist[$i]) = "READY") Then Return $Drivelist[$i]
	Next
	Return 0
EndFunc   ;==>_GetDriveFromSerial

;===============================================================================
; Function Name:    _Sleep()
; Description:		Simple modification to sleep to allow for adlib functions to run
; Call With:		_Sleep($iTime)
; Parameter(s): 	$iTime - Time in MS
; Return Value(s):  none
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		10/15/2014  --  v1.0
;===============================================================================
Func _Sleep($iTime)
	$iTime = Round($iTime / 10)
	For $i = 1 To $iTime
		Sleep(10)
	Next
EndFunc   ;==>_Sleep

;===============================================================================
; Function Name:    _IsInternet()
; Description:		Gets internet connection state as determined by Windows
; Call With:		_IsInternet()
; Parameter(s): 	none
; Return Value(s):  Success - 1
;					Failure - 0
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		10/15/2014  --  v1.0
;===============================================================================
Func _IsInternet()
	Local $Ret = DllCall("wininet.dll", 'int', 'InternetGetConnectedState', 'dword*', 0x20, 'dword', 0)

	If (@error) Then
		Return SetError(1, 0, 0)
	EndIf

	Local $wError = _WinAPI_GetLastError()

	Return SetError((Not ($wError = 0)), $wError, $Ret[0])
EndFunc   ;==>_IsInternet

;===============================================================================
; Function Name:    _WinGetClientPos
; Description:		Retrieves the position and size of the client area of given window
; Call With:		_WinGetClientPos($hWin)
; Parameter(s): 	$hWnd - Handle to window
;					$Absolute - 1 = Get coordinates relative to the screen (deafult)
;								0 = Get coordinates relative to the window
; Return Value(s):	On Success - Returns an array containing location values for client area of the specified window
;									$Array[0] = X position
;									$Array[1] = Y position
;									$Array[2] = Width
;									$Array[3] = Height
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		11/21/2012  --  v1.2
;===============================================================================
Func _WinGetClientPos($hWnd, $Absolute = 1)
	Local $aPos[4], $aSize[4], $aWinPos[4]

	Local $tpoint = DllStructCreate("int X;int Y")
	DllStructSetData($tpoint, "X", 0)
	DllStructSetData($tpoint, "Y", 0)
	DllCall("user32.dll", "bool", "ClientToScreen", "hwnd", $hWnd, "struct*", $tpoint)

	$aSize = WinGetClientSize($hWnd)

	$aPos[0] = DllStructGetData($tpoint, "X")
	$aPos[1] = DllStructGetData($tpoint, "Y")
	$aPos[2] = $aSize[0]
	$aPos[3] = $aSize[1]

	If $Absolute = 0 Then
		$aWinPos = WinGetPos($hWnd)
		$aPos[0] = $aPos[0] - $aWinPos[0]
		$aPos[1] = $aPos[1] - $aWinPos[1]
	EndIf

	Return $aPos
EndFunc   ;==>_WinGetClientPos

;===============================================================================
; Function Name:    _WinMoveClient
; Description:		Position and size the client area of given window
; Call With:		_WinMoveClient()
; Parameter(s):
; Return Value(s):	Success - Handle to the window
;					Failure - 0
; Author(s):        JohnMC - JohnsCS.com
; Date/Version:		10/15/2014  --  v1.0
;===============================================================================
Func _WinMoveClient($sTitle, $sText, $X, $Y, $Width = Default, $Height = Default, $Speed = Default)

	Local $WinPos = WinGetPos($sTitle, $sText)
	Local $ClientSize = WinGetClientSize($sTitle, $sText)

	If $Width <> Default Then $Width = $Width + $WinPos[2] - $ClientSize[0]
	If $Height <> Default Then $Height = $Height + $WinPos[3] - $ClientSize[1]

	Return WinMove($sTitle, $sText, $X, $Y, $Width, $Height, $Speed)
EndFunc   ;==>_WinMoveClient

;=============================================================================================
; Name:				 _HighPrecisionSleep()
; Description:		Sleeps down to 0.1 microseconds
; Syntax:			_HighPrecisionSleep( $iMicroSeconds, $hDll=False)
; Parameter(s):		$iMicroSeconds        - Amount of microseconds to sleep
;					$hDll  - Can be supplied so the UDF doesn't have to re-open the dll all the time.
; Return value(s):	None
; Author:			Andreas Karlsson (monoceres)
; Remarks:			Even though this has high precision you need to take into consideration that it will take some time for autoit to call the function.
;=============================================================================================
Func _HighPrecisionSleep($iMicroSeconds, $dll = "")
	Local $hStruct, $bLoaded
	If $dll <> "" Then $HPS_hDll = $dll
	If Not IsDeclared("HPS_hDll") Then
		Global $HPS_hDll
		$HPS_hDll = DllOpen("ntdll.dll")
		$bLoaded = True
	EndIf
	$hStruct = DllStructCreate("int64 time;")
	DllStructSetData($hStruct, "time", -1 * ($iMicroSeconds * 10))
	DllCall($HPS_hDll, "dword", "ZwDelayExecution", "int", 0, "ptr", DllStructGetPtr($hStruct))
EndFunc   ;==>_HighPrecisionSleep

;===============================================================================
; Function:		_ProcessGetWin
; Purpose:		Return information on the Window owned by a process (if any)
; Syntax:		_ProcessGetWin($iPID)
; Parameters:	$iPID = integer process ID
; Returns:  	On success returns an array:
; 					[0] = Window Title (if any)
;					[1] = Window handle
;				If $iPID does not exist, returns empty array and @error = 1
; Notes:		Not every process has a window, indicated by an empty array and
;   			@error = 0, and not every window has a title, so test [1] for the handle
;   			to see if a window existed for the process.
; Author:		PsaltyDS at www.autoitscript.com/forum
;===============================================================================
Func _ProcessGetWin($iPid)
	Local $avWinList = WinList(), $avRET[2]
	For $n = 1 To $avWinList[0][0]
		If WinGetProcess($avWinList[$n][1]) = $iPid Then
			$avRET[0] = $avWinList[$n][0] ; Title
			$avRET[1] = $avWinList[$n][1] ; HWND
			ExitLoop
		EndIf
	Next
	If $avRET[1] = "" Then SetError(1)
	Return $avRET
EndFunc   ;==>_ProcessGetWin

;===============================================================================
; Function Name:    _ProcessListProperties()
; Description:   Get various properties of a process, or all processes
; Call With:       _ProcessListProperties( [$Process [, $sComputer]] )
; Parameter(s):  (optional) $Process - PID or name of a process, default is "" (all)
;          (optional) $sComputer - remote computer to get list from, default is local
; Requirement(s):   AutoIt v3.2.4.9+
; Return Value(s):  On Success - Returns a 2D array of processes, as in ProcessList()
;            with additional columns added:
;            [0][0] - Number of processes listed (can be 0 if no matches found)
;            [1][0] - 1st process name
;            [1][1] - 1st process PID
;            [1][2] - 1st process Parent PID
;            [1][3] - 1st process owner
;            [1][4] - 1st process priority (0 = low, 31 = high)
;            [1][5] - 1st process executable path
;            [1][6] - 1st process CPU usage
;            [1][7] - 1st process memory usage
;            [1][8] - 1st process creation date/time = "MM/DD/YYY hh:mm:ss" (hh = 00 to 23)
;            [1][9] - 1st process command line string
;            ...
;            [n][0] thru [n][9] - last process properties
; On Failure:     	Returns array with [0][0] = 0 and sets @Error to non-zero (see code below)
; Author(s):      	PsaltyDS at http://www.autoitscript.com/forum
; Date/Version:   	12/01/2009  --  v2.0.4
; Notes:        	If an integer PID or string process name is provided and no match is found,
;             		then [0][0] = 0 and @error = 0 (not treated as an error, same as ProcessList)
;          			This function requires admin permissions to the target computer.
;          			All properties come from the Win32_Process class in WMI.
;            		To get time-base properties (CPU and Memory usage), a 100ms SWbemRefresher is used.
;===============================================================================
Func _ProcessListProperties($Process = "", $sComputer = ".")
	Local $sUsername, $sMsg, $sUserDomain, $avProcs, $dtmDate
	Local $avProcs[1][2] = [[0, ""]], $n = 1

	; Convert PID if passed as string
	If StringIsInt($Process) Then $Process = Int($Process)

	; Connect to WMI and get process objects
	$oWMI = ObjGet("winmgmts:{impersonationLevel=impersonate,authenticationLevel=pktPrivacy, (Debug)}!\\" & $sComputer & "\root\cimv2")
	If IsObj($oWMI) Then
		; Get collection processes from Win32_Process
		If $Process == "" Then
			; Get all
			$colProcs = $oWMI.ExecQuery("select * from win32_process")
		ElseIf IsInt($Process) Then
			; Get by PID
			$colProcs = $oWMI.ExecQuery("select * from win32_process where ProcessId = " & $Process)
		Else
			; Get by Name
			$colProcs = $oWMI.ExecQuery("select * from win32_process where Name = '" & $Process & "'")
		EndIf

		If IsObj($colProcs) Then
			; Return for no matches
			If $colProcs.count = 0 Then Return $avProcs

			; Size the array
			ReDim $avProcs[$colProcs.count + 1][10]
			$avProcs[0][0] = UBound($avProcs) - 1

			; For each process...
			For $oProc In $colProcs
				; [n][0] = Process name
				$avProcs[$n][0] = $oProc.name
				; [n][1] = Process PID
				$avProcs[$n][1] = $oProc.ProcessId
				; [n][2] = Parent PID
				$avProcs[$n][2] = $oProc.ParentProcessId
				; [n][3] = Owner
				If $oProc.GetOwner($sUsername, $sUserDomain) = 0 Then $avProcs[$n][3] = $sUserDomain & "\" & $sUsername
				; [n][4] = Priority
				$avProcs[$n][4] = $oProc.Priority
				; [n][5] = Executable path
				$avProcs[$n][5] = $oProc.ExecutablePath
				; [n][8] = Creation date/time
				$dtmDate = $oProc.CreationDate
				If $dtmDate <> "" Then
					; Back referencing RegExp pattern from weaponx
					Local $sRegExpPatt = "\A(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})(?:.*)"
					$dtmDate = StringRegExpReplace($dtmDate, $sRegExpPatt, "$2/$3/$1 $4:$5:$6")
				EndIf
				$avProcs[$n][8] = $dtmDate
				; [n][9] = Command line string
				$avProcs[$n][9] = $oProc.CommandLine

				; increment index
				$n += 1
			Next
		Else
			SetError(2) ; Error getting process collection from WMI
		EndIf
		; release the collection object
		$colProcs = 0

		; Get collection of all processes from Win32_PerfFormattedData_PerfProc_Process
		; Have to use an SWbemRefresher to pull the collection, or all Perf data will be zeros
		Local $oRefresher = ObjCreate("WbemScripting.SWbemRefresher")
		$colProcs = $oRefresher.AddEnum($oWMI, "Win32_PerfFormattedData_PerfProc_Process").objectSet
		$oRefresher.Refresh

		; Time delay before calling refresher
		Local $iTime = TimerInit()
		Do
			Sleep(20)
		Until TimerDiff($iTime) >= 100
		$oRefresher.Refresh

		; Get PerfProc data
		For $oProc In $colProcs
			; Find it in the array
			For $n = 1 To $avProcs[0][0]
				If $avProcs[$n][1] = $oProc.IDProcess Then
					; [n][6] = CPU usage
					$avProcs[$n][6] = $oProc.PercentProcessorTime
					; [n][7] = memory usage
					$avProcs[$n][7] = $oProc.WorkingSet
					ExitLoop
				EndIf
			Next
		Next
	Else
		SetError(1) ; Error connecting to WMI
	EndIf

	; Return array
	Return $avProcs
EndFunc   ;==>_ProcessListProperties

;===============================================================================
; Function:		_IsIP
; Purpose:		Validate if string is an IP address
; Syntax:		_IsIP($sIP)
; Parameters:	$sIP = String to validate as IP address
; Returns:  	Success - 1=IP 2=Subnet
;				Failure - 0 ()
; Notes:
; Author:
; Date/Version:   	10/15/2014  --  v2.0.4
;===============================================================================
Func _IsIP($sIP, $P_strict = 0)
	$t_ip = $sIP
	$port = StringInStr($t_ip, ":") ;check for : (for the port)
	If $port Then ;has a port attached
		$t_ip = StringLeft($sIP, $port - 1) ;remove the port from the rest of the ip
		If $P_strict Then ;return 0 if port is wrong
			$zport = Int(StringTrimLeft($sIP, $port)) ;retrieve the port
			If $zport > 65000 Or $zport < 0 Then Return 0 ;port is wrong
		EndIf
	EndIf
	$zip = StringSplit($t_ip, ".")
	If $zip[0] <> 4 Then Return 0 ;incorect number of segments
	If Int($zip[1]) > 255 Or Int($zip[1]) < 1 Then Return 0 ;xxx.ooo.ooo.ooo
	If Int($zip[2]) > 255 Or Int($zip[1]) < 0 Then Return 0 ;ooo.xxx.ooo.ooo
	If Int($zip[3]) > 255 Or Int($zip[3]) < 0 Then Return 0 ;ooo.ooo.xxx.ooo
	If Int($zip[4]) > 255 Or Int($zip[4]) < 0 Then Return 0 ;ooo.ooo.ooo.xxx
	$BC = 1 ; is it 255.255.255.255 ?
	For $i = 1 To 4
		If $zip[$i] <> 255 Then $BC = 0 ;no
		;255.255.255.255 can never be a ip but anything else that ends with .255 can be
		;ex:10.10.0.255 can actually be an ip address and not a broadcast address
	Next
	If $BC Then Return 0 ;a broadcast address is not really an ip address...
	If $zip[4] = 0 Then ;subnet not ip
		If $port Then
			Return 0 ;subnet with port?
		Else
			Return 2 ;subnet
		EndIf
	EndIf
	Return 1 ;;string is a ip
EndFunc   ;==>_IsIP

;==============================================================================================
; Description:		_FileRegister($FileExt, $Command, $Verb[, $Default = Default[, $Icon = Default[, $Description = Default]]])
;					Registers a file type in Explorer
; Parameter(s):		$FileExt - 	File Extension without period eg. "zip"
;					$Command - 	Program path with arguments eg. '"C:\test\testprog.exe" "%1"'
;								(%1 is 1st argument, %2 is 2nd, etc.)
;								Setting $Command to an empty string "" will skip setting that key and return the value of an existing key
;					$Verb 	- 	Name of action to perform on file
;								eg. "Open with ProgramName" or "Extract Files"
;					$Default - 	(True/False) The verb will be the default for this filetype
;								If the file is not already associated, this will be the default.
;								Setting default to False will return the current default verb
;					$Icon - 	Default icon for filetype including resource # if needed
;								eg. "C:\test\testprog.exe,0" or "C:\test\filetype.ico"
;					$Description - File Description eg. "Zip File" or "ProgramName Document"
; Returns:  		Returns the new verb is setting that key was a success, the old verb if it was not
;					Sets @extended to the new command if setting that key was a success or the old command if not
; Notes:
; Author:
; Date/Version:   	10/15/2014  --  v2.0.4
;===============================================================================================
Func _FileRegister($FileExt, $Command, $Verb = Default, $Default = Default, $Icon = Default, $Description = Default)
	; FileExt is a key representing the file extention and specifying which FileType to use for that extention
	; FileType is a key containing various properties and actions (verbs) that can be used for files of this type

	If $Default = Default Then $Default = False ; Do not make the new verb the default
	If $Command = Default Or $Command = "" Then $Command = "" ; Allow use of Default to be treated as ""

	; Remove the "." if it was included
	If StringLeft($FileExt, 1) = "." Then $FileExt = StringTrimLeft($FileExt, 1)

	; Get the current FileType for the extention
	Local $FileType = RegRead("HKCR\." & $FileExt, "")
	If @error And $Verb = Default Then
		; If FileExt doesn't exist but a verb wasn't specified then we have nothing to do.
		Return SetError(1, 0, "")
	ElseIf @error Then ; The extention doesn't exist so create it
		RegWrite("HKCR\." & $FileExt, "", "REG_SZ", $FileExt & "file") ; Create a new FileType to use with it
		$FileType = $FileExt & "file" ; Make the new Verb default since it doesn't have one
		$Default = True
	EndIf

	; Verb keys can't have spaces, still use $Verb for a display name
	$VerbKey = StringReplace($Verb, " ", "")

	; Set the default verb to use
	Local $CurrentDefaultVerb = RegRead("HKCR\" & $FileType & "\shell", "")
	If $Default Then
		If Not @error Then RegWrite("HKCR\" & $FileType & "\shell", "oldverb", "REG_SZ", $CurrentDefaultVerb) ; Backup default verb
		RegWrite("HKCR\" & $FileType & "\shell", "", "REG_SZ", $VerbKey) ; Make new verb the default
		If Not @error Then $CurrentDefaultVerb = $VerbKey
	EndIf

	If $Verb <> Default Then
		; The display name for the verb
		Local $CurrentVerbName = RegRead("HKCR\" & $FileType & "\shell\" & $VerbKey, "")
		If Not @error Then RegWrite("HKCR\" & $FileType & "\shell\" & $VerbKey, "oldname", "REG_SZ", $CurrentVerbName) ; Backup verb name
		RegWrite("HKCR\" & $FileType & "\shell\" & $VerbKey, "", "REG_SZ", $Verb) ; Set the display name

		; Command is a subkey reprenseting what happens when a particular verb is selected
		Local $CurrentCommand = RegRead("HKCR\" & $FileType & "\shell\" & $VerbKey & "\command", "")
		If $Command <> Default Then
			If Not @error Then RegWrite("HKCR\" & $FileType & "\shell\" & $VerbKey & "\command", "oldcmd", "REG_SZ", $CurrentCommand) ; Backup command
			RegWrite("HKCR\" & $FileType & "\shell\" & $VerbKey & "\command", "", "REG_SZ", $Command) ; Set the command
			If Not @error Then $CurrentCommand = $Command
		EndIf
	EndIf

	; Specify the icon to be used for the FileType
	If $Icon <> Default Then
		Local $CurrentIcon = RegRead("HKCR\" & $FileType & "\DefaultIcon", "")
		If @error Then RegWrite("HKCR\" & $FileType & "\DefaultIcon", "oldicon", "REG_SZ", $CurrentIcon) ; Backup icon
		RegWrite("HKCR\" & $FileType & "\DefaultIcon", "", "REG_SZ", $Icon)
	EndIf

	; Set the description for the the file type
	If $Description <> Default Then
		Local $CurrentDescription = RegRead("HKCR\" & $FileType, "")
		If @error Then RegWrite("HKCR\" & $FileType, "olddesc", "REG_SZ", $CurrentDescription) ; Backup description
		RegWrite("HKCR\" & $FileType, "", "REG_SZ", $Description) ; Write the description
	EndIf

	Return SetError(0, $CurrentCommand, $CurrentDefaultVerb)
EndFunc   ;==>_FileRegister

;===============================================================================
; Description:		FileUnRegister($ext, $verb)
;					UnRegisters a verb for a file type in Explorer
; Parameter(s):		$ext - File Extension without period eg. "zip"
;					$verb - Name of file action to remove
;							eg. "Open with ProgramName" or "Extract Files"
;===============================================================================
Func _FileUnRegister($ext, $Verb)
	$loc = RegRead("HKCR\." & $ext, "")
	If Not @error Then
		$oldicon = RegRead("HKCR\" & $loc & "\shell", "oldicon")
		If Not @error Then
			RegWrite("HKCR\" & $loc & "\DefaultIcon", "", "REG_SZ", $oldicon)
		Else
			RegDelete("HKCR\" & $loc & "\DefaultIcon", "")
		EndIf
		$oldverb = RegRead("HKCR\" & $loc & "\shell", "oldverb")
		If Not @error Then
			RegWrite("HKCR\" & $loc & "\shell", "", "REG_SZ", $oldverb)
		Else
			RegDelete("HKCR\" & $loc & "\shell", "")
		EndIf
		$olddesc = RegRead("HKCR\" & $loc, "olddesc")
		If Not @error Then
			RegWrite("HKCR\" & $loc, "", "REG_SZ", $olddesc)
		Else
			RegDelete("HKCR\" & $loc, "")
		EndIf
		$oldcmd = RegRead("HKCR\" & $loc & "\shell\" & $Verb & "\command", "oldcmd")
		If Not @error Then
			RegWrite("HKCR\" & $loc & "\shell\" & $Verb & "\command", "", "REG_SZ", $oldcmd)
			RegDelete("HKCR\" & $loc & "\shell\" & $Verb & "\command", "oldcmd")
		Else
			RegDelete("HKCR\" & $loc & "\shell\" & $Verb)
		EndIf
	EndIf
EndFunc   ;==>_FileUnRegister

