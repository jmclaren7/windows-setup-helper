
$Default = ""
While 1
	$Command=InputBox(@ScriptName,"Enter command", $Default)
	if @error then continueloop
	$Default = $Command
	$aCommand=StringSplit($Command," ")

	Switch $aCommand[1]
		Case "exit"
			Exit

		Case "res"
			_ChangeScreenRes($aCommand[2],$aCommand[3])

		Case "1"
			_ChangeScreenRes(800,600)

		Case "2"
			_ChangeScreenRes(1024,768)

		Case "3"
			_ChangeScreenRes(1152,864)

	EndSwitch

WEnd

;===============================================================================
;
; Function Name:    _ChangeScreenRes()
; Description:      Changes the current screen geometry, colour and refresh rate.
; Version:          2
; Parameter(s):     $i_Width - Width of the desktop screen in pixels. (horizontal resolution)
;                   $i_Height - Height of the desktop screen in pixels. (vertical resolution)
;					$i_BitsPP - Depth of the desktop screen in bits per pixel.
;					$i_RefreshRate - Refresh rate of the desktop screen in hertz.
; Requirement(s):   AutoIt 3.3.6.1
; Return Value(s):  On Success - Screen is adjusted, @ERROR = 0
;                   On Failure - sets @ERROR = 1
; Forum(s):         http://www.autoitscript.com/forum/index.php?showtopic=20121
; Author(s):        Original code - psandu.ro
;                   Modifications - PartyPooper
;
;                   HEAVY Modifications by KaFu
;
;===============================================================================
Func _ChangeScreenRes($i_Width = @DesktopWidth, $i_Height = @DesktopHeight, $i_BitsPP = @DesktopDepth, $i_RefreshRate = @DesktopRefresh)

   Local Const $ENUM_CURRENT_SETTINGS = -1
   Local Const $ENUM_REGISTRY_SETTINGS = -2
   Local Const $DMDO_90 = 1
   Local Const $DMDO_270 = 3
   Local Const $_tag_POINTL = "long x;long y"
   Local Const $_tag_DEVMODE = "char dmDeviceName[32];ushort dmSpecVersion;ushort dmDriverVersion;short dmSize;" & _
		"ushort dmDriverExtra;dword dmFields;" & $_tag_POINTL & ";dword dmDisplayOrientation;dword dmDisplayFixedOutput;" & _
		"short dmColor;short dmDuplex;short dmYResolution;short dmTTOption;short dmCollate;" & _
		"byte dmFormName[32];ushort LogPixels;dword dmBitsPerPel;int dmPelsWidth;dword dmPelsHeight;" & _
		"dword dmDisplayFlags;dword dmDisplayFrequency"

   Local $h_DLL_user32 = DllOpen("user32.dll")
   Local $b_Res_Orientation_Rotation_Supported = False

    Local Const $DM_ORIENTATION = 0x00000001
	Local Const $DM_PAPERSIZE = 0x00000002
	Local Const $DM_PAPERLENGTH = 0x00000004
	Local Const $DM_PAPERWIDTH = 0x00000008
	Local Const $DM_SCALE = 0x00000010
	Local Const $DM_COPIES = 0x00000100
	Local Const $DM_DEFAULTSOURCE = 0x00000200
	Local Const $DM_PRINTQUALITY = 0x00000400
	Local Const $DM_POSITION = 0x00000020
	Local Const $DM_DISPLAYORIENTATION = 0x00000080
	Local Const $DM_DISPLAYFIXEDOUTPUT = 0x20000000
	Local Const $DM_COLOR = 0x00000800
	Local Const $DM_DUPLEX = 0x00001000
	Local Const $DM_YRESOLUTION = 0x00002000
	Local Const $DM_TTOPTION = 0x00004000
	Local Const $DM_COLLATE = 0x00008000
	Local Const $DM_FORMNAME = 0x00010000
	Local Const $DM_LOGPIXELS = 0x00020000
	Local Const $DM_BITSPERPEL = 0x00040000
	Local Const $DM_PELSWIDTH = 0x00080000
	Local Const $DM_PELSHEIGHT = 0x00100000
	Local Const $DM_DISPLAYFLAGS = 0x00200000
	Local Const $DM_NUP = 0x00000040
	Local Const $DM_DISPLAYFREQUENCY = 0x00400000
	Local Const $DM_ICMMETHOD = 0x00800000
	Local Const $DM_ICMINTENT = 0x01000000
	Local Const $DM_MEDIATYPE = 0x02000000
	Local Const $DM_DITHERTYPE = 0x04000000
	Local Const $DM_PANNINGWIDTH = 0x08000000
	Local Const $DM_PANNINGHEIGHT = 0x10000000

	Local Const $DM_DISPLAYQUERYORIENTATION = 0x01000000

	Local Const $CDS_TEST = 0x00000002
	Local Const $CDS_UPDATEREGISTRY = 0x00000001
	Local Const $CDS_RESET = 0x40000000
	Local Const $CDS_SET_PRIMARY = 0x00000010

	Local Const $CDS_VIDEOPARAMETERS = 0x00000020
	Local Const $CDS_ENABLE_UNSAFE_MODES = 0x00000100
	Local Const $CDS_DISABLE_UNSAFE_MODES = 0x00000200

	; error 2 = EnumDisplaySettingsEx for $ENUM_CURRENT_SETTINGS failed
	Local Const $DISP_CHANGE_SUCCESSFUL = 0
	Local Const $DISP_CHANGE_FAILED = -1
	Local Const $DISP_CHANGE_BADMODE = -2
	Local Const $DISP_CHANGE_NOTUPDATED = -3
	Local Const $DISP_CHANGE_BADFLAGS = -4
	Local Const $DISP_CHANGE_BADPARAM = -5
	Local Const $DISP_CHANGE_BADDUALVIEW = -6
	Local Const $DISP_CHANGE_RESTART = 1

	Local Const $HWND_BROADCAST = 0xffff
	Local Const $WM_DISPLAYCHANGE = 0x007E

	If $i_Width = "" Or $i_Width = -1 Then $i_Width = @DesktopWidth ; default to current setting
	If $i_Height = "" Or $i_Height = -1 Then $i_Height = @DesktopHeight ; default to current setting
	If $i_BitsPP = "" Or $i_BitsPP = -1 Then $i_BitsPP = @DesktopDepth ; default to current setting
	If $i_RefreshRate = "" Or $i_RefreshRate = -1 Then $i_RefreshRate = @DesktopRefresh ; default to current setting

	Local $DEVMODE = DllStructCreate($_tag_DEVMODE)
	DllStructSetData($DEVMODE, "dmSize", DllStructGetSize($DEVMODE))

	; Using the dmFields flag of DM_DISPLAYORIENTATION, ChangeDisplaySettingsEx can be used to dynamically rotate the screen orientation. However, the DM_PELSWIDTH and DM_PELSHEIGHT flags cannot be used to change the screen resolution.

	Local $i_DllRet = DllCall($h_DLL_user32, "int", "EnumDisplaySettingsEx", "ptr", 0, "dword", $ENUM_CURRENT_SETTINGS, "ptr", DllStructGetPtr($DEVMODE), "dword", 0)
	If $i_DllRet[0] = 0 Then
		$i_DllRet = DllCall($h_DLL_user32, "int", "EnumDisplaySettingsEx", "ptr", 0, "dword", $ENUM_REGISTRY_SETTINGS, "ptr", DllStructGetPtr($DEVMODE), "dword", 0)
	EndIf

	#cs
		ConsoleWrite("dmDeviceName " & DllStructGetData($DEVMODE, "dmDeviceName") & @CRLF)
		ConsoleWrite("dmSpecVersion " & DllStructGetData($DEVMODE, "dmSpecVersion") & @CRLF)
		ConsoleWrite("dmDriverVersion " & DllStructGetData($DEVMODE, "dmDriverVersion") & @CRLF)
		ConsoleWrite("dmSize " & DllStructGetData($DEVMODE, "dmSize") & @CRLF)
		ConsoleWrite("dmDriverExtra " & DllStructGetData($DEVMODE, "dmDriverExtra") & @CRLF)
		ConsoleWrite("dmFields " & DllStructGetData($DEVMODE, "dmFields") & @CRLF)
		ConsoleWrite("dmPositionx " & DllStructGetData($DEVMODE, "dmPositionx") & @CRLF)
		ConsoleWrite("dmPositiony " & DllStructGetData($DEVMODE, "dmPositiony") & @CRLF)
		ConsoleWrite("- dmDisplayOrientation " & DllStructGetData($DEVMODE, "dmDisplayOrientation") & @CRLF)
		ConsoleWrite("dmDisplayFixedOutput " & DllStructGetData($DEVMODE, "dmDisplayFixedOutput") & @CRLF)
		ConsoleWrite("dmColor " & DllStructGetData($DEVMODE, "dmColor") & @CRLF)
		ConsoleWrite("dmDuplex " & DllStructGetData($DEVMODE, "dmDuplex") & @CRLF)
		ConsoleWrite("dmYResolution " & DllStructGetData($DEVMODE, "dmYResolution") & @CRLF)
		ConsoleWrite("dmTTOption " & DllStructGetData($DEVMODE, "dmTTOption") & @CRLF)
		ConsoleWrite("dmCollate " & DllStructGetData($DEVMODE, "dmCollate") & @CRLF)
		ConsoleWrite("dmFormName " & DllStructGetData($DEVMODE, "dmFormName") & @CRLF)
		ConsoleWrite("dmLogPixels " & DllStructGetData($DEVMODE, "dmLogPixels") & @CRLF)
		ConsoleWrite("dmBitsPerPel " & DllStructGetData($DEVMODE, "dmBitsPerPel") & @CRLF)
		ConsoleWrite("dmPelsWidth " & DllStructGetData($DEVMODE, "dmPelsWidth") & @CRLF)
		ConsoleWrite("dmPelsHeight " & DllStructGetData($DEVMODE, "dmPelsHeight") & @CRLF)
		ConsoleWrite("dmDisplayFlags " & DllStructGetData($DEVMODE, "dmDisplayFlags") & @CRLF)
		ConsoleWrite("dmDisplayFrequency " & DllStructGetData($DEVMODE, "dmDisplayFrequency") & @CRLF & @CRLF)
	#ce

	If @error Then
		$DEVMODE = 0
		SetError(1)
		Return 1
	Else
		$i_DllRet = $i_DllRet[0]
	EndIf

	If $i_DllRet <> 0 Then

		DllStructSetData($DEVMODE, "dmPelsWidth", $i_Width)
		DllStructSetData($DEVMODE, "dmPelsHeight", $i_Height)
		DllStructSetData($DEVMODE, "dmBitsPerPel", $i_BitsPP)
		DllStructSetData($DEVMODE, "dmDisplayFrequency", $i_RefreshRate)

		If $b_Res_Orientation_Rotation_Supported Then
			If $i_Height < $i_Width Then
				If IniRead(@ScriptDir & "\HRC.ini", 'Settings', 'Default_Orientation', "090") = "090" Then
					DllStructSetData($DEVMODE, "dmDisplayOrientation", $DMDO_90)
				Else
					DllStructSetData($DEVMODE, "dmDisplayOrientation", $DMDO_270)
				EndIf
			Else
				DllStructSetData($DEVMODE, "dmFields", BitOR($DM_POSITION, $DM_PELSWIDTH, $DM_PELSHEIGHT, $DM_BITSPERPEL, $DM_DISPLAYFREQUENCY))
			EndIf
		Else
			DllStructSetData($DEVMODE, "dmFields", BitOR($DM_POSITION, $DM_PELSWIDTH, $DM_PELSHEIGHT, $DM_BITSPERPEL, $DM_DISPLAYFREQUENCY))
		EndIf

		$i_DllRet = DllCall($h_DLL_user32, "int", "ChangeDisplaySettingsEx", "ptr", 0, "ptr", DllStructGetPtr($DEVMODE), "hwnd", 0, "int", $CDS_TEST, "ptr", 0)
		If @error Then
			$DEVMODE = 0
			SetError(2)
			Return 2
		Else
			$i_DllRet = $i_DllRet[0]
		EndIf

		Select
			Case $i_DllRet = $DISP_CHANGE_SUCCESSFUL
				$i_DllRet = DllCall($h_DLL_user32, "int", "ChangeDisplaySettingsEx", "ptr", 0, "ptr", DllStructGetPtr($DEVMODE), "hwnd", 0, "int", $CDS_UPDATEREGISTRY, "ptr", 0)
				If @error Then
					$DEVMODE = 0
					SetError(2)
					Return 3
				Else
					$i_DllRet = $i_DllRet[0]
				EndIf

				If $i_DllRet <> $DISP_CHANGE_SUCCESSFUL Then
					$DEVMODE = 0
					SetError($i_DllRet)
					Return 3
				EndIf

				;WinMove($hGUI_HRC_Main, "", @DesktopWidth / 2 - (420 / 2), @DesktopHeight / 2 - ((228 + (($iHotKeyBoxOffset - 3) * 80)) / 2))
				;_SendMessageTimeout_Ex($HWND_BROADCAST, $WM_DISPLAYCHANGE, $i_BitsPP, $i_Height * 2 ^ 16 + $i_Width)
				Return 0 ; Success !

			Case Else
				$DEVMODE = 0
				SetError($i_DllRet)
				Return 2

		EndSelect
	EndIf
	$DEVMODE = 0
	SetError(2)
	Return 1
EndFunc   ;==>_ChangeScreenRes
