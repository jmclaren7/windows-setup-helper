#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=icon.ico
#AutoIt3Wrapper_Outfile_x64=VNCHelper.exe
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Res_Fileversion=1.0.0.49
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_Language=1033
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include <Array.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <GuiStatusBar.au3>
#include <ListViewConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>

Global $Title = "VNCHelper"
Global $ViewerFullPath = @TempDir & "\tvnviewer.exe"

ConsoleWrite(@CRLF&"Start"&@CRLF)


Opt("GUIOnEventMode", 1)
Opt("TCPTimeout", 60)


TCPStartup()

Global $PortInput, $PasswordInput, $HostListView
$IPDefault = _DefaultIPRange()

ConsoleWrite("GUI"&@CRLF)

#Region ### START Koda GUI section
$MainGUI = GUICreate("VNCWatch", 371, 204, -1, -1)
$Button1 = GUICtrlCreateButton("Connect", 184, 153, 179, 25)
$PortInput = GUICtrlCreateInput("5950", 8, 21, 153, 21)
$PasswordInput = GUICtrlCreateInput("vncwatch", 8, 64, 153, 21, $ES_PASSWORD)
$Label2 = GUICtrlCreateLabel("Port", 8, 4, 23, 17)
GUICtrlCreateLabel("Password", 8, 47, 50, 17)
$AutoConnectCheckbox = GUICtrlCreateCheckbox("Automatically Connect", 8, 134, 153, 17)
GUICtrlSetState(-1, $GUI_UNCHECKED)
$HostListView = GUICtrlCreateListView("", 184, 8, 178, 142, BitOR($GUI_SS_DEFAULT_LISTVIEW,$LVS_SMALLICON,$LVS_NOCOLUMNHEADER,$LVS_NOLABELWRAP,$LVS_AUTOARRANGE,$WS_VSCROLL))
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 0, 178)
$Input1 = GUICtrlCreateInput($IPDefault, 8, 106, 153, 21)
GUICtrlCreateLabel("IP Range (CIDR)", 9, 89, 84, 17)
$StatusBar1 = _GUICtrlStatusBar_Create($MainGUI)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

WinSetTitle($MainGUI, "", $Title)
GUISetOnEvent($GUI_EVENT_CLOSE, "_Exit")
GUICtrlSetOnEvent($Button1, "_Connect" )

Global $aListViewItems[0]

ConsoleWrite("Ready"&@CRLF)

While 1
	$IPInput = GUICtrlRead ($Input1)
	$aIPList = _GetIpAddressList($IPInput)
	_GUICtrlStatusBar_SetText($StatusBar1, "Scanning... " & Ubound($aIPList))

	For $i = 0 To Ubound($aIPList)-1
		$IP = $aIPList[$i]
		;ConsoleWrite("$IP="&$IP&@CRLF)
		_GUICtrlStatusBar_SetText($StatusBar1, "Scanning: "&$IP)
		If $IPInput <> GUICtrlRead ($Input1) Then ExitLoop

		$iSocket = TCPConnect($IP, GUICtrlRead ($PortInput))
		ConsoleWrite("TCPConnect: "&@error&@CRLF)

		If $iSocket > 0 Then
			$MarkedForRemoval = False

			ConsoleWrite("Success: " & $IP & @CRLF)
			If _GUICtrlListView_FindText($HostListView, $IP, -1, False) = -1 Then
				$ListViewItem = GUICtrlCreateListViewItem($IP, $HostListView)
				ConsoleWrite("Added " & $IP & @CRLF)

				If GUICtrlRead($AutoconnectCheckbox) = $GUI_CHECKED Then
					$Index = _GUICtrlListView_FindText($HostListView, $IP, -1, False)
					ConsoleWrite("Index: "&$Index&@CRLF)
					$Return = _GUICtrlListView_SetItemSelected($HostListView, $Index, True, True)
					ConsoleWrite("Return: "&$Return&@CRLF)

					_Connect()
				Endif

				$iLV_Width = 0
				For $b = 0 To 2
					GUICtrlSendMsg($HostListView, $LVM_SETCOLUMNWIDTH, $b, $LVSCW_AUTOSIZE)
					;$iLV_Width += GUICtrlSendMsg($List1, $LVM_GETCOLUMNWIDTH, $b, 0)
				Next
			EndIf

		Else
			;ConsoleWrite("Failed: " & $IP & @CRLF)
			$ListItemIndex = _GUICtrlListView_FindText($HostListView, $IP, -1, False, True)
			If $ListItemIndex <> -1 Then
				ConsoleWrite("$ListItemIndex=" & $ListItemIndex & @CRLF)
				If $MarkedForRemoval = True Then
					_GUICtrlListView_DeleteItem($HostListView, $ListItemIndex)
					ConsoleWrite("Removed " & $IP & @CRLF)
					$MarkedForRemoval = False
				Else
					$MarkedForRemoval = True
					TCPCloseSocket($iSocket)
					Sleep(200)
					$i = $i - 1
				EndIf


			Endif


		EndIf

		TCPCloseSocket($iSocket)
	Next

	Sleep(10)
WEnd

TCPShutdown()

;=========== =========== =========== =========== =========== =========== =========== ===========

Func _DefaultIPRange()
	Local $IPConfig
	$objWMIService = ObjGet("winmgmts:\\" & "localhost" & "\root\cimv2")
	$IPConfigSet = $objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration ")

	For $IPConfig in $IPConfigSet
			For $x = 0 to UBound($IPConfig.IPAddress) - 1
				If Not StringInStr($IPConfig.IPAddress($x), ".") Then ContinueLoop
				ConsoleWrite("IP: "&$IPConfig.IPAddress($x)&@CRLF)

				For $y = 0 to UBound($IPConfig.DefaultIPGateway) - 1
					ConsoleWrite("Gateway: "&$IPConfig.DefaultIPGateway($y)&@CRLF)
					$CIDR = StringLeft($IPConfig.IPAddress($x),StringInStr($IPConfig.DefaultIPGateway($y),".",0,-1)) & "*"
					ConsoleWrite($CIDR&@CRLF)
					Return $CIDR
				Next
			Next
	Next

	Return StringLeft(@IPAddress1,StringInStr(@IPAddress1,".",0,-1)) & "*"
Endfunc

Func _Connect()
	ConsoleWrite("_Connect" & @CRLF)

	Local $Selected = Int(_GUICtrlListView_GetSelectedIndices($HostListView))

	ConsoleWrite("$Selected=" & $Selected & @CRLF)

	If $Selected >= 0 Then
		Local $IP = _GUICtrlListView_GetItemText($HostListView, $Selected)
		$Execute = $ViewerFullPath & " " & $IP & "::" & GUICtrlRead ($PortInput) & " -password=" & GUICtrlRead ($PasswordInput)
		ConsoleWrite("$Execute=" & $Execute & @CRLF)
		FileInstall("tvnviewer.exe", $ViewerFullPath)
		If $IP <> "" Then Run(@ComSpec & " /c " & $Execute, "", @SW_HIDE)
	Endif

EndFunc

Func _Exit()
	Exit
Endfunc

Func _Log($data)
	ConsoleWrite(@CRLF & $data)
EndFunc

Func _GetIpAddressList($ipFormat)
    Local $aResult[1]

    If StringInStr(StringStripWS($ipFormat, 3), " ") Or StringInStr($ipFormat, ".", "", 4) Then Return $aResult

    $ipFormat = StringReplace($ipFormat, "*", "1-255")
    $ipSplit = StringSplit($ipFormat, ".")

    If $ipSplit[0] <> 4 Then Return $aResult

    Local $ipRange[4][2], $totalPermu = 1

    For $i = 0 To 3
        If StringInStr($ipSplit[$i + 1], "-") Then
            If StringInStr($ipSplit[$i + 1], "-", "", 2) Then Return $aResult
            $m = StringSplit($ipSplit[$i +1 ],"-")
            For $i2 = 1 to $m[0]
                If Number($m[$i2]) > 255 Or Number($m[$i2]) < 0 Then Return $aResult
                $ipRange[$i][$i2 - 1] = Number($m[$i2])
            Next
        Else
            $n = Number($ipSplit[$i + 1])
            If $n > 255 Or $n < 0 Then Return $aResult
            $ipRange[$i][0] = $n
            $ipRange[$i][1] = $n
        EndIf

        $totalPermu *= $ipRange[$i][1] - $ipRange[$i][0] + 1
    Next

    Local $aResult[$totalPermu], $i = 0

    For $a = $ipRange[0][0] To $ipRange[0][1]
        For $b = $ipRange[1][0] To $ipRange[1][1]
            For $c = $ipRange[2][0] To $ipRange[2][1]
                For $d = $ipRange[3][0] To $ipRange[3][1]
                    $aResult[$i] = $a & "." & $b & "." & $c & "." & $d
                    $i += 1
                Next
            Next
        Next
    Next

    Return $aResult
EndFunc   ;==>_GetIpAddressList
