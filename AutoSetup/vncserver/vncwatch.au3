#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Change2CUI=n
#AutoIt3Wrapper_Res_Fileversion=1.0.0.25
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_Language=1033
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

;#include <ButtonConstants.au3>
;#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
;#include <GUIListBox.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
;#include <Array.au3>
;#include <ListViewConstants.au3>
#include <GuiListView.au3>

ConsoleWrite(@CRLF&"Start"&@CRLF)

Opt("GUIOnEventMode", 1)
Opt("TCPTimeout", 10)

TCPStartup()

Global $Title = "VNCWatch"
Global $PortInput, $PasswordInput, $HostListView

#Region ### START Koda GUI section
$MainGUI = GUICreate($Title, 373, 191, -1, -1)
$Label1 = GUICtrlCreateLabel("Loading...", 8, 144, 158, 33, $SS_CENTER)
$Button1 = GUICtrlCreateButton("Connect", 184, 153, 179, 25)
$PortInput = GUICtrlCreateInput("5950", 8, 26, 153, 21)
$PasswordInput = GUICtrlCreateInput("vncwatch", 8, 74, 153, 21)
$Label2 = GUICtrlCreateLabel("Port", 8, 9, 23, 17)
GUICtrlCreateLabel("Password", 8, 57, 50, 17)
$AutoConnectCheckbox = GUICtrlCreateCheckbox("Automatically Connect", 8, 112, 153, 17)
GUICtrlSetState(-1, $GUI_CHECKED)
$HostListView = GUICtrlCreateListView("", 184, 8, 178, 142, BitOR($GUI_SS_DEFAULT_LISTVIEW,$LVS_SMALLICON,$LVS_NOCOLUMNHEADER,$LVS_NOLABELWRAP,$LVS_AUTOARRANGE,$WS_VSCROLL))
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 0, 178)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

GUISetOnEvent($GUI_EVENT_CLOSE, "_Exit")
GUICtrlSetOnEvent($Button1, "_Connect" )

Global $aListViewItems[0]


While 1
	$NetInfo = _NetAdapterInfo()
	$aIPList = _GetIPList($NetInfo[6], $NetInfo[7])
	GUICtrlSetData($Label1, "Looking For VNC Servers..." & @CRLF & $NetInfo[6] & "-" & $NetInfo[7])

	For $i = 0 To Ubound($aIPList)-1
		$IP = $aIPList[$i]


		$iSocket = TCPConnect($IP, GUICtrlRead ($PortInput))
		If $iSocket > 0 Then
			$MarkedForRemoval = False

			;ConsoleWrite("Success: " & $IP & @CRLF)
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
				For $i = 0 To 2
					GUICtrlSendMsg($HostListView, $LVM_SETCOLUMNWIDTH, $i, $LVSCW_AUTOSIZE)
					;$iLV_Width += GUICtrlSendMsg($List1, $LVM_GETCOLUMNWIDTH, $i, 0)
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


Func _Connect()
	ConsoleWrite("_Connect" & @CRLF)

	Local $Selected = Int(_GUICtrlListView_GetSelectedIndices($HostListView))
	ConsoleWrite("$Selected=" & $Selected & @CRLF)

	If $Selected >= 0 Then
		Local $IP = _GUICtrlListView_GetItemText($HostListView, $Selected)
		$Execute = "vncviewer.exe " & $IP & "::" & GUICtrlRead ($PortInput) & " -password=" & GUICtrlRead ($PasswordInput)
		ConsoleWrite("$Execute=" & $Execute & @CRLF)
		If $IP <> "" Then Run(@ComSpec & " /c " & $Execute, "", @SW_HIDE)
	Endif

EndFunc

Func _ConnectOld()
	ConsoleWrite("_Connect" & @CRLF)

	Local $IP, $Item

	$Item = GUICtrlRead($HostListView)
	$IP = GUICtrlRead($Item)
	If StringRight($IP, 1) = "|" Then $IP = StringLeft($IP, StringInStr($IP,"|")-1)

	ConsoleWrite("$IP: " & $IP & @CRLF)

	;If $IP <> "" Then Run(@ComSpec & " /c " & "vncviewer.exe " & $IP & "::" & GUICtrlRead ($PortInput) & " -password=" & GUICtrlRead ($PasswordInput), "", @SW_HIDE)


Endfunc

Func _Exit()
	Exit
Endfunc

Func _GetIPList($Start, $End)
	Local $aStart = StringSplit($Start, ".")
	Local $aEnd = StringSplit($End, ".")
	Local $aReturn[0]

	For $Oct1=$aStart[1] to $aEnd[1]

		Local $Start2 = $aStart[2]
		If $Oct1 <> $aStart[1] Then $Start2 = 0
		Local $End2 = $aEnd[2]
		If $Oct1 < $aEnd[1] Then $End2 = 254

		For $Oct2=$Start2 to $End2

			Local $Start3 = $aStart[3]
			If $Oct2 <> $aStart[2] Then $Start3 = 0
			Local $End3 = $aEnd[3]
			If $Oct2 < $aEnd[2] Then $End3 = 254

			For $Oct3=$Start3 to $End3

				Local $Start4 = $aStart[4]
				If $Oct3 <> $aStart[3] Then $Start4 = 1
				Local $End4 = $aEnd[4]
				If $Oct3 <> $aEnd[3] Then $End4 = 254

				For $Oct4=$Start4 to $End4
					$IP = $Oct1&"."&$Oct2&"."&$Oct3&"."&$Oct4
					;ConsoleWrite($IP & @CRLF)
					$Index = UBound($aReturn)
					ReDim $aReturn[$Index+1]
					$aReturn[$Index] = $IP
				Next
			Next
		Next
	Next

	Return $aReturn
Endfunc

Func _NetAdapterInfo()
	Local $aData[9]
	Local $objWMIService = ObjGet('winmgmts:\\' & @ComputerName & '\root\CIMV2')
	Local $colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=true")
	For $objAdapter In $colItems
		If StringStripWS($objAdapter.DefaultIPGateway(0), 8) <> "" Then
				$aData[1] = $objAdapter.DNSDomain
				$aData[2] = $objAdapter.MACAddress
				$aData[3] = $objAdapter.IPAddress(0)
				$aData[4] = $objAdapter.DefaultIPGateway(0)
				$aData[5] = $objAdapter.IPSubnet(0)

				Local $aIP = StringSplit($aData[3], ".")
				Local $aSubnetMask = StringSplit($aData[5], ".")
				Local $aSubnetAddress[5]
				Local $aInverseMask[5]
				Local $aBroadcastAddress[5]

				For $i = 1 To $aIP[0]
				   $aSubnetAddress[$i] = BitAND($aIP[$i], $aSubnetMask[$i])
				   $aInverseMask[$i] = BitNOT($aSubnetMask[$i] - 256)
				   $aBroadcastAddress[$i] = BitOR($aSubnetAddress[$i], $aInverseMask[$i])
				Next

				;Start
				$aData[6] = $aSubnetAddress[1] & "." & $aSubnetAddress[2] & "." & $aSubnetAddress[3] & "." & $aSubnetAddress[4] + 1
				;End
				$aData[7] = $aBroadcastAddress[1] & "." & $aBroadcastAddress[2] & "." & $aBroadcastAddress[3] & "." & $aBroadcastAddress[4] - 1
				;Address Count
				$aData[8] = ($aInverseMask[4] + 1) * ($aInverseMask[3] + 1) * ($aInverseMask[2] + 1) - 2

				;ConsoleWrite(_ArrayToString ($aData, @CRLF) & @CRLF)

				Return $aData
		EndIf
	Next

	Return SetError(1, 0, $aData)
EndFunc