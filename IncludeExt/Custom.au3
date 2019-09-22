Func _NetAdapterInfo()
	_Log("Collecting Adapter Information")
	Local $Data[6]
	Local $objWMIService = ObjGet('winmgmts:\\' & @ComputerName & '\root\CIMV2')
	Local $colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=true")
	For $objAdapter In $colItems
		If StringStripWS($objAdapter.DefaultIPGateway(0), 8) <> "" Then
				$Data[1] = $objAdapter.DNSDomain
				$Data[2] = $objAdapter.MACAddress
				$Data[3] = $objAdapter.IPAddress(0)
				$Data[4] = $objAdapter.DefaultIPGateway(0)
				$Data[5] = $objAdapter.IPSubnet(0)
				Return $Data
		EndIf
	Next

	Return SetError(1, 0, $Data)
EndFunc

#cs ===============================================================================
 Function:      _RenameComputer( $iCompName , $iUserName = "" , $iPassword = "" )

 Description:   Renames the local computer

 Parameter(s):  $iCompName: The new computer name

                Required Only if PC is joined to Domain:
                    $iUserName: Username in DOMAIN\UserNamefFormat
                    $iPassword: Password of the specified account

 Returns:       1 - Succeeded (Reboot to take effect)

                0 - Invalid parameters
                    @error 2 - Computername contains invalid characters.
                    @error 3 - Current account does not have sufficient rights
                    @error 4 - Failed to create COM Object

                Returns error code returned by WMI
                    Sets @error 1

 Author(s):  Kenneth Morrissey (ken82m)
#ce ===============================================================================

Func _RenameComputer($iCompName, $iUserName = "", $iPassword = "")
    $Check = StringSplit($iCompName, "`~!@#$%^&*()=+_[]{}\|;:.'"",<>/? ")
    If $Check[0] > 1 Then
        SetError(2)
        Return 0
    EndIf
    If Not IsAdmin() Then
        SetError(3)
        Return 0
    EndIf

    $objWMIService = ObjGet("winmgmts:\root\cimv2")
    If @error Then
        SetError(4)
        Return 0
    EndIf

    For $objComputer In $objWMIService.InstancesOf("Win32_ComputerSystem")
        $oReturn = $objComputer.rename($iCompName,$iPassword,$iUserName)
        If $oReturn <> 0 Then
            SetError(1)
            Return $oReturn
        Else
            Return 1
        EndIf
    Next
EndFunc

Func _WinHTTPRead($sURL, $Agent = "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:15.0) Gecko/20100101 Firefox/15.0.1", $AddHeader = "")
	_Log("_WinHTTPRead " & $sURL)
	; Open needed handles
	Local $hOpen = _WinHttpOpen($Agent)

	Local $iStart = StringInStr($sURL,"/",0,2)+1
	Local $Connect = StringMid($sURL, $iStart, StringInStr($sURL,"/",0,3) - $iStart)

	Local $hConnect = _WinHttpConnect($hOpen, $Connect)

	; Specify the reguest:
	Local $RequestURL = StringTrimLeft($sURL,StringInStr($sURL,"/",0,3))
	Local $hRequest = _WinHttpOpenRequest($hConnect, "GET", $RequestURL, Default, Default, Default, $WINHTTP_FLAG_SECURE + $WINHTTP_FLAG_ESCAPE_DISABLE + $WINHTTP_FLAG_BYPASS_PROXY_CACHE)

	_WinHttpAddRequestHeaders ($hRequest, "Cache-Control: no-cache")
	_WinHttpAddRequestHeaders ($hRequest, "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3")
	_WinHttpAddRequestHeaders ($hRequest, "content-type: application/json")

	If $AddHeader <> "" Then
		_WinHttpAddRequestHeaders ($hRequest, $TokenAddHeader)
	Endif

	; Send request
	_WinHttpSendRequest($hRequest)
	If @error Then
		_WinHttpCloseHandle($hRequest)
		_WinHttpCloseHandle($hConnect)
		_WinHttpCloseHandle($hOpen)
		_Log("Connection error (Send)")
		Return SetError(1, 0, 0)
	Endif

	; Wait for the response
	_WinHttpReceiveResponse($hRequest)
	If @error Then
		_WinHttpCloseHandle($hRequest)
		_WinHttpCloseHandle($hConnect)
		_WinHttpCloseHandle($hOpen)
		_Log("Connection error (Receive)")
		Return SetError(2, 0, 0)
	Endif

	Local $sHeader = _WinHttpQueryHeaders($hRequest) ; ...get full header

	Local $bData, $bChunk
	While 1
		$bChunk = _WinHttpReadData($hRequest, 2)
		If @error Then ExitLoop
		$bData = _WinHttpBinaryConcat($bData, $bChunk)
	WEnd

	; Clean
	_WinHttpCloseHandle($hRequest)
	_WinHttpCloseHandle($hConnect)
	_WinHttpCloseHandle($hOpen)

	Return $bData

EndFunc