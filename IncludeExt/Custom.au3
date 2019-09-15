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