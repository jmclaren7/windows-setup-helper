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
        $Return = $objComputer.rename($iCompName,$iPassword,$iUserName)
        If $oReturn <> 0 Then
            SetError(1)
            Return $oReturn
        Else
            Return 1
        EndIf
    Next
EndFunc