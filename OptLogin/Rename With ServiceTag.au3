

$wbemFlagReturnImmediately = 0x10
$wbemFlagForwardOnly = 0x20
$colItems = ""
$Name = "Windows"

$objWMIService = ObjGet("winmgmts:\\localhost\root\CIMV2")
$colItems = $objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct", "WQL", _
                                          $wbemFlagReturnImmediately + $wbemFlagForwardOnly)

If IsObj($colItems) then
	For $objItem In $colItems
   		$Name = $Name&"-"&$objItem.IdentifyingNumber
	Next
Endif

_RenameComputer($Name)



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