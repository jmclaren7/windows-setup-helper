Func IsActivated()
 $oWMIService = ObjGet("winmgmts:\\.\root\cimv2")
 If IsObj($oWMIService) Then
  $oCollection = $oWMIService.ExecQuery("SELECT Description, LicenseStatus, GracePeriodRemaining FROM SoftwareLicensingProduct WHERE PartialProductKey <> null")
  If IsObj($oCollection) Then
   For $oItem In $oCollection
    Switch $oItem.LicenseStatus
     Case 0
      ConsoleWrite("Unlicensed" & @CRLF)
      Return False
     Case 1
      If $oItem.GracePeriodRemaining Then
       If StringInStr($oItem.Description, "TIMEBASED_") Then
        ConsoleWrite("Timebased activation will expire in " & $oItem.GracePeriodRemaining & " minutes" & @CRLF)
        Return False
       Else
        ConsoleWrite("Volume activation will expire in " & $oItem.GracePeriodRemaining & " minutes" & @CRLF)
        Return False
       EndIf
      Else
       ConsoleWrite("The machine is permanently activated." & @CRLF)
       Return True
      EndIf
     Case 2
      ConsoleWrite("Initial grace period ends in " & $oItem.GracePeriodRemaining & " minutes" & @CRLF)
      Return False
     Case 3
      ConsoleWrite("Additional grace period ends in " & $oItem.GracePeriodRemaining & " minutes" & @CRLF)
      Return False
     Case 4
      ConsoleWrite("Non-genuine grace period ends in " & $oItem.GracePeriodRemaining & " minutes" & @CRLF)
      Return False
     Case 5
      ConsoleWrite("Windows is in Notification mode" & @CRLF)
      Return False
     Case 6
      ConsoleWrite("Extended grace period ends in " & $oItem.GracePeriodRemaining & " minutes" & @CRLF)
      Return False
    EndSwitch
   Next
  EndIf
 EndIf
EndFunc