Func IsActivated()
  $oWMIService = ObjGet("winmgmts:\\.\root\cimv2")
  If IsObj($oWMIService) Then
    $oCollection = $oWMIService.ExecQuery("SELECT Description, LicenseStatus, GracePeriodRemaining FROM SoftwareLicensingProduct WHERE PartialProductKey <> null")
    If IsObj($oCollection) Then
      For $oItem In $oCollection
        Switch $oItem.LicenseStatus
          Case 0
            Return SetError(0, 1, "Unlicensed")

          Case 1
            If $oItem.GracePeriodRemaining Then
              If StringInStr($oItem.Description, "TIMEBASED_") Then
                Return SetError(0, 0, "Timebased activation will expire in " & $oItem.GracePeriodRemaining & " minutes")

              Else
                Return SetError(0, 0, "Volume activation will expire in " & $oItem.GracePeriodRemaining & " minutes")

			  EndIf
            Else
              Return SetError(0, 0, "The machine is permanently activated.")

			EndIf

          Case 2
            Return SetError(0, 2, "Initial grace period ends in " & $oItem.GracePeriodRemaining & " minutes")

          Case 3
            Return SetError(0, 3, "Additional grace period ends in " & $oItem.GracePeriodRemaining & " minutes")

          Case 4
            Return SetError(0, 4, "Non-genuine grace period ends in " & $oItem.GracePeriodRemaining & " minutes")

          Case 5
            Return SetError(0, 5, "Windows is in Notification mode")

          Case 6
            Return SetError(0, 6, "Extended grace period ends in " & $oItem.GracePeriodRemaining & " minutes")

        EndSwitch
      Next

    Else
      Return SetError(2, 0, "WMI Query Error")
      
    EndIf

  Else
    Return SetError(1, 0, "WMI Object Error")

  EndIf
EndFunc