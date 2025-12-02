; Script to generate a One-Time Password (OTP) from challenge code displayed on WinPE boot.

#include "Crypt.au3"

$ConfigFilePath = @ScriptDir & "\Helper\Config.ini"
$AccessOTPSecret = IniRead($ConfigFilePath, "Access", "OTPSecret", "")
$AccessSalt = IniRead($ConfigFilePath, "Access", "Salt", "3b194da2")
$ChallengeResponseLength = 6
$ValidOTPResponse = ""
$Input = ""
$ValidOTPResponseText = ""

While 1

    $Input = InputBox("OTP Generator", "Enter the challenge code displayed on the WinPE boot screen." & @CRLF & @CRLF & $ValidOTPResponseText, $Input)
    If @error Then Exit


    $ValidOTPResponse = StringRight(_Crypt_HashData($AccessOTPSecret & StringLower($Input) & $AccessSalt, $CALG_SHA_256), $ChallengeResponseLength)
    $ValidOTPResponseText = "OTP response is: " & $ValidOTPResponse
WEnd