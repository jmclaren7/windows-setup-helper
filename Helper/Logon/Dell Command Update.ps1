If ($(Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer -like "*Dell*"){
    "Waiting for user input to continue Dell command update"

    Start-Sleep 2

    $Wscript_Shell = New-Object -ComObject "Wscript.Shell"
    $MsgBox = $Wscript_Shell.Popup("Dell command update will run automaticly in 20 seconds, press 'ok' to run now or 'cancel' to abort.", 20, "Setup Helper", 1+32)

    If ($MsgBox -ne 1 -AND $MsgBox -ne -1) { Exit }

    $ProgressPreference = 'SilentlyContinue';
    $ua='Mozilla/5.0 (Windows NT; Windows NT 10.0; en-US) AppleWebKit/534.6 (KHTML, like Gecko) Chrome/7.0.500.0 Safari/534.6';
    Invoke-WebRequest 'https://dl.dell.com/FOLDER10408436M/1/Dell-Command-Update-Windows-Universal-Application_1WR6C_WIN_5.0.0_A00.EXE' -useragent $ua -outfile 'dcu.exe'

    Start-Process "dcu.exe" -Args "/s" -Wait -NoNewWindow

    echo 'Waiting for Dell Command Update to install...'
    do{$count++; if(Test-Path "$env:ProgramFiles\Dell\CommandUpdate\dcu-cli.exe"){ Write-Host "Found"; Break }; Start-Sleep 1} until ($count -ge 10)


    $null = Reg.exe add "HKLM\SOFTWARE\DELL\UpdateService\Clients\CommandUpdate\Preferences\CFG" /v "ShowSetupPopup" /t REG_DWORD /d "0" /f
    $null = Reg.exe add "HKLM\SOFTWARE\DELL\UpdateService\Clients\CommandUpdate\Preferences\Settings\AdvancedDriverRestore" /v "IsAdvancedDriverRestoreEnabled" /t REG_DWORD /d "0" /f
    $null = Reg.exe add "HKLM\SOFTWARE\DELL\UpdateService\Clients\CommandUpdate\Preferences\Settings\General" /v "UserConsentDefault" /t REG_DWORD /d "0" /f
    $null = Reg.exe add "HKLM\SOFTWARE\DELL\UpdateService\Clients\CommandUpdate\Preferences\Settings\General" /v "SuspendBitLocker" /t REG_DWORD /d "1" /f
    $null = Reg.exe add "HKLM\SOFTWARE\DELL\UpdateService\Clients\CommandUpdate\Preferences\Settings\Schedule" /v "ScheduleMode" /t REG_SZ /d "ManualUpdates" /f
    $null = Reg.exe add "HKLM\SOFTWARE\DELL\UpdateService\Service\UpdateScheduler" /v  "CurrentUpdateState" /t REG_SZ /d "WaitForScan" /f


    Start-Sleep 2

    Start-Process "$Env:ProgramW6432\Dell\CommandUpdate\dcu-cli.exe" -Args "/scan" -Wait -NoNewWindow

    Start-Sleep 4

    Start-Process "$Env:ProgramW6432\Dell\CommandUpdate\dcu-cli.exe" -Args "/applyupdates" -Wait -NoNewWindow

    PowerShell.exe -NoLogo
}