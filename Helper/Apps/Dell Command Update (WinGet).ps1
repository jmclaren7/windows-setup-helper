# This script will check to see if the computer is manufactured by Dell, download and install Dell Command Update, and attempt to install updates

If ($(Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer -like "*Dell*"){
    If (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){
        Write-Host "Not admin, exiting"
        Start-Sleep 3
        Exit
    }
    Write-Host "Waiting for user input to continue Dell command update"
    $Wscript_Shell = New-Object -ComObject "Wscript.Shell"
    $MsgBox = $Wscript_Shell.Popup("Dell command update will run automatically in 20 seconds, press 'ok' to run now or 'cancel' to abort.", 20, "Setup Helper", 1+32)
    If ($MsgBox -ne 1 -AND $MsgBox -ne -1) { 
        Exit 
    }

    Start-Process "winget" -Args "install --id Dell.CommandUpdate.Universal --exact --source winget --accept-source-agreements --disable-interactivity --silent --accept-package-agreements --force" -Wait -NoNewWindow

    $DCU_CLI = "$($env:ProgramFiles)\Dell\CommandUpdate\dcu-cli.exe"
    Write-Host 'Waiting for Dell Command Update to install...'
    do{
        $count++
        if(Test-Path $DCU_CLI){ 
            Write-Host "Found"
            Break 
        } 
        Start-Sleep 1
    } until ($count -ge 10)

    $null = New-ItemProperty -Path "HKLM:\SOFTWARE\DELL\UpdateService\Clients\CommandUpdate\Preferences\CFG" -Name "ShowSetupPopup" -Value 0 -PropertyType DWord -Force -ErrorAction SilentlyContinue
    $null = New-ItemProperty -Path "HKLM:\SOFTWARE\DELL\UpdateService\Clients\CommandUpdate\Preferences\Settings\AdvancedDriverRestore" -Name "IsAdvancedDriverRestoreEnabled" -Value 0 -PropertyType DWord -Force -ErrorAction SilentlyContinue
    $null = New-ItemProperty -Path "HKLM:\SOFTWARE\DELL\UpdateService\Clients\CommandUpdate\Preferences\Settings\General" -Name "UserConsentDefault" -Value 0 -PropertyType DWord -Force -ErrorAction SilentlyContinue
    $null = New-ItemProperty -Path "HKLM:\SOFTWARE\DELL\UpdateService\Clients\CommandUpdate\Preferences\Settings\General" -Name "SuspendBitLocker" -Value 1 -PropertyType DWord -Force -ErrorAction SilentlyContinue
    $null = New-ItemProperty -Path "HKLM:\SOFTWARE\DELL\UpdateService\Clients\CommandUpdate\Preferences\Settings\Schedule" -Name "ScheduleMode" -Value "ManualUpdates" -PropertyType String -Force -ErrorAction SilentlyContinue
    $null = New-ItemProperty -Path "HKLM:\SOFTWARE\DELL\UpdateService\Service\UpdateScheduler" -Name "CurrentUpdateState" -Value "WaitForScan" -PropertyType String -Force -ErrorAction SilentlyContinue
    
    

    Write-Host "Starting Dell Command Update scan and install..."
    Start-Sleep 2
    Start-Process $DCU_CLI -Args "/scan" -Wait -NoNewWindow
    Start-Sleep 4
    Start-Process $DCU_CLI -Args "/applyupdates" -Wait -NoNewWindow
}else{
    Write-Host "This computer is not manufactured by Dell..."
}

Write-Host "Exiting..."
Start-Sleep 3