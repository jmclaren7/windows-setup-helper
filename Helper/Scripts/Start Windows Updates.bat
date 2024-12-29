@echo off
REM This script will start an automatic check and install of Windows updates through the update GUI

usoclient StartInteractiveScan ScanInstallWait
explorer.exe ms-settings:windowsupdate

REM Other related commands, not working or not needed
REM explorer.exe ms-settings:windowsupdate-action
REM (New-Object -ComObject Microsoft.Update.AutoUpdate).DetectNow()
REM $updateSession = new-object -com "Microsoft.Update.Session"; $updates=$updateSession.CreateupdateSearcher().Search($criteria).Updates
REM wuauclt /reportnow
REM wuauclt /UpdateNow
