@echo off
REM Set the desktop background color to red for the current user

Call :SetBackground Red
Exit

:SetBackground
SET BR=%~1
SET BG=%~2
SET BB=%~3
IF "%~1"=="Red"    (SET BR=D1 & SET BG=34 & SET BB=38)
IF "%~1"=="Green"  (SET BR=00 & SET BG=B3 & SET BB=36)
IF "%~1"=="Blue"   (SET BR=00 & SET BG=8C & SET BB=FF)
IF "%~1"=="Orange" (SET BR=FF & SET BG=8C & SET BB=00)
IF "%~1"=="Blue"   (SET BR=22 & SET BG=33 & SET BB=44)
IF "%~1"=="Purple" (SET BR=22 & SET BG=33 & SET BB=44)
IF "%~1"=="Gray"   (SET BR=7F & SET BG=7F & SET BB=7F)
set BGPath=%ProgramData%\Helper
set BGFullPath=%BGPath%\backgroundpixel.bmp
mkdir %BGPath%
>backgroundpixel.tmp echo(42 4D 3A 00 00 00 00 00 00 00 36 00 00 00 28 00 00 00 01 00 00 00 01 00 00 00 01 00 18 00 00 00 00 00 00 00 00 00 12 0B 00 00 12 0B 00 00 00 00 00 00 00 00 00 00 %BB% %BG% %BR% 00
certutil -f -decodehex backgroundpixel.tmp %BGFullPath% >nul
del backgroundpixel.tmp
reg add "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System" /v Wallpaper /t REG_SZ /d "%BGFullPath%" /f
reg add "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System" /v WallpaperStyle /t REG_SZ /d 1 /f
taskkill /F /IM explorer.exe & start explorer
exit /B 0