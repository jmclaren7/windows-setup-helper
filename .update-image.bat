@ECHO OFF
set image-root=H:\Windows 10 Images\21H2
set additions1=H:\Windows 10 Images\Additions

del "%image-root%\Auto-saved*.xml"
del "%image-root%\NTLite.log"

robocopy "%~dp0\" "%image-root%\sources\$OEM$\$$\IT" /mir /xd .git
robocopy "%additions1%" "%image-root%" /e /xx

pause
