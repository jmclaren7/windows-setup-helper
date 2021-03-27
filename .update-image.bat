set image-root=H:\Windows 10 Images\2009

del "%image-root%\Auto-saved*.xml"
del "%image-root%\NTLite.log"

robocopy "%~dp0\" "%image-root%\sources\$OEM$\$$\IT" /mir /xd .git

pause
