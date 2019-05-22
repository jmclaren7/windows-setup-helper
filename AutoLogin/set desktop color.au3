RegWrite("HKEY_CURRENT_USER\Control Panel\Desktop","WallPaper","REG_SZ","")
RegWrite("HKEY_CURRENT_USER\Control Panel\Colors","Background","REG_SZ","192 16 16")
_WinAPI_SetSysColors($COLOR_BACKGROUND, 0x1010C0)