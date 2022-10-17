For $i=1 to 10
	WinSetState("[CLASS:ConsoleWindowClass]", "", @SW_HIDE)
	Sleep(100)
Next

sleep(3000)

For $i=1 to 10
	WinSetState("[CLASS:ConsoleWindowClass]", "", @SW_SHOW)
	Sleep(100)
Next