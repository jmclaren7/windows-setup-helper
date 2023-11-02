For $i=1 to 10
	WinSetState(@comspec, "", @SW_HIDE)
	Sleep(100)
Next

exit
sleep(3000)

For $i=1 to 10
	WinSetState("[CLASS:ConsoleWindowClass]", "", @SW_SHOW)
	Sleep(100)
Next