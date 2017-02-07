Dim objWshShell, i
Set objWshShell = WScript.CreateObject("WScript.Shell")

If objWshShell.AppActivate("Super BuildÅ^SS3") Then
	WScript.Sleep 100
	objWshShell.sendKeys("{F2}{TAB}{TAB}^{END}{ENTER}")
	For i = 0 to WScript.Arguments.Count - 1
		objWshShell.Run "cmd.exe /c echo """ & WScript.Arguments(i) & """ | clip", 0, TRUE
		WScript.Sleep 100
		objWshShell.sendKeys("^V")
		WScript.Sleep 100
	Next
	objWshShell.sendKeys("{TAB}{ENTER}")
	Set objWshShell = Nothing
	WScript.Quit
End If

Set objWshShell = Nothing
WScript.Quit 1