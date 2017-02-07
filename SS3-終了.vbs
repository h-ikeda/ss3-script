Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

If objWshShell.AppActivate("Super BuildÅ^SS3") Then
	WScript.Sleep 100
	objWshShell.sendKeys("%fx")
	WScript.Sleep 100
	objWshShell.sendKeys("{ENTER}")
	Dim i
	For i = 1 to 1000
		If objWshShell.AppActivate("Super BuildÅ^SS3") Then
			WScript.Sleep 1
		Else
			Set objWshShell = Nothing
			WScript.Quit
		End If
	Next
End If

Set objWshShell = Nothing
WScript.Quit 1