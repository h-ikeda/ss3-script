Dim objWshShell, i, j
Set objWshShell = WScript.CreateObject("WScript.Shell")

For i = 1 to 1000
	If objWshShell.AppActivate("Super BuildÅ^SS3") Then
		For j = 1 to 100
			objWshShell.sendKeys("y{F2}")
			WScript.Sleep 100
			If objWshShell.AppActivate("ÉÅÉÇÇÃï“èW") Then
				WScript.Sleep 100
				objWshShell.sendKeys("{Enter}")
				objWshShell = Nothing
				WScript.Quit
			End If
		Next
	Else
		WScript.Sleep 1
	End If
Next

objWshShell = Nothing
WScript.Quit 1
