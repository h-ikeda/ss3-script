Dim objWshShell, i, j
Set objWshShell = WScript.CreateObject("WScript.Shell")

If WScript.Arguments.Count > 0 And objWshShell.AppActivate("Super Build／SS3") Then
	WScript.Sleep 100
	objWshShell.sendKeys("%fv{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{ENTER}{TAB}{TAB}{TAB}{ENTER}")
	For i = 1 to 100
		objWshShell.Run "cmd.exe /c echo """ & WScript.Arguments(0) & """ | clip", 0, TRUE
		WScript.Sleep 100
		objWshShell.sendKeys("^v^a^c")
		WScript.Sleep 100
		If Trim(WScript.CreateObject("HtmlFile").ParentWindow.ClipboardData.GetData("text")) = """" & WScript.Arguments(0) & """" Then
			objWshShell.sendKeys("{ENTER}")
			WScript.Sleep 100
			For j = 1 to 1000
				WScript.Sleep 1
				If objWshShell.AppActivate("SS3データ → CSVファイル") Then
					WScript. Sleep 100
					objWshShell.sendKeys("{TAB}{ENTER}%c")
					Set objWshShell = Nothing
					WScript.Quit
				End If
				objWshShell.sendKeys("y")
			Next
		End If
	Next
End If

Set objWshShell = Nothing
WScript.Quit 1