Dim objWshShell, i, j
Set objWshShell = WScript.CreateObject("WScript.Shell")

If WScript.Arguments.Count > 0 And objWshShell.AppActivate("Super Build／SS3") Then
	WScript.Sleep 100
	objWshShell.sendKeys("%fv{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}{ENTER}{TAB}{TAB}")
	For i = 1 to 15
		objWshShell.sendKeys("{ENTER}{TAB}{ENTER}{TAB}{TAB}{ENTER}{DOWN}{DOWN}{TAB}{TAB}{TAB}")
	Next
	objWshShell.sendKeys("{TAB}{ENTER}")
	For i = 1 to 100
		objWshShell.Run "cmd.exe /c echo """ & WScript.Arguments(0) & """ | clip", 0, TRUE
		WScript.Sleep 100
		objWshShell.sendKeys("^v^a^c")
		WScript.Sleep 100
		If Trim(WScript.CreateObject("HtmlFile").ParentWindow.ClipboardData.GetData("text")) = """" & WScript.Arguments(0) & """" Then
			objWshShell.sendKeys("{ENTER}")
			WScript.Sleep 100
			For j = 1 to 300
				WScript.Sleep 100
				If objWshShell.AppActivate("CSVファイル → SS3データ") Then
					WScript.Sleep 100
					objWshShell.sendKeys("{ENTER}%c")
					Set objWshShell = Nothing
					WScript.Quit
				End If
			Next
			Exit For
		End If
	Next
End If

Set objWshShell = Nothing
WScript.Quit 1