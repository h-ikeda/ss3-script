Dim objWshShell
Set objWshShell = WScript.CreateObject("WScript.Shell")

If objWshShell.AppActivate("Super Build／SS3") Then
	WScript.Sleep 100
	objWshShell.sendKeys("%dc")
	Dim i, j
	For i = 1 to 100
		WScript.Sleep 1
		If objWshShell.AppActivate("解析と結果出力") Then
			objWshShell.sendKeys("10{ENTER}0{ENTER}")
			WScript.Sleep 100
			If objWshShell.AppActivate("出力先ファイルの確認") Then
				objWshShell.sendKeys("y")
			End If
			WScript.Sleep 1000
			Do
				objWshShell.sendKeys("c")
				WScript.Sleep 100
			Loop Until objWshShell.AppActivate("解析と結果出力")
			WScript.Sleep 100
			objWshShell.sendKeys("%c")
			Set objWshShell = Nothing
			WScript.Quit
		End If
	Next
End If

Set objWshShell = Nothing
WScript.Quit 1