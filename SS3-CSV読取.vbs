Function CSV_getData(title, csv)
	Dim result, objFile
	Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(csv, 1)
	Do
		If objFile.AtEndOfStream Then
			WScript.Quit 1
		End If
		result = objFile.ReadLine
	Loop Until InStr(1, result, "<" & title & ">") = 1
	objFile.ReadLine
	Set result = WScript.CreateObject("System.Collections.ArrayList")
	Do
		Dim line
		line = objFile.ReadLine
		If line = "" Then
			Exit Do
		End If
		result.Add Split(line, ",")
	Loop Until objFile.AtEndOfStream
	CSV_getData = result.ToArray
End Function

Function CSV_getDataHeader(title, csv)
	Dim result, objFile
	Set objFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(csv, 1)
	Do
		If objFile.AtEndOfStream Then
			WScript.Echo "Reached the end of the file."
			WScript.Quit 1
		End If
		result = objFile.ReadLine
	Loop Until InStr(1, result, "<" & title & ">") = 1
	CSV_getDataHeader = Split(objFile.ReadLine, ",")
End Function