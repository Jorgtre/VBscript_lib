Class fileSystem
	Private f_oShell
	Private f_fso
	
	Public Sub Class_Initialize()
		Set f_oShell = CreateObject("Wscript.Shell")
		Set f_fso = CreateObject("Scripting.FileSystemObject")
	End Sub
	
	Public Function createFolder(folderPath)
		path = ""
		For Each chunk In Split(folderPath, "\")
			path = path & chunk & "\"
			If Not(folderExists(path)) Then
				f_fso.CreateFolder(path)
			End If
		Next
	End Function
	
	Public Function folderExists(folderPath)
		folderExists = f_fso.FolderExists(folderPath)
	End Function
	
	Public Function copyFile(location, destination)
		f_fso.CopyFile location, destination, true
	End Function
End Class
