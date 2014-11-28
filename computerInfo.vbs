Class computerInfo
	' *** Version 0.00.02 *** '
	Private c_oShell
	
	Public Sub Class_initialize()
		Set c_oShell = CreateObject("Wscript.Shell")
	End Sub
	
	Public Function getOSArchitecture()
		arch = c_oShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
		Select Case arch
		Case "AMD64"
			getOSArchitecture = "x64"
		Case "x86"
			getOSArchitecture = "x86"
		Case Else
			getOSArchitecture = "unknown"
		End Select
	End Function

End Class
