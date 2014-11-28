Class computerInfo
	' *** Version 0.00.02 *** '
	Private c_oShell
	Private c_oWMI
	Private strComputer
	
	Public Sub Class_initialize()
		strComputer = "."
		Set c_oShell = CreateObject("Wscript.Shell")
		Set oWMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	End Sub
	
	Function getLoggedOnUsersByDomain(d)
		query = "SELECT * FROM Win32_Process"
		Set processCollection = oWMI.ExecQuery(query)
		Set userList = CreateObject("System.Collections.ArrayList")
		For Each process In processCollection
			process.GetOwner username, domain
			If domain = d Then
				userList.Add username
			End if
		Next
		Set getLoggedOnUsersByDomain = userList
	End Function
	
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
