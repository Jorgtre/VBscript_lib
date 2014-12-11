Class computerInfo
	' *** Version 0.00.05 *** '
	Private c_oShell
	Private c_oWMI
	Private strComputer
	
	Public Sub Class_initialize()
		strComputer = "."
		Set c_oShell = CreateObject("Wscript.Shell")
		Set c_oWMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	End Sub
	
	Public Function getLoggedInUsersByDomain(d)
		'Returns Dictionary
		query = "SELECT * FROM Win32_Process"
		Set processCollection = c_oWMI.ExecQuery(query)
		Set userDic = CreateObject("Scripting.Dictionary")
		For Each process In processCollection
			process.GetOwner username, domain
			If domain = d Then
				userDic(username) = 0
			End if
		Next
		Set getLoggedInUsersByDomain = userDic
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
