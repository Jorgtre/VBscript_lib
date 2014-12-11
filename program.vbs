
' *** Example *** '

'set java = new program
'java.killAssociatedProcesses(Array("iexplore.exe", "firefox.exe"))
'java.uninstallByMatchingDisplayName("Java ")
'java.installByCommand("msiexec /I java.msi TRANSFORMS=transform.mst /passive /norestart")
'java.installUpdate("JavaUpdate.msp")






Class program
	' *** Version 0.01.02 *** '
	Private p_oShell
	Private p_WMI
	Private p_oReg
	
	Private strComputer
	Private HKEY_LOCAL_MACHINE
	
	Public sub Class_Initialize
		strComputer = "."
		HKEY_LOCAL_MACHINE = &H80000002
	
		Set p_oShell = CreateObject("Wscript.Shell")
		Set p_WMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
		Set p_oReg = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
		
	End Sub
	
	' *** Public Methods *** '
	Public Function installWithMSI(msiName)
		p_oShell.Run "msiexec /i " & msiName & " /passive /norestart ALLUSERS=2", 1, True
	End Function
	
	Public Function installByCommand(command)
		p_oShell.Run command, 1, True
	End Function
	
	Public Function uninstallByCommand(command)
		p_oShell.Run command, 1, True
	End Function
	
	Public Function installUpdate(mspFile)
		p_oShell.Run "msiexec /update " & mspFile & " /passive /norestart", 1, True
	End Function
	
	Public Function uninstallByMatchingDisplayName(sw_displayName)
		Set guids = getGUIdsByDisplayName(sw_displayName)
		For Each guid in guids
			uninstallByCommand("msiexec /X " & guid & " /passive /norestart")
		Next
	End Function
	
	Public Function killAssociatedProcesses(associatedProcesses)
		For Each processName In associatedProcesses
			Set processInstances = p_WMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name='" & processName & "'")  
			For Each processInstance In processInstances
				p_oShell.Run "cmd.exe /c taskkill /IM " & processInstance.name & " /f", 0, True
			Next
		Next
	End Function
	
	' *** Private Methods *** '	
	Private Function getGUIdsByDisplayName(software)
		Set guidList = CreateObject("System.Collections.ArrayList")
		regUninstallArray = Array("software\microsoft\windows\currentversion\uninstall", _
								  "software\Wow6432Node\microsoft\windows\currentversion\uninstall")
		For Each regPath In regUninstallArray
			p_oReg.EnumKey HKEY_LOCAL_MACHINE, regPath, GUIDs
			For Each GUID In GUIDs
			On Error Resume next
				displayName = p_oShell.RegRead("HKLM\" & regPath & "\" & GUID & "\DisplayName")
				if InStr( Chr(34) & displayName & Chr(34), software) Then
					guidList.add GUID
				End If
			Next
		Next
		Set getGUIdsByDisplayName = guidList
	End Function
	
End Class
