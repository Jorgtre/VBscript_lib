


'set java = new program
'java.setAssociatedProcesses(Array("iexplore.exe", "firefox.exe"))
'java.killAssociatedProcesses()
'java.uninstallByMatchingDisplayName("Java ")
'java.installByCommand("msiexec /I java.msi TRANSFORMS=transform.mst /passive /norestart")






Class program
	' *** Version 0.01.01 *** '
	Private p_oShell
	Private p_MSI
	Private p_associatedProcesses
	Private p_WMI
	Private p_oReg
	
	Private strComputer
	Private HKEY_LOCAL_MACHINE
	
	Public sub Class_Initialize
		p_MSI = "NULL"
		strComputer = "."
		HKEY_LOCAL_MACHINE = &H80000002
	
		Set p_oShell = CreateObject("Wscript.Shell")
		p_associatedProcesses = Array("")
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
	end Function
	
	Public Function uninstallByMatchingDisplayName(sw_displayName)
		Set guids = getGUIdsByDisplayName(sw_displayName)
		For Each guid in guids
			uninstallByCommand("msiexec /X " & guid & " /passive /norestart")
		Next
	End Function
	
	Public Function killAssociatedProcesses()
		For Each processName In p_associatedProcesses
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
	
	' *** Getters And Setters *** '
	
	Public Function setAssociatedProcesses(processes)
		p_associatedProcesses = processes
	End Function
	
End Class