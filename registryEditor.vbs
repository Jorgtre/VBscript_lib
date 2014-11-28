'Set redit = new registryEditor

' *** Create Value *** '
'redit.createValue "HKEY_CURRENT_USER\Test", "ValueName", "REG_SZ", "value"

' *** Create many values by setting a root  *** '
'redit.setRoot "HKEY_CURRENT_USER"
'redit.createValue "Test", "ValueName1", "REG_SZ", "value"
'redit.createValue "Test", "ValueName2", "REG_SZ", "value"
'redit.createValue "Test", "ValueName3", "REG_SZ", "value"

' *** Get SubKeys *** '
'redit.setRoot "HKEY_CURRENT_USER"
'subKeys = redit.getSubKeys("Console")

' *** Get Single value *** '
'redit.setRoot("HKEY_CURRENT_USER")
'redit.getValue("Test", "ValueName")

' *** Get Multiple values *** '
'redit.setRoot "HKEY_CURRENT_USER"
'subValues = redit.getSubValues("Test")


Class registryEditor
	' Version 0.00.10 '
	
	private h_root
	private h_fileLocation
	private h_oShell
	private h_oReg
	private isLoaded
	
	private strComputer
	private HKEY_CLASSES_ROOT
	private HKEY_CURRENT_USER
	private HKEY_LOCAL_MACHINE
	private HKEY_USERS
	private HKEY_CURRENT_CONFIG
	
	' *** Constructor *** '
	Public Sub Class_Initialize()
		strComputer = "."
		
		h_root = ""
		h_fileLocation = ""
		
		HKEY_CLASSES_ROOT 	= &H80000000
		HKEY_CURRENT_USER 	= &H80000001
		HKEY_LOCAL_MACHINE 	= &H80000002
		HKEY_USERS 		= &H80000003
		HKEY_CURRENT_CONFIG     = &H80000005
	
		Set h_oShell = CreateObject("WScript.Shell")
		Set h_oReg = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
	End Sub
	
	' *** Destructor *** '
	Public Sub Class_Terminate()
		
	End Sub
	
	' *** Public Methods *** '
	Public Function createKey(key)
		createNonExistingKeys(h_root & "\" & key)
	End Function
	
	Public Function createValue(key, subKey, keyType, value)
		h_oShell.RegWrite fixPath(key) & "\" & subkey, value
	End Function
	
	Public Function getSubKeys(regPath)
		arr = split(fixPath(regPath), "\")
		HKEY = arr(0)
		arr(0) = ""
		h_oReg.EnumKey getHKEY_VALUE(HKEY), remFirstChar(Join(arr, "\")), keys
		getSubKeys = keys
	End Function
	
	Public Function getSubValues(regPath)
		arr = split(fixPath(regPath), "\")
		HKEY = arr(0)
		arr(0) = ""
		h_oReg.EnumValues getHKEY_VALUE(HKEY), remFirstChar(Join(arr, "\")), values
		getSubValues = values
	End Function
	
	Public Function getData(regPath, value)
		getData = h_oShell.RegRead(fixPath(regPath) & "\" & value)
	End Function
	
	Public Function loadHive()
		h_oShell.Run "reg load " & encapChr34(h_root) & " " & encapChr34(h_fileLocation), 0, true
		isLoaded = true
	End Function
	
	Public Function unLoadHive()
		h_oShell.Run "reg unload " & encapChr34(h_root), 0, true
		isLoaded = false
	End Function
	
	Public Function keyExists(key)
		on Error Resume Next
		h_oShell.RegRead(key)
		if Err.Number <> 0 Then
			keyExists = False
		Else
			keyExists = True
		End if
	End Function
	
	Public Function valueExists(key, value)
		values = getSubValues(key)
		For Each val In values
			if val = value Then
				valueExists = true
				Exit Function
			End if
		Next
		valueExists = false
	End Function
	
	' *** Private Methods ***
	Private Function createNonExistingKeys(key)
		Dim str
		
		For Each s in split(key, "\")
			str = str & s & "\"
			if Not (keyExists(str)) Then
				h_oShell.Run "reg ADD " & str & " /f", 0, true
			End if
		Next
	End Function
	
	Private Function remLastChar(str)
		remLastChar = mid(str, 1, (Len(str)-1))
	End Function
	
	Private Function getStringKeyRoot(key)
		keys = Split(key, "\")
		getStringKeyRoot = keys(0)
	End Function
	
	Private Function remFirstChar(str)
		remFirstChar = mid(str, 2)
	End Function
	
	Private Function encapChr34(str)
		encapChr34 = Chr(34) & str & Chr(34)
	End Function
	
	Private Function getHKEY_VALUE(HKEY)
		HK = getStringKeyRoot(HKEY)
		
		HKEY_ENUM_VALUE = ""
		
	
		Select Case HK
			Case "HKEY_CURRENT_USER", "HKCU"
				HKEY_ENUM_VALUE = HKEY_CURRENT_USER
			Case "HKEY_LOCAL_MACHINE", "HKLM"
				HKEY_ENUM_VALUE = HKEY_LOCAL_MACHINE
			Case "HKEY_USERS"
				HKEY_ENUM_VALUE = HKEY_USERS
			Case "HKEY_CURRENT_CONFIG", "HKCC"
				HKEY_ENUM_VALUE = HKEY_CURRENT_CONFIG
			Case "HKEY_CLASSES_ROOT", "HKCR"
				HKEY_ENUM_VALUE = HKEY_CLASSES_ROOT
			Case Else
				HKEY_ENUM_VALUE = "undefined"
		End Select
			
		getHKEY_VALUE = HKEY_ENUM_VALUE
	End Function
	
	Private Function fixPath(path)
		if getHKEY_VALUE(path) = "undefined" Then
				fixPath = h_root & "\" & path
			Else
				fixPath = path
		End If
	End Function
	
	' *** Setters ***
	Public Function setRoot(name)
		h_root = name
	End Function
	
	Public Function setHiveFile(location)
		h_fileLocation = location
	End Function
	
	' *** Getters *** '
	Public Function getRoot()
		getRoot = h_root
	End Function
	
End Class