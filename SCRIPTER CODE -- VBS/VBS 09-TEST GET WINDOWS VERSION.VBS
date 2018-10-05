

CALL WINDOWS_VERSION 

SUB WINDOWS_VERSION 
	Set SystemSet = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 
	for each System in SystemSet 
	' WScript.Echo System.Caption 
	' WScript.Echo System.Manufacturer 
	' WScript.Echo System.BuildType 
	 WScript.Echo " Version: " + System.Version 
	' WScript.Echo " Locale: " + System.Locale 
	' WScript.Echo " Windows Directory: " + System.WindowsDirectory 
	' WScript.Echo " Total memory: " + System.TotalVisibleMemorySize + " bytes" 
	' WScript.Echo " Serial Number: " + System.SerialNumber 
	' Wscript.Echo "" 
		Version_OS = System.Version 
	next 
END SUB