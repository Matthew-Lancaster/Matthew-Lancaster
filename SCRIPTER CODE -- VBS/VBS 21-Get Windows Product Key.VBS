'Location -- D:\03_MICROSOFT\#_OPPO_SYS\#_OP_SYS_-_TOOL_KITT\Windows 7 Product Key VBS
'Location -- D:\03_MICROSOFT\#_OPPO_SYS\#_OP_SYS_-_TOOL_KITT\Windows 7 Product Key VBS\VBS 21-Get Windows Product Key.VBS

IF GET_WINDOWS_VERSION=>6 THEN 
	MsgBox "Get Windows Product Key __ Not Work OS-VER Above Windows 7",vbMsgBoxSetForeground 
	WScript.Quit(0)
end if

Set WshShell = CreateObject("WScript.Shell")
MsgBox ConvertToKey(WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId")),vbMsgBoxSetForeground 


'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId
'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\12.0\Registration\{90120000-0021-0000-0000-0000000FF1CE}
'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\12.0\Registration\{90120000-0030-0000-0000-0000000FF1CE}
'HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\12.0\Registration\{90120000-0021-0000-0000-0000000FF1CE}
'HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office\12.0\Registration\{90120000-0030-0000-0000-0000000FF1CE}


Function ConvertToKey(Key)
Const KeyOffset = 52
i = 28
Chars = "BCDFGHJKMPQRTVWXY2346789"
Do
Cur = 0
x = 14
Do
Cur = Cur * 256
Cur = Key(x + KeyOffset) + Cur
Key(x + KeyOffset) = (Cur \ 24) And 255
Cur = Cur Mod 24
x = x -1
Loop While x >= 0
i = i -1
KeyOutput = Mid(Chars, Cur + 1, 1) & KeyOutput
If (((29 - i) Mod 6) = 0) And (i <> -1) Then
i = i -1
KeyOutput = "-" & KeyOutput
End If
Loop While i >= 0
ConvertToKey = KeyOutput
End Function

'--------------------------------------------------------
Function GET_WINDOWS_VERSION
'--------------------------------------------------------
	GET_WINDOWS_VERSION=0
	Set SystemSet = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 
	for each System in SystemSet 
	' WScript.Echo System.Caption 
	' WScript.Echo System.Manufacturer 
	' WScript.Echo System.BuildType 
	' WScript.Echo " Version: " + System.Version 
	' WScript.Echo " Locale: " + System.Locale 
	' WScript.Echo " Windows Directory: " + System.WindowsDirectory 
	' WScript.Echo " Total memory: " + System.TotalVisibleMemorySize + " bytes" 
	' WScript.Echo " Serial Number: " + System.SerialNumber 
	' Wscript.Echo "" 
	Version_OS = CINT(MID(System.Version,1,INSTR(System.Version,".")-1))
	'WINDOWS 7 = 6.1.7601
	GET_WINDOWS_VERSION=Version_OS
	next 
END Function