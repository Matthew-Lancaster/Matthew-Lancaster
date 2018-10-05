Dim strLINK, strEXE, WSHShell
strLINK = """C:\Program Files (X86)\Microsoft Visual Studio\VB98\LINK.EXE"""
strEXE = """C:\SCRIPTER\SCRIPTER CODE -- BAT\NET_SHARE\Multiple_Thread Port Scanner 02 CON\Multiple_Port_Scanner.EXE"""
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.Run strLINK & " /EDIT /SUBSYSTEM:CONSOLE " & strEXE
Set WSHShell = Nothing
WScript.Echo "Complete! __ Microsoft Visual Studio\VB98\LINK.EXE"
