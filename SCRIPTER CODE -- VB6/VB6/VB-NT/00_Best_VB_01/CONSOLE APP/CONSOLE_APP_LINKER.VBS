Dim strLINK, strEXE, WSHShell

strLINK = """C:\Program Files (X86)\Microsoft Visual Studio\VB98\LINK.EXE"""
' strEXE = """" & WScript.Arguments(0) & """"
strEXE = """" & "D:\VB6\VB-NT\00_Best_VB_01\CONSOLE APP\CONSOLE_APP.exe" & """"
Set WSHShell = CreateObject("WScript.Shell")
 
WSHShell.Run strLINK & " /EDIT /SUBSYSTEM:CONSOLE " & strEXE
 
Set WSHShell = Nothing
WScript.Echo "Complete!"
