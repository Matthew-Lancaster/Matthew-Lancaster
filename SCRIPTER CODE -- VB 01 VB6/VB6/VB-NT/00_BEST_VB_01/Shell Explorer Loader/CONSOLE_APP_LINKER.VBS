Dim strLINK, strEXE, WSHShell
strLINK = """C:\Program Files (X86)\Microsoft Visual Studio\VB98\LINK.EXE"""
strEXE = """D:\VB6\VB-NT\00_Best_VB_01\Shell Explorer LoaderShell Explorer Loader.EXE"""
Set WSHShell = CreateObject("WScript.Shell")
WSHShell.Run strLINK & " /EDIT /SUBSYSTEM:CONSOLE " & strEXE
Set WSHShell = Nothing
WScript.Echo "Complete! __ Microsoft Visual Studio\VB98\LINK.EXE"
