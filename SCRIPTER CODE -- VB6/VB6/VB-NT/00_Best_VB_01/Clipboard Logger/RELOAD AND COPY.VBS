Dim FSO
Set FSO = CreateObject("SCRIPTING.FILESYSTEMOBJECT")
FSO.CopyFile "D:\VB6-EXE\VB-NT\00_Best_VB_01\Clipboard Logger\ClipBoard Logger.EXE","D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\"
Set UAC = CreateObject("SHELL.APPLICATION")
UAC.ShellExecute "D:\VB6\VB-NT\00_Best_VB_01\Clipboard Logger\ClipBoard Logger", "", "", "RUNAS", 1

