Dim FSO
Set FSO = CreateObject("SCRIPTING.FILESYSTEMOBJECT")
FSO.CopyFile "D:\VB6-EXE\VB-NT\00_Send_To\Send To Text List of Files & Sub Folders IRFAN\#0 Send To Text List of Files and Sub Folders IRFAN.EXE","D:\VB6\VB-NT\00_Send_To\Send To Text List of Files & Sub Folders IRFAN\"
Set UAC = CreateObject("SHELL.APPLICATION")
UAC.ShellExecute "D:\VB6\VB-NT\00_Send_To\Send To Text List of Files & Sub Folders IRFAN\#0 Send To Text List of Files and Sub Folders IRFAN", "", "", "RUNAS", 1

