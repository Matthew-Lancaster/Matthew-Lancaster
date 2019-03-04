Dim FSO
Set FSO = CreateObject("SCRIPTING.FILESYSTEMOBJECT")
On Error Resume Next
Do
    Err.Clear
    FC_1 = "D:\VB6-EXE\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.EXE"
    FC_2 = "D:\VB6\VB-NT\00_Best_VB_01\EliteSpy\EliteSpy.EXE"
    FN_1 = Mid(FC_1, InStrRev(FC_1, "\") + 1)
    FSO.CopyFile FC_1, FC_2
Loop Until Err.Number = 0 Or X_COUNT > 100
If X_COUNT > 100 Then
    MsgBox "ERROR COPY FILE RETRY COUNT 100 FOR VB UPDATE" + vbCrLf + FN_1
End If
Err.Clear
Set UAC = CreateObject("SHELL.APPLICATION")
UAC.ShellExecute FC_2, "", "", "RUNAS", 1
If Err.Number > 0 Then
MsgBox "ERROR LAUNCH VB PROGRAM FROM UPDATE" + vbCrLf + FN_1 + vbCrLf + Err.Description
End If
WScript.Quit 0

