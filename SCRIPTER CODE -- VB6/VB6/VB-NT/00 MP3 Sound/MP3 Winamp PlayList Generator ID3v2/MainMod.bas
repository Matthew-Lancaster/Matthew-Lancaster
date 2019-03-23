Attribute VB_Name = "MainMod"
Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Rrt2$
Public B1$
Public c1 As String * 400
    
Public Const DRIVE_CDROM = 5
Public Const DRIVE_FIXED = 3
Public Const DRIVE_RAMDISK = 6
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_REMOVABLE = 2

Public Power

Sub Main()
'Call ScanPath.StartUp
Call ScanPath.Jeepers
'Call ScanPath.cmdScan_Click
End Sub

Sub junk()

F1 = FreeFile
Open App.Path + "CRC Such.txt" For Input As #F1
Do
Line Input #F1, F2$
wdm3$ = wdm3$ + F2$ + vbCrLf
If InStr(F2$, "*******" + ScanPath.txtPath) = 1 Then Exit Do
Loop Until EOF(F1)


If InStr(F2$, "*******" + ScanPath.txtPath) = 1 Then
wdm3$ = wdm3$ + wdm2$
End If

Do
Line Input #F1, F2$
wdm3$ = wdm3$ + F2$ + vbCrLf
If InStr(F2$, "*******") = 1 Then Exit Do
Loop Until EOF(F1)


Do
Line Input #F1, F2$
wdm3$ = wdm3$ + F2$ + vbCrLf
Loop Until EOF(F1)

Close #F1

Open App.Path + "CRC Such.txt" For Output As #F1
Print #F1, wdm3$
Close #F1

End
'Timer1.Enabled = True

End Sub
