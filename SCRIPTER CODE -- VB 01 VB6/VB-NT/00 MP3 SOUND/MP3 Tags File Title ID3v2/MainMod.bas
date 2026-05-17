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

f1 = FreeFile
Open App.Path + "CRC Such.txt" For Input As #f1
Do
Line Input #f1, f$
wdm3$ = wdm3$ + f$ + vbCrLf
If InStr(f$, "*******" + ScanPath.txtPath) = 1 Then Exit Do
Loop Until EOF(f1)


If InStr(f$, "*******" + ScanPath.txtPath) = 1 Then
wdm3$ = wdm3$ + wdm2$
End If

Do
Line Input #f1, f$
wdm3$ = wdm3$ + f$ + vbCrLf
If InStr(f$, "*******") = 1 Then Exit Do
Loop Until EOF(f1)


Do
Line Input #f1, f$
wdm3$ = wdm3$ + f$ + vbCrLf
Loop Until EOF(f1)

Close #f1

Open App.Path + "CRC Such.txt" For Output As #f1
Print #f1, wdm3$
Close #f1

End
'Timer1.Enabled = True

End Sub
