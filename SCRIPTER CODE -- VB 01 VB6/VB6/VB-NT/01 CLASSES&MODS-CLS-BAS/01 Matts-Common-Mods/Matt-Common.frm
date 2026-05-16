VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Email Back"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   600
      Top             =   780
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public KF1, KF2, EXE2, Key


Sub Email_Bak_Main()


End Sub


Private Sub Form_Load()

'M = Move
'R = Recurse
'-M5 Best
'-M2 Fast

On Local Error Resume Next
f1 = FreeFile
Open App.Path + "\Count.txt" For Input As #f1
Line Input #f1, k1$
Line Input #f1, k2$
Close #f1

KF1 = Val(k1$)
KF1 = KF1 + 1
If KF1 = 7 Then KF1 = 0

KF2 = Val(k2$)
KF2 = KF2 + 1
If KF2 = 28 Then KF2 = 0

f1 = FreeFile
Open App.Path + "\Count.txt" For Output As #f1
Print #f1, Str$(KF1)
Print #f1, Str$(KF2)
Close #f1


KF1 = KF1 - 1
KF2 = KF2 - 1

'KF1 = 0
'KF2 = 0
'If KF1 = 0 Then
''Name "E:\Email_BackUp_Rar\Email_Backup-2.rar" As "E:\Email_BackUp_Rar\Email_Backup-Week.rar"
'Shell "C:\Program Files\WinRAR\winrar.exe U -cfg- -av -IBCK -r -m2 E:\Email_BackUp_Rar\Email_Backup-Week.rar C:\Email\*.*", vbNormalNoFocus
'End If
'
'If KF2 = 0 Then
''Name "E:\Email_BackUp_Rar\Email_Backup-2.rar" As "E:\Email_BackUp_Rar\Email_Backup-28-Days.rar"
'Shell "C:\Program Files\WinRAR\winrar.exe U -cfg- -av -IBCK -r -m2 E:\Email_BackUp_Rar\Email_Backup-28-Days.rar C:\Email\*.*", vbNormalNoFocus
'End If

Kill "E:\Email_BackUp_Rar\Email_Backup-2.rar"
Name "E:\Email_BackUp_Rar\Email_Backup-1.rar" As "E:\Email_BackUp_Rar\Email_Backup-2.rar"
Name "E:\Email_BackUp_Rar\Email_Backup-0.rar" As "E:\Email_BackUp_Rar\Email_Backup-1.rar"

Shell "C:\Program Files\WinRAR\winrar.exe A -cfg- -av -IBCK -r -m3 E:\Email_BackUp_Rar\Email_Backup-0.rar E:\Email\*.*", vbNormalNoFocus

'Form1.Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
'If Timer1.Interval <> 1000 Then Timer1.Interval = 1000
a$ = "Creating archive Email_Backup-0.rar"
ew = FindWindow(vbNullString, a$)
If ew > 0 Then Key = 1
If ew = 0 And Key = 1 Then

If KF1 = 0 Then
'Name "E:\Email_BackUp_Rar\Email_Backup-2.rar" As "E:\Email_BackUp_Rar\Email_Backup-Week.rar"
Set Fs22 = CreateObject("Scripting.FileSystemObject")
Fs22.CopyFile "E:\Email_BackUp_Rar\Email_Backup-0.rar", "E:\Email_BackUp_Rar\Email_Backup-Week.rar "

'Shell "C:\Program Files\WinRAR\winrar.exe U -cfg- -av -IBCK -r -m2 E:\Email_BackUp_Rar\Email_Backup-Week.rar C:\Email\*.*", vbNormalNoFocus
End If

If KF2 = 0 Then
'Name "E:\Email_BackUp_Rar\Email_Backup-2.rar" As "E:\Email_BackUp_Rar\Email_Backup-28-Days.rar"
Fs22.CopyFile "E:\Email_BackUp_Rar\Email_Backup-0.rar", "E:\Email_BackUp_Rar\Email_Backup-28-Days.rar "
'Shell "C:\Program Files\WinRAR\winrar.exe U -cfg- -av -IBCK -r -m2 E:\Email_BackUp_Rar\Email_Backup-28-Days.rar C:\Email\*.*", vbNormalNoFocus
End If
End If
End Sub
