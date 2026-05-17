VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

If Dir("C:\TEMP", vbDirectory) = "" Then
    MkDir ("C:\TEMP")
End If
    
On Error Resume Next

FR1 = FreeFile
Open "C:\TEMP\SYSTEM START TIME RECORDER.TXT" For Output As #FR1
    Print #FR1, Now
Close #FR1

If Err.Number > 0 Then
    'Me.Show
    DoEvents
    Me.WindowState = vbNormal
    DoEvents
    
    MsgBox "ERROR MAKE SYSTEM START TIME" + vbCrLf + "C:\TEMP\SYSTEM START TIME RECORDER.TXT" + Format(Now, "DDD DD MMM YYYY HH-MM-SS"), vbMsgBoxSetForeground
    
End If

Unload Me

End Sub
