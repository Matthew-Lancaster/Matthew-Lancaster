VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "VB6_Code.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer3 
      Interval        =   10000
      Left            =   2790
      Top             =   2250
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2700
      Top             =   1395
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1845
      Top             =   1260
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FF$

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


Private Sub Form_Load()

FF$ = Clipboard.GetText

'MsgBox FF$

'FF$ = "..................."

Shell "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE " + Command$, vbNormalFocus

'Clipboard.Clear
'Clipboard.SetText FF$

If FF$ = Trim("") Then End

'End
End Sub

Private Sub Timer1_Timer()

'New Project

'If FindWindow("wndclass_desked_gsk", vbNullString) > 0 Then
If FindWindow(vbNullString, "New Project") > 0 Then
'MsgBox "New P"
Timer2.Enabled = True
Clipboard.Clear
Clipboard.SetText FF$
End If

End Sub

Private Sub Timer2_Timer()
Clipboard.Clear
Clipboard.SetText FF$
End
End Sub

Private Sub Timer3_Timer()
MsgBox "VB6 Not Started after 10 Sec Stop this Clipboard Loader"
Clipboard.Clear
Clipboard.SetText FF$
End
End Sub
