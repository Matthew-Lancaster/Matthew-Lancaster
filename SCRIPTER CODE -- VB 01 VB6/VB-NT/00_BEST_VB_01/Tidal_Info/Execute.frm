VERSION 5.00
Begin VB.Form Execute 
   Caption         =   "Form1"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8220
   Icon            =   "Execute.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1890
      Top             =   1890
   End
   Begin VB.ListBox List1 
      Height          =   4935
      Left            =   0
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   8205
   End
   Begin VB.Menu mnu_edit 
      Caption         =   "Edit"
   End
   Begin VB.Menu mnu_refresh 
      Caption         =   "Refresh"
   End
   Begin VB.Menu mnu_execute 
      Caption         =   "Execute"
   End
End
Attribute VB_Name = "Execute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Call mnu_refresh_Click

End Sub

Private Sub List1_DblClick()
On Local Error Resume Next
Shell List1.List(List1.ListIndex), vbNormalFocus
Execute.Hide
If Err.Number > 0 Then MsgBox "Error Loading Program" + vbCrLf + List1.List(List1.ListIndex) + vbCrLf + Err.Description

End Sub

Private Sub mnu_edit_Click()
Shell "notepad.exe D:\# MY DOCS\CallerID\Executie.txt", vbNormalFocus
End Sub

Private Sub mnu_execute_Click()

For R = 0 To List1.ListCount - 1
    If List1.Selected(R) = True Then
        List1.Selected(R) = False
        On Local Error Resume Next
        Shell List1.List(R), vbNormalFocus
        
        If Err.Number > 0 Then MsgBox "Error ::- " + Err.Description
        On Local Error GoTo 0
    End If
Next

'Timer1.Enabled = True
Execute.Hide

End Sub

Private Sub mnu_refresh_Click()

fg5 = FreeFile
Open "D:\# MY DOCS\CallerID\Executie.txt" For Input As #fg5
List1.Clear
Do
Line Input #fg5, LK$
List1.AddItem LK$
Loop Until EOF(fg5)
Close fg5

End Sub

Private Sub Timer1_Timer()

Execute.Hide
Timer1.Enabled = False

End Sub
