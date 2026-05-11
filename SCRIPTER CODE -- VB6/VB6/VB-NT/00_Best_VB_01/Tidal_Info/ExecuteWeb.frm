VERSION 5.00
Begin VB.Form ExecuteWeb 
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8220
   Icon            =   "ExecuteWeb.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   15
      TabIndex        =   1
      Top             =   4950
      Width           =   630
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
Attribute VB_Name = "ExecuteWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Call mnu_refresh_Click

End Sub

Private Sub List1_DblClick()

'Shell List1.List(List1.ListIndex), vbNormalFocus
'ExecuteWeb.Hide
        cc$ = "C:\Program Files\Internet Explorer\IEXPLORE.EXE "
        abc$ = cc$ + ExecuteWeb.List1.List(List1.ListIndex)
        Shell (abc$), vbNormalFocus
        

End Sub

Private Sub mnu_edit_Click()
Shell "notepad.exe D:\# MY DOCS\CallerID\ExecutieWeb.txt", vbNormalFocus
End Sub

Private Sub mnu_execute_Click()

For R = 0 To List1.ListCount - 1
    If List1.Selected(R) = True Then
        List1.Selected(R) = False
        On Local Error Resume Next
        'Shell List1.List(R), vbNormalFocus
        cc$ = "C:\Program Files\Internet Explorer\IEXPLORE.EXE "
        abc$ = cc$ + ExecuteWeb.List1.List(R)
        Tr = Shell(abc$)

        If Err.Number > 0 Then MsgBox "Error ::- " + Err.Description
        On Local Error GoTo 0
    End If
Next

'Timer1.Enabled = True
ExecuteWeb.Hide

End Sub

Private Sub mnu_refresh_Click()
If Dir$("D:\# MY DOCS\CallerID\ExecutieWeb.txt") = "" Then
    fg5 = FreeFile
    Open "D:\# MY DOCS\CallerID\ExecutieWeb.txt" For Output As #fg5
    Close #fg5
End If
fg5 = FreeFile
Open "D:\# MY DOCS\CallerID\ExecutieWeb.txt" For Input As #fg5
List1.Clear
Do
If Not (EOF(fg5)) Then
    Line Input #fg5, lk$
    If Trim(lk$) <> "" Then List1.AddItem lk$
End If
Loop Until EOF(fg5)
Close fg5

Label1.Caption = Str(List1.ListCount - 1)

End Sub

Private Sub Timer1_Timer()

ExecuteWeb.Hide
Timer1.Enabled = False

End Sub
