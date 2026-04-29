VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Spy_Form 
   Caption         =   "Spying Menus <Enumerating Menus>"
   ClientHeight    =   4275
   ClientLeft      =   2550
   ClientTop       =   2370
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   4935
   Begin MSComctlLib.TreeView Tree 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5953
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   4695
      Begin VB.CommandButton Command2 
         Caption         =   "&Expand All"
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   225
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Close"
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   225
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Bring Window to Top"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "When Menu action takes place, the relative window will be given focus, otherwise not"
         Top             =   290
         Value           =   1  'Checked
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Spy_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private size As New CResize

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Dim i As Integer
    For i = 1 To Tree.Nodes.Count
      Tree.Nodes(i).Expanded = True
   Next i

End Sub

Private Sub Form_Load()
    With size
        .hParam = Me.Height
        .wParam = Me.Width
        .Map Tree, RS_Height_Width
        .Map Frame1, RS_TopOnly
        .Map Frame1, RS_WidthOnly
        '.Map Check1, RS_LeftOnly
        .Map Command1, RS_LeftOnly
        .Map Command2, RS_LeftOnly
        
    End With
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub Form_Resize()
    size.rSize Me
    
End Sub

Private Sub Tree_DblClick()
    On Error GoTo hwndE
    
    Dim id As Long, l As Long
    If Check1.Value = vbChecked Then
        BringWindowToTop SpyHwnd
    End If
    With Tree.SelectedItem
        id = CLng(Right(.Key, Len(.Key) - 1)) - 15000
    End With
    l = PostMessage(SpyHwnd, WM_COMMAND, id, 0&)
    'Debug.Print l

    Exit Sub
hwndE:
    MsgBox "error"
    

End Sub
