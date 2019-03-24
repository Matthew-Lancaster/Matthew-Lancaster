VERSION 5.00
Object = "{DEEABA2C-9ED3-11D4-92CF-9E7BD6A8DC73}#1.0#0"; "MEDIAPLAYLIST.OCX"
Begin VB.Form frmTrans 
   BackColor       =   &H008080FF&
   BorderStyle     =   0  'None
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin MediaPlaylist.comFunc comFunc1 
      Height          =   195
      Left            =   3480
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   344
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TRANS"
      Height          =   375
      Index           =   4
      Left            =   2550
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O"
      Height          =   375
      Index           =   3
      Left            =   3720
      TabIndex        =   3
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "M"
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   2
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "D"
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   3240
      Width           =   375
   End
End
Attribute VB_Name = "frmTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Sub Command1_Click(index As Integer)
    Dim transFile As String
    Dim WinAmpFT As Boolean
    Me.Width = 6000
    Select Case index
        Case Is = 0: transFile = App.Path & "\trans.txt": Me.Caption = "Shape1":
        Case Is = 1: transFile = App.Path & "\trans1.txt": Me.Caption = "Shape2":
        Case Is = 2: transFile = App.Path & "\trans2.txt": Me.Caption = "Shape3":
        Case Is = 3: transFile = App.Path & "\trans3.txt": Me.Caption = "Shape4":
        Case Else: transFile = App.Path & "\region.txt": Me.Width = 8250:  WinAmpFT = True:
    End Select
    If WinAmpFT = False Then
        Unload frmSecond
        If comFunc1.FileExist(transFile) Then
            comFunc1.TransRestore Me.hWnd, 6000, 4500
            comFunc1.TransArea Me.hWnd, 6000, 4500, transFile, False  'true = dblSize
        Else
            comFunc1.TransRestore Me.hWnd, 6000, 4500
        End If
    Else
        If comFunc1.FileExist(transFile) Then
            comFunc1.TransRestore Me.hWnd, 8250, 3480
            comFunc1.TransRestore frmSecond.hWnd, 8250, 3480
            comFunc1.TransWinamp Me.hWnd, frmSecond.hWnd, transFile, True
            frmSecond.Move Me.Left, Me.Top + 3480, 8250, 3480
            frmSecond.Show
        End If
    End If
End Sub

Private Sub Form_DblClick()
    Unload frmSecond
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "Form Shaper"
    Me.Width = 6000
    Me.Height = 4500
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

