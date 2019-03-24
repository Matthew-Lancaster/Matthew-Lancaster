VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   516
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   788
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "6"
      Height          =   255
      Index           =   6
      Left            =   11280
      TabIndex        =   6
      Top             =   4920
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5"
      Height          =   255
      Index           =   5
      Left            =   10560
      TabIndex        =   5
      Top             =   6360
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      Height          =   255
      Index           =   4
      Left            =   8760
      TabIndex        =   4
      Top             =   6360
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3"
      Height          =   255
      Index           =   3
      Left            =   7680
      TabIndex        =   3
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2"
      Height          =   255
      Index           =   2
      Left            =   5520
      TabIndex        =   2
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   1
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   5520
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type vector2D
    X As Single
    Y As Single
End Type



Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


If Button = 1 Then
    Command1(Index).Top = Y / 15 + Command1(Index).Top
    Command1(Index).Left = X / 15 + Command1(Index).Left
    drawline
ElseIf Button = 2 Then
    Select Case Index
        Case 0
            Command1(0).Top = Y / 15 + Command1(0).Top
            Command1(0).Left = X / 15 + Command1(0).Left
            Command1(1).Top = Y / 15 + Command1(1).Top
            Command1(1).Left = X / 15 + Command1(1).Left
        Case 3
            Command1(2).Top = Y / 15 + Command1(2).Top
            Command1(2).Left = X / 15 + Command1(2).Left
            Command1(3).Top = Y / 15 + Command1(3).Top
            Command1(3).Left = X / 15 + Command1(3).Left
            Command1(4).Top = Y / 15 + Command1(4).Top
            Command1(4).Left = X / 15 + Command1(4).Left
        Case 6
            Command1(6).Top = Y / 15 + Command1(6).Top
            Command1(6).Left = X / 15 + Command1(6).Left
            Command1(5).Top = Y / 15 + Command1(5).Top
            Command1(5).Left = X / 15 + Command1(5).Left
    End Select
    drawline
End If

End Sub

Private Function drawline()
Dim p2 As vector2D
Dim p(100) As vector2D
Form1.Cls

For T = 0 To 6
    p(T).X = Command1(T).Left
    p(T).Y = Command1(T).Top
Next T
Form1.PSet (p(0).X, p(0).Y)
For v = 0 To 1

    
    For T = 0 To 1 Step 0.1
        p2 = IntBezier2D(p(0 + v * 3), p(1 + v * 3), p(2 + v * 3), p(3 + v * 3), T)
        Form1.Line -(p2.X, p2.Y)
        Form1.Circle (p2.X, p2.Y), 5
    Next T
Next v


End Function




Private Function IntBezier2D(Vec0 As vector2D, Vec1 As vector2D, Vec2 As vector2D, Vec3 As vector2D, ByVal a As Single) As vector2D
IntBezier2D.X = ((1 - a) ^ 3) * Vec0.X + (3 * a) * ((1 - a) ^ 2) * Vec1.X + (3 * a * a) * (1 - a) * Vec2.X + (a ^ 3) * Vec3.X
IntBezier2D.Y = ((1 - a) ^ 3) * Vec0.Y + (3 * a) * ((1 - a) ^ 2) * Vec1.Y + (3 * a * a) * (1 - a) * Vec2.Y + (a ^ 3) * Vec3.Y
End Function


Private Sub Form_Load()
Me.Show
drawline
End Sub
