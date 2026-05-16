VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Joystick Example..."
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   2925
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      DrawWidth       =   5
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2355
      ScaleMode       =   0  'User
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3120
      Top             =   960
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   2160
      TabIndex        =   8
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   1560
      TabIndex        =   7
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   960
      TabIndex        =   6
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   5
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   2160
      TabIndex        =   4
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   1560
      TabIndex        =   3
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   2640
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim j As JOYINFO
Dim je As JOYCAPS
Const red = &HFF&
Const White = &HFFFFFF
Dim be As B
Private Sub Form_Load()
If IsWinNT Then
    joyGetDevCapsW 0, je, Len(je)
Else
    joyGetDevCapsA 0, je, Len(je)
End If
If je.wXmax > 0 Then Picture1.ScaleHeight = je.wXmax
If je.wYmax > 0 Then Picture1.ScaleWidth = je.wYmax
Picture1.ScaleLeft = 0
Picture1.ScaleTop = 0
End Sub

Private Sub Timer1_Timer()
joyGetPos 0, j
Picture1.DrawWidth = 100
Picture1.Cls
Picture1.PSet (j.X, j.Y)
GetBPress j.Buttons, be
For n = 0 To 7
    If be.B(n + 1) = True Then
    Label1(n).BackColor = red
    Else
    Label1(n).BackColor = White
    End If
Next n
End Sub
