VERSION 5.00
Begin VB.Form frmAuthor 
   BackColor       =   &H0080FF80&
   BorderStyle     =   0  'None
   Caption         =   "About the Author"
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "X"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.PictureBox picAuthor 
      AutoSize        =   -1  'True
      Height          =   2370
      Left            =   120
      Picture         =   "frmAuthor.frx":0000
      ScaleHeight     =   2310
      ScaleWidth      =   1965
      TabIndex        =   1
      Top             =   480
      Width           =   2025
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Caption         =   "About the Author"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6495
   End
   Begin VB.Label lblEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail: Abazabam@Yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      MousePointer    =   2  'Cross
      TabIndex        =   2
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label lblAbout 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   2280
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
End
Attribute VB_Name = "frmAuthor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim Y As Integer
    Call SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    lblAbout.Caption = "My name is VB Beginner. I am the worst programmer ever but I'm always bored so I thought I'd start a hobby. That's all I have to say."
    Me.AutoRedraw = True
    Me.DrawStyle = 6
    Me.DrawMode = 13
    Me.DrawWidth = 13
    Me.ScaleMode = 3

    Me.ScaleHeight = (256)
    For i = 0 To 255
        Me.Line (0, Y)-(Me.Width, Y + 1), (vbRed) - (i * (vbRed - vbGreen) / 255), BF
        Y = Y + 1
    Next i
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Enabled = True
End Sub

Private Sub lblEmail_Click()
    ShellExecute 0&, vbNullString, "mailto:Abazabam@Yahoo.com", vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call ReleaseCapture
        Call SendMessage(Me.hwnd, &HA1, 2, 0)
    End If
End Sub
