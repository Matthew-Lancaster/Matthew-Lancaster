VERSION 5.00
Begin VB.Form Yahoo_Login 
   AutoRedraw      =   -1  'True
   Caption         =   "CID Yahoo Login"
   ClientHeight    =   1020
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   2565
   LinkTopic       =   "Form1"
   ScaleHeight     =   1020
   ScaleWidth      =   2565
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "MattLan2008"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   2580
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "******"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   2580
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "MattLancaster2000"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2580
   End
End
Attribute VB_Name = "Yahoo_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Yahoo_Login.SetFocus
Yahoo_Login.Show
AlwaysOnTop (Me.hWnd)

End Sub

Private Sub Form_Unload(Cancel As Integer)
If yahoo = 0 Then yahoo = -1
End Sub

Private Sub Label1_Click()
yahoo = 1
Unload Yahoo_Login
End Sub


Private Sub Label3_Click()
yahoo = 3
Unload Yahoo_Login
End Sub

