VERSION 5.00
Begin VB.Form Yahoo_Login 
   AutoRedraw      =   -1  'True
   Caption         =   "CID Yahoo Login"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   ScaleHeight     =   1785
   ScaleWidth      =   3720
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Verify Password"
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
      TabIndex        =   4
      Top             =   1425
      Width           =   3720
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Voidman2004xx Ebay ID"
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
      TabIndex        =   3
      Top             =   1080
      Width           =   3720
   End
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
      Width           =   3705
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "BrainBashing - Chat ID"
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
      Width           =   3705
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "MattLancaster2000 - Groups ID"
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
      Width           =   3720
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

Private Sub Label2_Click()
yahoo = 2
Unload Yahoo_Login

End Sub
Private Sub Label3_Click()
yahoo = 3
Unload Yahoo_Login
End Sub

Private Sub Label4_Click()
yahoo = 4
Unload Yahoo_Login
End Sub

Private Sub Label5_Click()
yahoo = 5
Unload Yahoo_Login
End Sub



