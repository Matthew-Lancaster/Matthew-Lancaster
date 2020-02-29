VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1185
   LinkTopic       =   "Form1"
   ScaleHeight     =   510
   ScaleWidth      =   1185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   720
      Top             =   75
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'AlwaysOnTop (Me.hWnd)

End Sub

Private Sub Form_Unload(Cancel As Integer)
ATidalX.TimerVolLab.Enabled = True
End Sub

Private Sub Label1_Change()
'Form1.Show
End Sub

Private Sub Timer1_Timer()
'Form1.Hide
'Unload Form1
'form1.label1
End Sub
