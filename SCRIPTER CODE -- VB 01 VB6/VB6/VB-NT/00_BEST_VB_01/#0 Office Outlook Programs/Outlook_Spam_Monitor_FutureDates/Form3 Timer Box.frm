VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   Caption         =   "OutLook Email Chk"
   ClientHeight    =   1395
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3885
   Icon            =   "Form3 Timer Box.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   3885
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "-----"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   1665
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3915
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Activate()
Call Form1.xy

End Sub

Private Sub Form_Load()
Call Form1.Timer1_Timer
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label1_Click()
Timer4 = 0
Call Form1.Timer1_Timer
End Sub
