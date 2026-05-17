VERSION 5.00
Begin VB.Form frmSplash 
   Caption         =   "Desktop Icon Tools"
   ClientHeight    =   975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2400
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   975
   ScaleWidth      =   2400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Show
Me.Refresh
End Sub

