VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   1920
      TabIndex        =   1
      Top             =   1680
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   600
      TabIndex        =   0
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Call GROUP_IN_500_JPG_MEDIA_PLAYER

End Sub
