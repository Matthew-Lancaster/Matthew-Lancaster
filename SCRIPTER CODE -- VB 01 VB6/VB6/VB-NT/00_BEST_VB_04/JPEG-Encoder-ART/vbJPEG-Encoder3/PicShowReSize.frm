VERSION 5.00
Begin VB.Form PicShowReSize 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame PicBack 
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   915
      TabIndex        =   0
      Top             =   750
      Width           =   3195
      Begin VB.Image ViewImage 
         Height          =   1155
         Left            =   435
         Top             =   435
         Width           =   2610
      End
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
   End
End
Attribute VB_Name = "PicShowReSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()

Call PicShow(FileName, PicShowReSize)

End Sub

Private Sub Form_Resize()

Call PicShow(FileName, PicShowReSize)

End Sub

Public Sub ViewImage_Click()

End Sub
