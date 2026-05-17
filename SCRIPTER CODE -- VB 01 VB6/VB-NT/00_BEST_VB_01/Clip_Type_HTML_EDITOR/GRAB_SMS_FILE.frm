VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "GRAB A SMS FILE"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   8745
   LinkTopic       =   "Form2"
   ScaleHeight     =   6360
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   5745
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   8415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub File1_Click()
FILESMSGET = File1.Path + "\" + File1.FileName
Me.Hide

FR = FreeFile
Open FILESMSGET For Input As #FR
II = Input(LOF(FR), FR)
Close FR
DoEvents
Form1.Text1 = II
DoEvents

EDITED = False

End Sub
