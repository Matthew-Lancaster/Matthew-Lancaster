VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "GRAB A SMS FILE"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   396
   ClientWidth     =   8748
   LinkTopic       =   "Form2"
   ScaleHeight     =   6360
   ScaleWidth      =   8748
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5580
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
Form1.RTB1 = II
DoEvents

EDITED = False

End Sub
