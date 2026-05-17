VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MP3 Tag Info"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   7605
   Icon            =   "MP3 Tags.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   2205
      Left            =   1170
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   630
      Width           =   4995
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1935
      Top             =   3555
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   7605
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************
'Windows API/Global Declarations for :Dr
'     ive Type Finder
'**************************************



'--------------End This Form Code Jeepers








Private Sub Form_Activate()

Exit Sub


Private Sub Form_Load()

Call Jeepers

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer1_Timer()
End
End Sub



'--------------End This Form Code Jeepers



