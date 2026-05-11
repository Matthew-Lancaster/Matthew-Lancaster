VERSION 5.00
Begin VB.Form frmWinampCOMSample 
   Caption         =   "Winamp COM Sample"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Start Nullsoft Visulization"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Copy Winamp Playlist"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   4455
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   $"frmWinampCOMSample.frx":0000
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmWinampCOMSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oWinamp As New WINAMPCOMLib.Application

Private Sub Command1_Click()
Dim n As Long

List1.Clear

For n = 0 To oWinamp.PlayListCount - 1
    List1.AddItem oWinamp.SongTitle(n)
    a$ = oWinamp.CurrentSongFileName
Next n

End Sub

Private Sub Command2_Click()
    oWinamp.StartPlugIn "vis_w.dll"
End Sub

Private Sub List1_DblClick()
    oWinamp.PlayListPos = List1.ListIndex
    oWinamp.Play
End Sub
