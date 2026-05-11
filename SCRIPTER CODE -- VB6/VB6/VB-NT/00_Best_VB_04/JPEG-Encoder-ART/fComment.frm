VERSION 5.00
Begin VB.Form fComment 
   Caption         =   "Text Comment"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFinish 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "OK"
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtComment 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "fComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_Comment As String



Public Property Get Comment() As String
    Comment = m_Comment
End Property
Public Property Let Comment(Value As String)
    m_Comment = Value
    txtComment.Text = m_Comment
End Property



Private Sub Form_Resize()
    Dim NewWidth As Long
    Dim NewHeight As Long

    NewWidth = Me.Width - 240
    NewHeight = Me.Height - 1095
    If NewWidth < 0 Then NewWidth = 0
    If NewHeight < 0 Then NewHeight = 0

    txtComment.Move 60, 120, NewWidth, NewHeight
    cmdFinish(0).Move txtComment.Left, txtComment.Top + txtComment.Height + 120
    cmdFinish(1).Move txtComment.Left + txtComment.Width - cmdFinish(1).Width, cmdFinish(0).Top
End Sub



Private Sub cmdFinish_Click(Index As Integer)
    If Index = 1 Then m_Comment = txtComment.Text
    Unload Me
End Sub
