VERSION 5.00
Begin VB.Form frmProgress 
   Caption         =   "Rendering To File..     [ESC] to Cancel"
   ClientHeight    =   615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5655
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   41
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   377
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picProg 
      AutoRedraw      =   -1  'True
      Height          =   540
      Left            =   30
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   368
      TabIndex        =   0
      Top             =   30
      Width           =   5580
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Dim T1 As Long, T2 As Long
Dim nLeft As Long, nTop As Long

'Position:
nLeft = frmParent.Left + ((frmParent.Width \ 2) - (Me.Width \ 2))
nTop = frmParent.Top + ((frmParent.Height \ 2) - (Me.Height \ 2))
Me.Move nLeft, nTop
DoEvents

'Calculate to file:
T1 = GetTickCount
CalculateToFile Me.Tag, True
T2 = GetTickCount

If ((T2 - T1) > 1000) Then
    Win32.PlaySound " ", 0, 0
End If

'Turn off file output option
Frac(Me.Tag).Calc2File = False

Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyEscape Then
    Frac(Me.Tag).Flag_Stop = True
End If

End Sub
