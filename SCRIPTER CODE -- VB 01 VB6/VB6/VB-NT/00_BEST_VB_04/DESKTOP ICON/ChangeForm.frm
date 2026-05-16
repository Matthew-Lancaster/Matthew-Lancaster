VERSION 5.00
Begin VB.Form ChangeForm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   533
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Screen Resolution Change Detected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   7725
   End
End
Attribute VB_Name = "ChangeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

ChangeForm.Width = changeWidth * Screen.TwipsPerPixelX
ChangeForm.Height = changeHeight * Screen.TwipsPerPixelY
ChangeForm.Label1.Move (ChangeForm.ScaleWidth \ 2) - (ChangeForm.Label1.Width \ 2), (ChangeForm.ScaleHeight \ 2) - (ChangeForm.Label1.Height \ 2)
ChangeForm.Show
ChangeForm.Refresh

End Sub
