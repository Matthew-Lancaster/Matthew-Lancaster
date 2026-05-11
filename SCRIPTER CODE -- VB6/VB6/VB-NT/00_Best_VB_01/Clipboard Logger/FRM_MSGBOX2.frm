VERSION 5.00
Begin VB.Form FRM_MSGBOX2 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Clipboard Logger Error and Info Messages"
   ClientHeight    =   4690
   ClientLeft      =   110
   ClientTop       =   740
   ClientWidth     =   11590
   LinkTopic       =   "Form2"
   ScaleHeight     =   4690
   ScaleWidth      =   11590
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1234"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11560
   End
   Begin VB.Menu MNU_SHOW_ERROR_SCRIPT_LIST_HISTROY 
      Caption         =   "SHOW ERROR SCRIPT LIST HISTORY"
   End
End
Attribute VB_Name = "FRM_MSGBOX2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MNU_SHOW_ERROR_SCRIPT_LIST_HISTROY_Click()
    
    FRM_MSGBOX.Show

End Sub
