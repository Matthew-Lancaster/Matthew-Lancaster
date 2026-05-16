VERSION 5.00
Begin VB.Form Form_Menu 
   Caption         =   "Tidal Menu"
   ClientHeight    =   420
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   2115
   LinkTopic       =   "Form1"
   ScaleHeight     =   420
   ScaleWidth      =   2115
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1200
      Top             =   0
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Sound Off"
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1155
   End
End
Attribute VB_Name = "Form_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
'form_menu.Check1.Value = vbChecked

Soff = True
If Form_Menu.Check1.Value = vbChecked Then
    Soff = True
Else
    Soff = False
End If
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
Form_Menu.Top = ATidalX.Top + 1000
Form_Menu.Left = ATidalX.Left + 100

End Sub

Private Sub Form_Unload(Cancel As Integer)
If QuitForms = False Then Cancel = True
Form_Menu.Hide

End Sub

Private Sub Timer1_Timer()

Form_Menu.Hide
Timer1.Enabled = False

End Sub
