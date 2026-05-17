VERSION 5.00
Begin VB.Form ProcLB 
   Caption         =   "Processes List"
   ClientHeight    =   5372
   ClientLeft      =   68
   ClientTop       =   340
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   5372
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox List1 
      Height          =   5151
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
   End
End
Attribute VB_Name = "ProcLB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)


Me.Hide

If MustUnload = True Then Exit Sub
Cancel = True
End Sub

