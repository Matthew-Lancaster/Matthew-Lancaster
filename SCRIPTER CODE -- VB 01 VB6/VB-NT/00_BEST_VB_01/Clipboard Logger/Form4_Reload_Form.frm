VERSION 5.00
Begin VB.Form Form4_Reload_Form 
   Caption         =   "Form2"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form2"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   576
      Top             =   660
   End
End
Attribute VB_Name = "Form4_Reload_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Load Form1
Form1.Show
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Unload Form4_Reload_Form
End Sub
