VERSION 5.00
Begin VB.Form Restarter 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   990
      Top             =   810
   End
End
Attribute VB_Name = "Restarter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()

If ReRun = True And Forms.Count = 1 Then
    Call ReRunner
    
End If

If Forms.Count = 1 Then End

End Sub
