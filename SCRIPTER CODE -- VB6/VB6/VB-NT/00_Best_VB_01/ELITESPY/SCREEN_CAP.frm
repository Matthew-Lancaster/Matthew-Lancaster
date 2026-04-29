VERSION 5.00
Begin VB.Form SCREEN_CAP 
   Caption         =   "Form1"
   ClientHeight    =   4944
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9804
   LinkTopic       =   "Form1"
   ScaleHeight     =   4944
   ScaleWidth      =   9804
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   6660
      Top             =   1560
   End
End
Attribute VB_Name = "SCREEN_CAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Unload Me
End Sub
