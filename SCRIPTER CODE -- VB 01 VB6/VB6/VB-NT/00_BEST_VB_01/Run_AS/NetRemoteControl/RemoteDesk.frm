VERSION 5.00
Begin VB.Form RemoteDesk 
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5505
   ScaleWidth      =   6915
End
Attribute VB_Name = "RemoteDesk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
Dim W As Single, H As Single
    
    W = Me.ScaleWidth - Screen.TwipsPerPixelX * 1
    H = Me.Height - Screen.TwipsPerPixelY * 27
    'Picture1.Move 1, 1, W, H
    Picture2.Move 1, 1, W, H
        
End Sub
