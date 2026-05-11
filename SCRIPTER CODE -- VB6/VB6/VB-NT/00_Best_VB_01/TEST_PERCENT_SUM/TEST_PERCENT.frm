VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

'Debug.Print (5.32) + (5.32 / 100) * 25

'END IN .32
Var = 7.45 ' 9.3125
Var = 7.46 ' 9.325
Var = 8.26 ' 3649.005
Var = 7.46 ' 3648.005
Var = 9.07 ' 3648.005
Var = 1.8   ' 3648.005
Debug.Print 3638.68 + ((Var) + (Var / 100) * 25)
Debug.Print 9.07
Debug.Print (Var / 100) * 25
Debug.Print Var + (Var / 100) * 25
Debug.Print 3638.68 + 11.34
Debug.Print 3650.02

End

End Sub
