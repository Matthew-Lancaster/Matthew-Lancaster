VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'D:\My Webs\MatthewLan.Com Web\LoveFolder\Cool\2sort-Lxx2.txt


Open "D:\My Webs\MatthewLan.Com Web\LoveFolder\Cool\2sort-Lxx2.txt" For Input As #1
Do

Loop Until 1 = 1

Close #1



End Sub
