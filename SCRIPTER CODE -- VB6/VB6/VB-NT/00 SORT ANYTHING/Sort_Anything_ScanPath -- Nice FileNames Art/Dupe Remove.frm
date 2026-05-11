VERSION 5.00
Begin VB.Form Dupe 
   Caption         =   "Form1"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15435
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   15435
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   150
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   180
      Width           =   15075
   End
End
Attribute VB_Name = "Dupe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub List1_Click()


q = List1.ListIndex

b$ = Mid$(List1.List(q), iit)
On Local Error Resume Next
'Kill b$

End Sub
