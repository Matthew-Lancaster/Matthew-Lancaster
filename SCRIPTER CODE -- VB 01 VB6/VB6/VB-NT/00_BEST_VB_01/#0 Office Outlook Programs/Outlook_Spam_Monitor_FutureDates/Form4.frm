VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   Caption         =   "Form3"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13605
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   2700
   ScaleWidth      =   13605
   Visible         =   0   'False
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13590
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
Call Form1.xy
Form4.Hide
End Sub

Public Sub List1_Click()
If Title$ <> "" Then
Form4.List1.AddItem Title$
Title$ = ""
Form4.List1.ListIndex = Form4.List1.ListCount - 1
If List1.ListCount > 1 Then Form4.List1.AddItem "---------------------------"
DoEvents
Exit Sub
End If

If t1$ = "" Then Exit Sub
Form4.List1.AddItem t1$ + t2$ + t3$ + t4$
Form4.List1.AddItem t1$ + t2$ + t3$ + t5$
Form4.List1.AddItem t1$ + t2$ + t3$ + t6$
Form4.List1.AddItem t1$ + t2$ + t3$ + t7$
Form4.List1.AddItem "---------------------------"
t1$ = ""
Form4.List1.ListIndex = Form4.List1.ListCount - 1
DoEvents

End Sub

