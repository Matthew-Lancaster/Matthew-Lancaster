VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save DeskTop Icon Positions"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Please type in a Name for the current configuration"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'save

If Text1.Text = "" Then
MsgBox "ERROR:  No text entered"
    Form1.Show
    Unload Form2
Else
    
Form1.List1.Clear
FindIcons
FileNum = FreeFile






Open CurrentDirectory + Text1.Text + ".pps" For Output As #FileNum
'TODO:
'check for errors

Print #FileNum, Str(icount)
For i = 0 To icount - 1
    Print #FileNum, IconName(i)
    Print #FileNum, IconPosition(i).x
    Print #FileNum, IconPosition(i).y
Next i
    'next will be width
    'then height
    'for multiple restore configurations
SWidth = Screen.Width \ Screen.TwipsPerPixelX
SHeight = Screen.Height \ Screen.TwipsPerPixelY
Print #FileNum, SWidth
Print #FileNum, SHeight



Close #FileNum



    
    
    
    
    Form1.Show
    Unload Form2
End If

End Sub

Private Sub Command2_Click()
'cancel
Form1.Show
Unload Form2
End Sub

Private Sub Form_Load()
Show
End Sub
