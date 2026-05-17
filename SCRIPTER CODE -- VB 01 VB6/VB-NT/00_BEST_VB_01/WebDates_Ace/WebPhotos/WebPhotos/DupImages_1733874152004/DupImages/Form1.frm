VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "DupImages"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "confirm file delete"
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   120
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3240
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   960
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "all"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   1440
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "duplicates"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   1440
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "search && display"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "search subfolders"
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "filetype"
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   960
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim fd As Boolean
  Screen.MousePointer = 11
  fd = False
  If Check1.Value = 1 Then fd = True
  If Option1 = True Then ' dups
      dups Label2.Caption, Combo1.List(Combo1.ListIndex), fd
  Else ' all
      all Label2.Caption, Combo1.List(Combo1.ListIndex), fd
  End If
  Screen.MousePointer = 0
End Sub

Private Sub Dir1_Change()
    Label2.Caption = Dir1.List(Dir1.ListIndex)
End Sub

Private Sub Dir1_Click()
    Label2.Caption = Dir1.List(Dir1.ListIndex)
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    Label2.Caption = Dir1.Path
    Combo1.AddItem "jpg"
    Combo1.AddItem "gif"
    Combo1.AddItem "bmp"
    Combo1.ListIndex = 0
    Drive1.Drive = Left(App.Path, 2)
End Sub
