VERSION 5.00
Begin VB.Form SortFrm 
   Caption         =   "SortFrm"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "SortFrm"
   ScaleHeight     =   4065
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1845
      Top             =   1125
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   1845
      Top             =   1530
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select What You Want To Delete and Hit Here"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   0
      Picture         =   "2Sort-Copy.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3285
      Width           =   4650
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   0
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   495
      Width           =   4650
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "... Duplicates ..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4650
   End
End
Attribute VB_Name = "SortFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

f5 = FreeFile
Open "C:\callerid\2double.txt" For Input As #f5
slurp1$ = Input$(LOF(f5), f5)
Close f5
pir = 0
For r = 0 To SortFrm.List1.ListCount - 1
    If SortFrm.List1.Selected(r) = True Then
        
        qw1 = InStrRev(slurp1$, SortFrm.List1.List(r))
        If qw1 > 0 And Mid$(SortFrm.List1.List(r), 1, 1) <> "#" Then
            'qw2 used else where
            qw3 = InStr(qw1 + 1, slurp1$, vbCr)
            xx1$ = Right$(Left$(slurp1$, qw1 - 3), 9)
            xx2$ = Mid$(slurp1$, qw3 + 2)
            slurp1$ = Left$(slurp1$, qw1 - 3) + Mid$(slurp1$, qw3)
            pir = 1
        End If
    End If
Next

For r = SortFrm.List1.ListCount - 1 To 0 Step -1
    If SortFrm.List1.Selected(r) = True Then
        SortFrm.List1.RemoveItem (r)
    End If
Next

If pir = 1 Then
    f9 = FreeFile
    Open "C:\callerid\2double.txt" For Output As #f9
    Print #f9, slurp1$;
    Close f9
    SortFrm.List1.ListIndex = SortFrm.List1.ListCount - 1
End If

Call UpDating

'SortFrm.Command1.Visible = False

If SortFrmPop = 0 Then
SortFrm.Hide
End If

Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
SortFrm.Hide
Timer1.Enabled = False
Form1.Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
'If Form1.WindowState = 0 Then SortFrm.WindowState = 0
If Form1.WindowState = 1 Then SortFrm.Hide

End Sub
