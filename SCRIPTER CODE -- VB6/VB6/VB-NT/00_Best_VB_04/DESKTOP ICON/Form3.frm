VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DeskTop Icon Tools Restore"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Restore"
      Height          =   2535
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   3615
      Begin VB.Label Label2 
         Caption         =   "Restoring Icons to Desktop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   3375
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   3360
      Pattern         =   "*.pps"
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Please Choose a configuration to restore"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If List1.ListIndex < 0 Then Exit Sub



Frame1.Visible = True
Refresh
Dim iconcount As Long
Dim found As Long
Dim IconPosition3() As POINTAPI
Dim IconName3() As String

FileNum = FreeFile
On Error Resume Next

Open CurrentDirectory + List1.List(List1.ListIndex) + ".pps" For Input As #FileNum

If Err > 0 Then   'oops
    MsgBox "ERROR: Opening File"
    Exit Sub
Else
    
    Input #FileNum, iconcount
        
    ReDim IconPosition3(iconcount) As POINTAPI
    ReDim IconName3(iconcount) As String
        'we should do some error checking here <g>
    For i = 1 To iconcount
        Input #FileNum, IconName3(i - 1)
        Input #FileNum, IconPosition3(i - 1).x
        Input #FileNum, IconPosition3(i - 1).y
    Next i
    'TODO:
    'will have resolution code here as well
    'next will be width
    'then height
    'for multiple restore configurations
    Dim SWidth As Long
    Dim SHeight As Long
    Input #FileNum, SWidth
    Input #FileNum, SHeight
    
    
    
    
    
    Close #FileNum

    'ok , now we have the
    'names and positions of previous save


        'scroll through the list we got at startup
        'checking to see if our saved icons are here
        'then respot them
        'we'll only restore if we find a name match
        'icount is the no. of icons found when we loaded
    

    Dim diff2 As Single
    Dim diff As Single
                     'restore to fit icons into the
    Dim tempone      'same pattern we have in the file
    Dim temptwo      'but scaled to fit the screen
    diff = SWidth / StartWidth
    diff2 = SHeight / StartHeight
    
    For i = 0 To icount - 1
        For n = 0 To iconcount - 1
            
            'need to RTrim here in case user has put
            'spaces after the icon name  -bugfix-
            
            If RTrim$(IconName3(n)) = RTrim$(IconName(i)) Then
            'bingo, same text as what we have in memory
            found = found + 1
            'set saved icon postion
            
            If SWidth = CurrentWidth And SHeight = CurrentHeight Then
                 
                Call SendMessageByLong(hdesk, LVM_SETITEMPOSITION, i, _
                CLng(IconPosition3(n).x + IconPosition3(n).y * &H10000))
                
            Else
           
                tempone = CLng(IconPosition3(n).x / diff)
           Debug.Print Str(IconPosition3(n).x) + Str(tempone)
                temptwo = CLng(IconPosition3(n).y / diff2)
                Call SendMessageByLong(hdesk, LVM_SETITEMPOSITION, i, _
                CLng(tempone + temptwo * &H10000))
            
            End If
                    
            
            Exit For 'exit this loop and look for next iconname
            Else
            
            End If
        Next n
    Next i
'all done tell the user
Form1.Text1.Text = Str(found) + " Icons returned to their Positions."
End If

FindIcons 'reload all Vars
Form1.Show

Unload Form3

End Sub

Private Sub Command2_Click()
'cancel
Form1.Show
Unload Form3



End Sub

Private Sub Command3_Click()
If List1.ListIndex < 0 Then Exit Sub

Kill CurrentDirectory + List1.List(List1.ListIndex) + ".pps"
'reload listbox
List1.Clear
File1.Refresh

For i = 0 To File1.ListCount - 1

    List1.AddItem Left$(File1.List(i), Len(File1.List(i)) - 4)

Next i

End Sub

Private Sub Form_Load()
Form1.Hide

Show

If File1.ListCount < 1 Then



    'MsgBox "ERROR:  No Saved Configurations"

Else

For i = 0 To File1.ListCount - 1

    List1.AddItem Left$(File1.List(i), Len(File1.List(i)) - 4)

Next i



End If



End Sub

