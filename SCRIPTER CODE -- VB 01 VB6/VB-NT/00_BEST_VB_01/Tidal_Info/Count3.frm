VERSION 5.00
Begin VB.Form Counter3 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Counter3"
   ClientHeight    =   3996
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4932
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3996
   ScaleWidth      =   4932
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4035
      Top             =   195
   End
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      Height          =   240
      Left            =   0
      TabIndex        =   5
      Top             =   1185
      Width           =   4875
   End
   Begin VB.Timer Timer3 
      Interval        =   10000
      Left            =   3555
      Top             =   195
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3120
      Top             =   195
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Back3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2775
      TabIndex        =   7
      ToolTipText     =   "Disable Hide Cursor"
      Top             =   3045
      Width           =   1515
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Back4"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2775
      TabIndex        =   6
      ToolTipText     =   "Disable Hide Cursor"
      Top             =   3405
      Width           =   1515
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Mem"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "View keyboard and Mouse Log With Notepad"
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "--"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   615
      TabIndex        =   3
      ToolTipText     =   "F12 to Command Prompt - Disable"
      Top             =   2400
      Width           =   2010
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Back2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2775
      TabIndex        =   2
      ToolTipText     =   "Disable Hide Cursor"
      Top             =   2685
      Width           =   1515
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Back1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2775
      TabIndex        =   1
      ToolTipText     =   "Disable Hide Cursor"
      Top             =   2325
      Width           =   1515
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "0--0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   840
      Left            =   345
      TabIndex        =   0
      Top             =   -120
      Width           =   1305
   End
End
Attribute VB_Name = "Counter3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FILENAME_IN_USE_CHECK As String

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'Public CHD, MHD, NHD
Public LK2, OldDate3 ', FreeMem, IIO2, IIO3


Private Sub Form_Load()
Me.BackColor = 0
Label10.Top = Counter1.Label10.Top
Label10.Height = Counter1.Label10.Height
Label10.Left = Counter1.Label10.Left
Label10.Width = Counter1.Label10.Width
Label10.FontSize = Counter1.Label10.FontSize

WC = 0: HC = 0
On Error Resume Next
For Each Control In Counter3
    If Control.Enabled = True Then
        If Control.Left + Control.Width > WC Then WC = Control.Left + Control.Width
        If Control.Top + Control.Height > HC Then HC = Control.Top + Control.Height
    End If
Next
On Error GoTo 0
'Counter3.Width = wc + 90 + hq
'Counter3.Height = hc + 380
Counter3.Width = WC + hq
Counter3.Height = HC




If Dir$(App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\00 Knocker Boy3.txt") <> "" Then
    Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\00 Knocker Boy3.txt" For Input As #1
    Do
    Line Input #1, R2$
'    If Mid$(r2$, 12, 3) <> mm$ Then TTCount1 = 0
'    TTCount1 = TTCount1 + 1
'    mm$ = Mid$(r2$, 12, 3)
    'List1.AddItem r2$
    Loop Until EOF(1)
Close #1
End If

TTCount5 = Val(Mid$(R2$, 1, 4))
OldDate3 = DateValue(Mid$(R2$, 16, 9)) + TimeValue(Mid$(R2$, 26, 9))


'If r2$ <> "" Then
'    TTCount5 = Val(Mid$(r2$, 1, 4))
'    OldDate3 = DateValue(Mid$(r2$, 16, 9))
    TTCount6 = Val(Mid$(R2$, 6, 5))
'    If Int(Now) <> DateValue(Mid$(r2$, 16, 9)) Then TTCount5 = 0: OldDate3 = Int(Now)
'End If

Label10.Caption = Trim(str(TTCount5))

'    Open App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\00 Knocker Boy.txt" For Output As #1
'    For R = 1 To List1.ListCount - 1
'    Print #1, List1.List(R)
'    Next
'    Close #1

End Sub

Public Sub WriteCounter()
    If BlockCountDupe = True Then Exit Sub
    BlockCountDupe = True

    ii = HOUR(Now)
    If ii < 25 Then xdate = Int(Now) + TimeSerial(21, 0, 0)
    If ii < 21 Then xdate = Int(Now) + TimeSerial(14, 0, 0)
    If ii < 14 Then xdate = Int(Now) + TimeSerial(7, 0, 0)
    If ii < 7 Then xdate = Int(Now) - 1 + TimeSerial(21, 0, 0)

    If OldDate3 < xdate Then
        TTCount5 = 0
    End If
    
    OldDate3 = Now
    
    FILENAME_IN_USE_CHECK = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\00 Knocker Boy3.txt"
    FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
    DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
    FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
    tr$ = FILENAME_IN_USE_CHECK
    FR1 = FreeFile
    Open FILENAME_IN_USE_CHECK For Append As #FR1
    
    TTCount5 = TTCount5 + 1
    TTCount6 = TTCount6 + 1
    Label10 = Trim(str(TTCount5))
    If Menu.MuteAllMal = vbChecked Then fg$ = fg$ + " -- Mute All Mal"
    QQ$ = Format$(TTCount5, "0000 ") + Format$(TTCount6, "00000-") + Format$(Now, "DDD-DD-MMM-YY HH:MM:SSa/p ") + "-" + fg$
    'List1.AddItem qq$
    'List1.ListIndex = List1.ListCount - 1
    Print #FR1, QQ$
    Close #FR1
    
    ADate3Old = 0

    Label10.BackColor = Label5.BackColor
    Counter3.BackColor = Label5.BackColor
    Timer4.Enabled = True
    
    BlockCountDupe = False

Exit Sub
    
If Menu.MuteAllMal = vbUnchecked Then

VolumeSet1 = GetVolume(SPEAKER)
    
tf = SetVolume(SPEAKER, 2)

    ATidalX.MMControl7.Command = "prev"
    ATidalX.MMControl7.Command = "Play"
    
    
Do
    DoEvents
    wer1 = ATidalX.MMControl7.mode
    Q = Q + 1
Loop Until wer1 = 525
    
    
tf = SetVolume(SPEAKER, VolumeSet1)


End If



End Sub

Private Sub Form_Resize()
Call Timer1_Timer
End Sub



Private Sub Mnu_Now2_Click()
'Shell "notepad " + App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\Idle and Active Logger Now2.txt", vbNormalFocus

End Sub

Private Sub Label8_Click()
'Drives.Label8
End Sub

Private Sub Label10_Change()
Call Counter1.Label10_Change
End Sub

Private Sub Label10_DblClick()
    
    Call ATidalX.Mnu_Knocker3_Click

End Sub

Private Sub Timer1_Timer()
hwndw = FindWindow(vbNullString, Counter2.Caption)
If hwndw = 0 Then Exit Sub
Dim Rect4 As RECT
Hwnd24 = GetWindowRect(hwndw, Rect4)
On Local Error Resume Next
If Ik2 <> Rect4.Top Then
    Counter3.Top = (Rect4.Bottom * Screen.TwipsPerPixelX)
    Counter3.Left = (Counter2.Left) '- Counter.Width)
End If
Ik2 = Rect4.Top
End Sub

Private Sub Timer4_Timer()
Label10.BackColor = Label6.BackColor
Counter3.BackColor = Label6.BackColor
Timer4.Enabled = False
End Sub

