VERSION 5.00
Begin VB.Form Counter1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Counter1"
   ClientHeight    =   4128
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4128
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer5 
      Interval        =   1000
      Left            =   4965
      Top             =   1740
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4485
      Top             =   1740
   End
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      Height          =   816
      Left            =   -15
      TabIndex        =   5
      Top             =   885
      Width           =   4875
   End
   Begin VB.Timer Timer3 
      Interval        =   10000
      Left            =   4035
      Top             =   1725
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3600
      Top             =   1725
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3150
      Top             =   1725
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
      TabIndex        =   7
      ToolTipText     =   "Disable Hide Cursor"
      Top             =   3405
      Width           =   1515
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
      TabIndex        =   6
      ToolTipText     =   "Disable Hide Cursor"
      Top             =   3045
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
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   135
      TabIndex        =   0
      Top             =   -195
      Width           =   1725
   End
End
Attribute VB_Name = "Counter1"
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
Public LK2, OldDate ', FreeMem, IIO2, IIO3


Private Sub Form_Load()
Me.BackColor = 0



tr$ = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\00 Knocker Boy1.txt"

If Dir$(tr$) <> "" Then
    Open tr$ For Input As #1
    Do
    Line Input #1, R2$
    
    If R2$ <> "" Then
        R3$ = R2$
    End If
'    If Mid$(r2$, 12, 3) <> mm$ Then TTCount1 = 0
'    TTCount1 = TTCount1 + 1
'    mm$ = Mid$(r2$, 12, 3)
    'List1.AddItem r2$
    Loop Until EOF(1)
    R2$ = R3$
Close #1
End If

TTCount1 = Val(Mid$(R2$, 1, 4))
If R2$ <> "" Then
OldDate = DateValue(Mid$(R2$, 16, 9)) + TimeValue(Mid$(R2$, 26, 9))
End If

If OldDate < xdate Then
'TTCount1 = 0: OldDate = Now
End If
'
'If Int(Now) <> DateValue(Mid$(r2$, 16, 9)) Then TTCount1 = 0: OldDate = Int(Now)

TTCount2 = Val(Mid$(R2$, 6, 5))
Label10.Caption = Trim(str(TTCount1))

'    Open App.Path + "\00_Text_Data\00 Knocker Boy.txt" For Output As #1
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

    If OldDate < xdate Then
        TTCount1 = 0
    End If
    
    OldDate = Now
    
    
    'Call FileInUseDelay(tr$, True)
    FILENAME_IN_USE_CHECK = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\00 Knocker Boy1.txt"
    FILENAME_IN_USE_CHECK_4 = FILENAME_IN_USE_CHECK
    DumVar = IsFileOpenDelay(FILENAME_IN_USE_CHECK)
    FILENAME_IN_USE_CHECK = FILENAME_IN_USE_CHECK_4
    tr$ = FILENAME_IN_USE_CHECK
    FR1 = FreeFile
    Open FILENAME_IN_USE_CHECK For Append As #FR1
    
    TTCount1 = TTCount1 + 1
    TTCount2 = TTCount2 + 1
    Label10 = Trim(str(TTCount1))
    If Vcodez = 145 Then fg$ = "Mal Knock- " Else fg$ = ""
    If Menu.SoundMall = vbChecked Then gag$ = "Hittback Vol =" + str(100 - Menu.Slider1.Value) + "- " Else gag$ = ""
    If fg$ = "" And Menu.SoundMall = vbChecked Then fg$ = " Armed On- ": gag$ = ""
    If Menu.MuteAllMal = vbChecked Then fg$ = fg$ + " Mute All Mal HittBacks- "
    If Menu.MuteAllMal = vbUnchecked Then fg$ = fg$ + " Quite Spit Sound- "
    If Ctrl = True Then fg$ = "Quick CTRL Key Vol =" + str(100 - Menu.Slider1.Value) + "- "
    QQ$ = Format$(TTCount1, "0000 ") + Format$(TTCount2, "00000-") + Format$(Now, "DDD-DD-MMM-YY HH:MM:SSa/p ") + fg$ + gag$
    MSGRESULT1 = ATidalX.WinAmpChkIsPlaying
    If MSGRESULT1 = 1 Then QQ$ = QQ$ + " Music Playing Vol =" + str(GetVolume(SPEAKER)) + "- "
    If MSGRESULT1 <> 1 Then QQ$ = QQ$ + " Music Not Playing- "
    
    Print #FR1, QQ$
    Close #FR1

    ADate1Old = 0
    
    Label10.BackColor = Label5.BackColor
    Counter1.BackColor = Label5.BackColor
    Timer4.Enabled = True
    
'    ATidalX.MMControl7.Command = "prev"
'    ATidalX.MMControl7.Command = "Play"
    
    BlockCountDupe = False

End Sub

Private Sub Form_Resize()
Call Timer1_Timer
End Sub



Private Sub Mnu_Now2_Click()
'Shell "notepad " + App.Path + "\00_Text_Data\Idle and Active Logger Now2.txt", vbNormalFocus

End Sub

Private Sub Label8_Click()
'Drives.Label8
End Sub

Public Sub Label10_Change()
Counter1.Label10.FontSize = 35
Counter2.Label10.FontSize = Counter1.Label10.FontSize
Counter3.Label10.FontSize = Counter1.Label10.FontSize
Counter4.Label10.FontSize = Counter1.Label10.FontSize
Counter2.Label10.Font = Counter1.Label10.Font
Counter3.Label10.Font = Counter1.Label10.Font
Counter4.Label10.Font = Counter1.Label10.Font
Counter1.Label10.Top = 0
Counter1.Label10.Left = 0
Counter2.Label10.Top = 0
Counter2.Label10.Left = 0
Counter3.Label10.Top = 0
Counter3.Label10.Left = 0
Counter4.Label10.Top = 0
Counter4.Label10.Left = 0

WC = 0: HC = 0
On Error Resume Next
For Each Control In Counter1
    If Control.Enabled = True And Mid$(Control.Name, 1, 4) <> "Time" Then
        'control.name
        If Control.Left + Control.Width > WC Then WC = Control.Left + Control.Width
        If Control.Top + Control.Height > HC Then HC = Control.Top + Control.Height
    End If
Next
For Each Control In Counter2
    If Control.Enabled = True And Mid$(Control.Name, 1, 4) <> "Time" Then
        If Control.Left + Control.Width > WC Then WC = Control.Left + Control.Width
        If Control.Top + Control.Height > HC Then HC = Control.Top + Control.Height
    End If
Next
For Each Control In Counter3
    If Control.Enabled = True And Mid$(Control.Name, 1, 4) <> "Time" Then
        If Control.Left + Control.Width > WC Then WC = Control.Left + Control.Width
        If Control.Top + Control.Height > HC Then HC = Control.Top + Control.Height
    End If
Next
For Each Control In Counter4
    If Control.Enabled = True And Mid$(Control.Name, 1, 4) <> "Time" Then
        If Control.Left + Control.Width > WC Then WC = Control.Left + Control.Width
        If Control.Top + Control.Height > HC Then HC = Control.Top + Control.Height
    End If
Next
On Error GoTo 0
Counter1.Width = WC
Counter1.Height = HC - 190
Counter2.Width = WC
Counter2.Height = HC - 190
Counter3.Width = WC
Counter3.Height = HC - 190
Counter4.Width = WC
Counter4.Height = HC - 190

Counter1.BackColor = Counter1.Label10.BackColor
Counter2.BackColor = Counter2.Label10.BackColor
Counter3.BackColor = Counter3.Label10.BackColor
Counter4.BackColor = Counter4.Label10.BackColor
Counter1.Label10.Top = -100
Counter1.Label10.Left = Counter1.Width - Counter1.Label10.Width
Counter2.Label10.Top = -100
Counter2.Label10.Left = (Counter2.Width - Counter2.Label10.Width) / 2
Counter3.Label10.Top = -100
Counter3.Label10.Left = (Counter3.Width - Counter3.Label10.Width) / 2
Counter4.Label10.Top = -100
Counter4.Label10.Left = (Counter4.Width - Counter4.Label10.Width) / 2

End Sub

Private Sub Label10_DblClick()
    
    Call ATidalX.Mnu_Knocker1_Click

End Sub

Private Sub Timer1_Timer()

hwndw = FindWindow(vbNullString, ATidalX.Caption)
If hwndw = 0 Then Exit Sub
Dim Rect4 As RECT
Hwnd24 = GetWindowRect(hwndw, Rect4)
On Local Error Resume Next
If Ik2 <> Rect4.Top Then
    Counter1.Top = (Rect4.Top * Screen.TwipsPerPixelX)
    Counter1.Left = (ATidalX.Left - Counter1.Width)
End If
Ik2 = Rect4.Top
End Sub

Private Sub Timer2_Timer()



Exit Sub
ErrHand:
Exit Sub

End Sub


Private Sub Timer4_Timer()
Label10.BackColor = Label6.BackColor
Counter1.BackColor = Label6.BackColor
Timer4.Enabled = False
End Sub

Private Sub Timer5_Timer()
    Exit Sub

    On Local Error Resume Next
    FileSpec = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\00 Knocker Boy1.txt"
    Set F = FS.GetFile((FileSpec))
    Adate1 = F.DateLastModified
    If ADate1Old > 0 And ADate1Old <> Adate1 Then
        Unload Counter1
        Counter.Show
    End If
    
    ADate1Old = Adate1

    FileSpec = App.Path + "\00_Text_Data" + GetComputerName + "-" + GetUserName + "\\00 Knocker Boy2.txt"
    Set F = FS.GetFile((FileSpec))
    Adate1 = F.DateLastModified
    If ADate2Old > 0 And ADate2Old <> Adate1 Then
        Unload Counter2
        Counter2.Show
    End If
    ADate2Old = Adate1

    FileSpec = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\00 Knocker Boy3.txt"
    Set F = FS.GetFile((FileSpec))
    Adate1 = F.DateLastModified
    If ADate3Old > 0 And ADate3Old <> Adate1 Then
        Unload Counter3
        Counter3.Show
    End If
    ADate3Old = Adate1
    
    On Local Error Resume Next
    FileSpec = App.Path + "\00_Text_Data\" + GetComputerName + "-" + GetUserName + "\00 Knocker Boy4.txt"
    Set F = FS.GetFile((FileSpec))
    Adate1 = F.DateLastModified
    If ADate4Old > 0 And ADate4Old <> Adate1 Then
        Unload Counter4
        Counter4.Show
    End If
    ADate4Old = Adate1
    
End Sub
