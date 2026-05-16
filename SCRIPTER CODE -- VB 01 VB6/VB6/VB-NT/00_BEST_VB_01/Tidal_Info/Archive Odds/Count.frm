VERSION 5.00
Begin VB.Form Counter 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Counter1"
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4485
      Top             =   1740
   End
   Begin VB.ListBox List1 
      Enabled         =   0   'False
      Height          =   840
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
         Size            =   15.75
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
         Size            =   15.75
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
         Size            =   14.25
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
         Size            =   14.25
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
         Size            =   15.75
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
         Size            =   15.75
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
      Height          =   960
      Left            =   0
      TabIndex        =   0
      Top             =   -180
      Width           =   1995
   End
End
Attribute VB_Name = "Counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'Public CHD, MHD, NHD
Public LK2, OldDate ', FreeMem, IIO2, IIO3


Private Sub Form_Load()
Me.BackColor = 0

wc = 0: hc = 0
On Error Resume Next
For Each Control In Counter
    If Control.Enabled = True Then
        If Control.Left + Control.Width > wc Then wc = Control.Left + Control.Width
        If Control.Top + Control.Height > hc Then hc = Control.Top + Control.Height
    End If
Next
On Error GoTo 0
Counter.Width = wc + 90 + hq
Counter.Height = hc + 380
Counter.Width = wc + hq
Counter.Height = hc


tr$ = App.Path + "\00 TextData\00 Knocker Boy1.txt"

If Dir$(tr$) <> "" Then
    Open tr$ For Input As #1
    Do
    Line Input #1, r2$
    
    If r2$ <> "" Then
        r3$ = r2$
    End If
'    If Mid$(r2$, 12, 3) <> mm$ Then TTCount1 = 0
'    TTCount1 = TTCount1 + 1
'    mm$ = Mid$(r2$, 12, 3)
    'List1.AddItem r2$
    Loop Until EOF(1)
    r2$ = r3$
Close #1
End If

TTCount1 = Val(Mid$(r2$, 1, 4))
If r2$ <> "" Then
OldDate = DateValue(Mid$(r2$, 16, 9)) + TimeValue(Mid$(r2$, 26, 9))
End If

If OldDate < xdate Then
'TTCount1 = 0: OldDate = Now
End If
'
'If Int(Now) <> DateValue(Mid$(r2$, 16, 9)) Then TTCount1 = 0: OldDate = Int(Now)

TTCount2 = Val(Mid$(r2$, 6, 5))
Label10.Caption = Trim(Str(TTCount1))

'    Open App.Path + "\00 TextData\00 Knocker Boy.txt" For Output As #1
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
    
    tr$ = App.Path + "\00 TextData\00 Knocker Boy1.txt"
    Open tr$ For Append As #1
    
    TTCount1 = TTCount1 + 1
    TTCount2 = TTCount2 + 1
    Label10 = Trim(Str(TTCount1))
    If Vcodez = 145 Then fg$ = "---" Else fg$ = ""
    'gag$ = "Off "
    If Menu.SoundMall = vbChecked Then gag$ = "Hittback" + Str(100 - Menu.Slider1.Value) + " Volume" Else gag$ = ""
    If fg$ = "" And Menu.SoundMall = vbChecked Then fg$ = "--- Armed On": gag$ = ""
    If Shift = True Then fg$ = "Quick Shift Key " + Str(100 - Menu.Slider1.Value) + " Volume"
    If Menu.MuteAllMal = vbChecked Then fg$ = fg$ + " -- Mute All Mal "
    qq$ = Format$(TTCount1, "0000 ") + Format$(TTCount2, "00000-") + Format$(Now, "DDD-DD-MMM-YY HH:MM:SSa/p ") + fg$ + gag$
    'List1.AddItem qq$
    'List1.ListIndex = List1.ListCount - 1
    Print #1, qq$
    Close #1

    ADate1Old = 0
    
    Label10.BackColor = Label5.BackColor
    Timer4.Enabled = True
    
    
'    ATidalX.MMControl7.Command = "prev"
'    ATidalX.MMControl7.Command = "Play"
    
BlockCountDupe = False

End Sub

Private Sub Form_Resize()
Call Timer1_Timer
End Sub



Private Sub Mnu_Now2_Click()
'Shell "notepad " + App.Path + "\00 TextData\Idle and Active Logger Now2.txt", vbNormalFocus

End Sub

Private Sub Label8_Click()
'Drives.Label8
End Sub

Private Sub Label10_DblClick()
    
    Call ATidalX.Mnu_Knocker_Click

End Sub

Private Sub Timer1_Timer()
hwndw = FindWindow(vbNullString, Drives.Caption)
If hwndw = 0 Then Exit Sub
Dim Rect4 As RECT
Hwnd24 = GetWindowRect(hwndw, Rect4)
If Ik2 <> Rect4.Top Then
    Counter.Top = (Rect4.Top * Screen.TwipsPerPixelX)
    Counter.Left = (Drives.Left - Counter.Width)
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
Timer4.Enabled = False
End Sub
