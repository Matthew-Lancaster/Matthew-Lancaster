VERSION 5.00
Begin VB.Form FrmJoy 
   BackColor       =   &H80000007&
   Caption         =   "Form2"
   ClientHeight    =   6924
   ClientLeft      =   60
   ClientTop       =   372
   ClientWidth     =   8316
   LinkTopic       =   "Form2"
   ScaleHeight     =   6924
   ScaleWidth      =   8316
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox List1 
      Height          =   1008
      Left            =   2760
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   1680
      Width           =   4110
   End
   Begin VB.Timer Timer_XxXJOY 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2985
      Top             =   3270
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      DrawWidth       =   5
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2355
      ScaleMode       =   0  'User
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   -15
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   9
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   8
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   840
      TabIndex        =   5
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   1440
      TabIndex        =   4
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   2040
      TabIndex        =   3
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   840
      TabIndex        =   2
      Top             =   3465
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   9
      Left            =   1440
      TabIndex        =   1
      Top             =   3465
      Width           =   375
   End
End
Attribute VB_Name = "FrmJoy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SMX, SMY, OX1, OY1, GF, JJ, OJJ, OKJ, KeyBounce, OJKY, OB1, OBT1, OBT2, JoyQuickTrig, XYN
Dim j As JOYINFO, HALO2, TimeWait, OldSluty3, PreviNextFF, PreviNextRW, OTT, JoyPress, PY, x1, y1
Dim jold As JOYINFO
Dim je As JOYCAPS
Dim TTB
Const Red = &HFF&
Const White = &HFFFFFF
Dim be As B
Dim be1 As B
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Const SW_HIDE = 0
Const SW_SHOWNORMAL = 1
Const SW_NORMAL = 1
Const SW_SHOWMINIMIZED = 2
Const SW_SHOWMAXIMIZED = 3
Const SW_MAXIMIZE = 3
Const SW_SHOWNOACTIVATE = 4
Const SW_SHOW = 5
Const SW_MINIMIZE = 6
Const SW_SHOWMINNOACTIVE = 7
Const SW_SHOWNA = 8
Const SW_RESTORE = 9
Const SW_SHOWDEFAULT = 10
Const SW_FORCEMINIMIZE = 11
Const SW_MAX = 11

Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Function GetWindowTitle(ByVal hwnd As Long) As String
   Dim L As Long
   Dim S As String
   L = GetWindowTextLength(hwnd)
   S = Space(L + 1)
   GetWindowText hwnd, S, L + 1
   GetWindowTitle = Left$(S, L)
End Function



Private Sub Form_Load()
If IsWinNT Then
    joyGetDevCapsW 0, je, Len(je)
Else
    joyGetDevCapsA 0, je, Len(je)
End If
If je.wXmax > 0 Then Picture1.ScaleHeight = je.wXmax
If je.wYmax > 0 Then Picture1.ScaleWidth = je.wYmax
Picture1.ScaleLeft = 0
Picture1.ScaleTop = 0
'FrmJoy.Show
'Timer1.Interval = 10
Me.Hide
End Sub




Public Sub Timer1_Timerxxx()

Exit Sub

GetBPress j.Buttons, be
joyGetPos 0, j
        
'Updated all this pucker on 28-07-07 3:56pm Knackarded Working Staving

        
KEY1 = Int(32767 / 1.2)
KEY2 = Int(65535 / 1.2)
'Debug.Print str(j.X) + " ---- " + str(j.Y)
'----------

JJ = 0
If j.x = 32767 And j.y = 32511 Then JJ = 0 'Idle
'----------
If j.x = 0 And j.y > KEY1 Then JJ = 2 'Left
If j.x > KEY1 And j.y = 0 Then JJ = 3 'Up
If j.x > KEY1 And j.y > KEY2 Then JJ = 4 'Down
If j.x > KEY2 And j.y > KEY1 Then JJ = 5 'Right
'----------
If j.x = 0 And j.y = 0 Then JJ = 6  'Left Up -- Top Left
If j.x > KEY2 And j.y = 0 Then JJ = 7 'Up Right -- Top Right
If j.x = 0 And j.y > KEY2 Then JJ = 8 'Down Left  Bottom Left
If j.x > KEY2 And j.y > KEY2 Then JJ = 9 'Right Down -- Bottom Right
'----------
        
'If JJ > 0 Then Debug.Print JJ
        
        
If IsIDE = True Then
   ' If FrmJoy.Visible = False Then Me.Show: DoEvents
    For N = 0 To 9
        If be.B(N + 1) = True Then
            Label1(N).BackColor = Red
        
        Else
        Label1(N).BackColor = White
        End If
    Next
End If
        
        
Dim Vertical_Pos As Long


notepad2rt = FindWindow("#32770", "Notepad2")
'notepad3rt = FindWindow("Notepad2", vbNullString)
'xxrp = GetParent(notepad2rt)
'cuecue$ = GetWindowTitle(notepad3rt)
'If xxrp <> notepad3rt Then Exit Sub

'If notepad2rt = 0 Then Exit Sub

'If NotePad2RT2 = notepad3rt Then Exit Sub

If GetForegroundWindow = notepad2rt Then yesnote = True


If FindWindow(vbNullString, "Tidal Information...") <> GetForegroundWindow Then
If (JJ = 3 And Me.WindowState = 0) Or (yesnote = True And JJ = 3) Then
    'Vertical_Pos = GetVerticalScrollPos(Form1.Text1)
    'Call SetVerticalScrollPos(Form1.Text1, Vertical_Pos + 10)
    SendKeys "{up}"
    HHH = Now + TimeSerial(0, 10, 0)
    Exit Sub
End If
If (JJ = 4 And Me.WindowState = 0) Or (yesnote = True And JJ = 4) Then
'    Vertical_Pos = GetVerticalScrollPos(Form1.Text1) - 100
    'Vertical_Pos = 100
'    Call SetVerticalScrollPos(Form1.Text1, Vertical_Pos)
    SendKeys "{down}"
    HHH = Now + TimeSerial(0, 10, 0)
    Exit Sub
End If
End If
     
                    
'xxrp = GetParent(notepad2rt)
'cuecue$ = GetWindowTitle(notepad3rt)

If GetForegroundWindow = FindWindow("Outlook Express Browser Class", vbNullString) Then yesnote = True

If GetForegroundWindow = FindWindow("ATH_Note", vbNullString) Then yesnote = True: XYN = True Else XYN = False

If JJ = 3 And yesnote = True Then
    'Vertical_Pos = GetVerticalScrollPos(Form1.Text1)
    'Call SetVerticalScrollPos(Form1.Text1, Vertical_Pos + 10)
    SendKeys "{up}", False
    Exit Sub
End If
        
If JJ = 4 And yesnote = True Then
'    Vertical_Pos = GetVerticalScrollPos(Form1.Text1) - 100
    'Vertical_Pos = 100
'    Call SetVerticalScrollPos(Form1.Text1, Vertical_Pos)
    SendKeys "{down}", False
    Exit Sub
End If
        
                    
        
                    
Exit Sub
                    
                    


'''' These One Press Only
'Change Test only if button press any combination but not releases Harry
'My Button got released do you like that
TT = 0
For N = 0 To 9
    TT = TT + Abs(be.B(N + 1)) * (N + 1)
    'If be1.B(n + 1) = True Then Stop
Next
TT = TT + JJ * 11
If TT > OTT Then
JoyPress = True
Else: JoyPress = False
End If

OTT = TT
If JoyPress = False Then Exit Sub

If be.B(1) = True Then

TTB = TTB + 1
If TTB > 6 Then TTB = 1

If TTB = 1 Then
    On Error Resume Next
        Form1.WindowState = 0
    On Error GoTo 0
    Call Form1.Mnu_Max_Click
End If

If TTB = 2 Then
    On Error Resume Next
        Form1.WindowState = 1
    On Error GoTo 0
    Call Form1.Mnu_Max_Click
End If

If TTB = 3 Then
    Form1.WindowState = 1
    ED = FindWindow("Notepad2", vbNullString)
    If ED > 0 Then
        ShowWindow ED, SW_NORMAL
        SetForegroundWindow (ED)
    Else
    Shell "NOTEPAD.EXE ""D:\VB6\VB-NT\00_Best_VB_01\NNTP\USENET POSTS -- alt.philosophy.txt"""
    
    End If

End If

If TTB = 4 Then
    Form1.WindowState = 1
    ED = FindWindow("Notepad2", vbNullString)
    If ED > 0 Then
        ShowWindow ED, SW_MINIMIZE
        'SetForegroundWindow (ED)
    End If

End If


If TTB = 5 Then
    Form1.WindowState = 1
    ED = FindWindow("Outlook Express Browser Class", vbNullString)
    If ED > 0 Then
        ShowWindow ED, SW_NORMAL
        SetForegroundWindow (ED)
    Else
    Shell "C:\Program Files\OE-QuoteFix\OELaunch.exe", vbNormalFocus
    End If
End If

If TTB = 6 Then
    Form1.WindowState = 1
    ED = FindWindow("Outlook Express Browser Class", vbNullString)
    If ED > 0 Then
        ShowWindow ED, SW_MINIMIZE
        'SetForegroundWindow (ED)
    End If
End If




End If


'BUTT 1 - AND THEN IF ON NOTE MEAN FULL SCREEN MSG
'Open MSG
'OUTLOOK EXPRESS
If be.B(1) = True And yesnote = True And XYN = False Then
'    SendKeys "^o", False
    Exit Sub
End If

'ESC Msg
'OUTLOOK EXPRESS
If be.B(1) = True And yesnote = True And XYN = True Then
'    SendKeys "{esc}", False
    Exit Sub
End If
    

'IN MSGS Not The Post
'UP
'OutLook Express
If JJ = 6 And yesnote = True Then
    'SendKeys "{up}", False
    SendKeys "^<", False
    
    Exit Sub
End If
        
'DOWN
If JJ = 8 And yesnote = True Then
    'SendKeys "{down}", False
    SendKeys "^>", False
    Exit Sub
End If
        



End Sub

Private Sub Form_Resize()
Me.Hide
End Sub



Private Sub Timer_XxXJOY_Timer()
Dim Handle
'Timer1.Enabled = False
'Exit Sub

I = FindWindow("Winamp v1.x", vbNullString)
If I <> OLDI Then
    OLDI = I
    'ADD NEW ITEM
    List1.AddItem (Str(I))
    
    'MAKE LIST
    For r = List1.ListCount - 1 To 1 Step -1
        Handle = List1.List(r)
        
        TTX1 = GetWindowTitle(Val(Handle))
        If TTX1 = "" Then
            List1.RemoveItem (r)
        End If
            
    Next
    
    'DUPE COPY REMOVED
    For r = List1.ListCount - 1 To 1 Step -1
        If List1.List(r) = List1.List(r + 1) Then List1.RemoveItem (r)
    Next
    
End If

End Sub
