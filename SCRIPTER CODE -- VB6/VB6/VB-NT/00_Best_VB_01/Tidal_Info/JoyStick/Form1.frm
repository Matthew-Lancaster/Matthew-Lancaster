VERSION 5.00
Begin VB.Form FrmJoy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Joystick Example..."
   ClientHeight    =   4764
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   3552
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4764
   ScaleWidth      =   3552
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1575
      Top             =   1770
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      DrawWidth       =   5
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2355
      ScaleMode       =   0  'User
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   9
      Left            =   1560
      TabIndex        =   10
      Top             =   3585
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   960
      TabIndex        =   9
      Top             =   3585
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   2160
      TabIndex        =   8
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   1560
      TabIndex        =   7
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   960
      TabIndex        =   6
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   5
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   2160
      TabIndex        =   4
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   1560
      TabIndex        =   3
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   2640
      Width           =   375
   End
End
Attribute VB_Name = "FrmJoy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FORM NAME - FrmJoy

Dim SMX, SMY, OX1, OY1, GF, JJ, OJJ, OKJ, KeyBounce, OJKY, OB1, OBT1, OBT2, JoyQuickTrig
Dim HALO2, TimeWait, OldSluty3, OTT, JoyPress, PY, x1, y1
Dim jold As JOYINFO, XJJ
Const Red = &HFF&
Const White = &HFFFFFF

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Type POINTAPI
        x As Long
        Y As Long
End Type

Private Sub Form_Load()
If IsWinNT Then
    joyGetDevCapsW 0, Je, Len(Je)
Else
    joyGetDevCapsA 0, Je, Len(Je)
End If
If Je.wXmax > 0 Then Picture1.ScaleHeight = Je.wXmax
If Je.wYmax > 0 Then Picture1.ScaleWidth = Je.wYmax
Picture1.ScaleLeft = 0
Picture1.ScaleTop = 0
'FrmJoy.Show
Timer1.Interval = 10
End Sub


Private Sub Form_Unload(Cancel As Integer)

If aa_MainCode.ALL_FORM_REQUEST_TO_END = False Then
    Cancel = True: Me.Hide: Exit Sub
End If

End Sub

Public Sub Timer1_Timer()
'FRMJOY.TIMER1_TIMER
'Timer1.Interval

'DEFAULT =SET 10
'SET QUICKER NOT IN IDE

'If Timer1.Interval = 10 Then
'    If IsIDE = False Then
'        Timer1.Interval = 1
'    Else
'        Timer1.Interval = 11
'    End If
'End If

If Timer1.Interval <> 1 Then Timer1.Interval = 1



GetBPress j.Buttons, be

If be.B(1) = True Then BE_B_01 = True
If be.B(2) = True Then BE_B_02 = True
If be.B(3) = True Then BE_B_03 = True
If be.B(4) = True Then BE_B_04 = True
If be.B(5) = True Then BE_B_05 = True
If be.B(6) = True Then BE_B_06 = True
If be.B(7) = True Then BE_B_07 = True
If be.B(8) = True Then BE_B_08 = True
If be.B(9) = True Then BE_B_09 = True
If be.B(10) = True Then BE_B_10 = True


'joyGetPos 0, j'LEFT JOY STICK

'Do
'DoEvents
joyGetPos 0, j
        
'Debug.Print j.X
'Loop Until 1 = 2
        
'Updated all this pucker on 28-07-07 3:56pm Knackarded Working Staving
        
'KEY1 = Int(32767 / 1.2)
'KEY2 = Int(65535 / 1.2)

'= 255   LOW
'= 32767 MIDDLE
'= 65275 HIGH

KEY1 = 1000 'Int(32767 / 1.5)
KEY2 = 50000 ' Int(65535 / 1.5)

'Debug.Print str(j.X) + " ---- " + str(j.Y)
'----------

'MOST LIKELY JOYPAD NOT IN IF HERE IS NAUGHT
'WANT THIS FOR LOAD DXDIRECT DRIVER
If j.x = 0 And j.Y = 0 Then
    JOYPAD_SOCKETED_IN = False
End If

If JOYPAD_SOCKETED_IN = False Then
    JOYPAD_SOCKETED_IN = True
    If frmDXGAMEJOY.JOYPAD_DX_LOADED = False Then
        Load frmDXGAMEJOY
    End If
End If

JJ = 0: JJ2 = 0
If j.x = 32767 And j.Y = 32511 Then JJ = 0 'Idle
'----------
'Debug.Print j.Y
If j.x = 0 And j.Y > KEY1 Then JJ = 2 'Left


'LEFT JOY THUMB PAD COMPUS
'Debug.Print j.x, j.Y
If j.x < KEY1 Then
    JJ = 1: JJ2 = 2  'Left
End If
If XJJ <> JJ2 And JJ2 = 2 Then
    'SayTime3 = True
End If
XJJ = JJ2

If j.x > KEY1 And j.Y < KEY1 Then JJ = 3 'Up
If j.x > KEY1 And j.Y > KEY2 Then JJ = 4 'Down
If j.x > KEY2 And j.Y > KEY1 Then JJ = 5 'Right
'----------
If j.x < KEY1 And j.Y < KEY1 Then JJ = 6 'Left Up -- Top Left
If j.x > KEY2 And j.Y < KEY1 Then JJ = 7 'Up Right -- Top Right
If j.x < KEY1 And j.Y > KEY2 Then JJ = 8 'Down Left  Bottom Left
If j.x > KEY2 And j.Y > KEY2 Then JJ = 9 'Right Down -- Bottom Right
        
        
'-------------------
'DRAW THE BUTTON SET
'-------------------
If FrmJoy.Visible = True Then
    For n = 0 To 9
        If be.B(n + 1) = True Then
            Label1(n).BackColor = Red
        Else
        Label1(n).BackColor = White
        End If
    Next
End If
                    
' One Press Only
' Change Test only if button press any combination but not releases Harry
' My Button got released do you like that
TT = 0
For n = 0 To 9
    TT = TT + Abs(be1.B(n + 1)) * (n + 1)
    'If be1.B(n + 1) = True Then Stop
Next
TT = TT + JJ * 11

If TT > OTT Then
    JoyPress = True
Else
    JoyPress = False
End If
OTT = TT

If JoyPress = False Then
    JOY_PREVIOUS_TRACK_BUTTON_GO = False
    JOY_NEXT_TRACK_BUTTON_GO = False
End If

''These QUICKER key
'TOP AND BOTTOM OTHER WAY ROUND - TOP MUSIC FORWARD BACK
'AND BOTTOM VOLUME
'If JOY_CONTROLLER = "" Or JOY_CONTROLLER <> "" Then
'    If be.B(6) Then 'And PreviNext = False Then
'        'Fast ForWard
'        Call WinAmpCommands(4)
'    End If
'
'    If be.B(5) Then
'        'ReWind
'        'Sluty3 = 3
'        Call WinAmpCommands(5)
'    End If
'End If

GAMEPAD_GO = False
If JOY_CONTROLLER = "Logitech Cordless RumblePad 2" Then
    GAMEPAD_GO = True
End If
If JOY_CONTROLLER = "Controller (Wireless Gamepad F710)" Then
    GAMEPAD_GO = True
End If

If GAMEPAD_GO = True Then
    If MOTION_FAST_FORWARD_BUTTON_GO = True Then
        Call WinAmpCommands(4)
    End If
    If MOTION_REVERSE_BUTTON_GO = True Then
        Call WinAmpCommands(5)
    End If
    
    If MOTION_FAST_FORWARD_BUTTON_GO = True Then
        Call WinAmpCommands(4)
    End If
    If MOTION_REVERSE_BUTTON_GO = True Then
        Call WinAmpCommands(5)
    End If
End If

If JOY_CONTROLLER = "Logitech Cordless RumblePad 2" Then
    If JOY_VOLUME_UP_BUTTON_GO = True Then
        'Vol Up
        'Sluty3 = 10
        'MUST SET HERE AS WHY WE THINK LAB CHANGE WE SET
        oVolumeSet2 = Val(ATidalX.LblVol) + 1
        ATidalX.LblVol = str(Val(ATidalX.LblVol) + 1) 'Raise Vol)
    
    End If
    If JOY_VOLUME_DOWN_BUTTON_GO = True Then
        'Vol Down
        'Sluty3 = 11
        'MUST SET HERE AS WHY WE THINK LAB CHANGE WE SET
        oVolumeSet2 = Val(ATidalX.LblVol) - 1
        ATidalX.LblVol = str(Val(ATidalX.LblVol) - 1)  'Lower Vol)
    
    End If
End If

If JOY_CONTROLLER = "Controller (Wireless Gamepad F710)" Then
    If JOY_VOLUME_UP_BUTTON_GO = True Then
        'Vol Up
        'Sluty3 = 10
        'MUST SET HERE AS WHY WE THINK LAB CHANGE WE SET
        oVolumeSet2 = Val(ATidalX.LblVol) + 1
        ATidalX.LblVol = str(Val(ATidalX.LblVol) + 1) 'Raise Vol)
    
    End If
    If JOY_VOLUME_DOWN_BUTTON_GO = True Then
        'Vol Down
        'Sluty3 = 11
        'MUST SET HERE AS WHY WE THINK LAB CHANGE WE SET
        oVolumeSet2 = Val(ATidalX.LblVol) - 1
        ATidalX.LblVol = str(Val(ATidalX.LblVol) - 1)  'Lower Vol)
    
    End If
End If


'If JOY_CONTROLLER = "" Then
'    If be.B(8) Then
'        'Vol Up
'        'Sluty3 = 10
'        'MUST SET HERE AS WHY WE THINK LAB CHANGE WE SET
'        oVolumeSet2 = Val(ATidalX.LblVol) + 1
'        ATidalX.LblVol = str(Val(ATidalX.LblVol) + 1) 'Raise Vol)
'
'    End If
'    If be.B(7) Then
'        'Vol Down
'        'Sluty3 = 11
'        'MUST SET HERE AS WHY WE THINK LAB CHANGE WE SET
'        oVolumeSet2 = Val(ATidalX.LblVol) - 1
'        ATidalX.LblVol = str(Val(ATidalX.LblVol) - 1)  'Lower Vol)
'
'    End If
'End If

    
'If JoyPress = True Then
'    If JOY_CONTROLLER = "" Then ' Or JOY_CONTROLLER <> "" Then
'        If be.B(4) = True Then
'            'Previous Track
'            If JOY_PREVIOUS_TRACK_BUTTON_GO = False Then
'                JOY_PREVIOUS_TRACK_BUTTON_GO = True
'                Call WinAmpCommands(6)
'            End If
'        End If
'    End If
'
'    If JOY_CONTROLLER = "" Then
'        If be.B(2) = True Then
'            'Next Track
'            If JOY_NEXT_TRACK_BUTTON_GO = False Then
'                JOY_NEXT_TRACK_BUTTON_GO = True
'                Call WinAmpCommands(3)
'            End If
'        End If
'    End If

'    If JOY_CONTROLLER = "Controller (Wireless Gamepad F710)" Then
'        If be.B(1) = True Then
'            'Next Track
'            If JOY_NEXT_TRACK_BUTTON_GO = False Then
'                JOY_NEXT_TRACK_BUTTON_GO = True
'                Call WinAmpCommands(3)
'            End If
'        End If
'    End If
    

'    If JOY_CONTROLLER = "Logitech Cordless RumblePad 2" Then
'        If JOY_PAUSE_BUTTON_GO = True And O_JOY_PAUSE_BUTTON_GO <> JOY_PAUSE_BUTTON_GO Then
'            Call WinAmpCommands(2)
'        End If
'    End If
'
'    If JOY_CONTROLLER = "Controller (Wireless Gamepad F710)" Then
'        If JOY_PAUSE_BUTTON_GO = True And O_JOY_PAUSE_BUTTON_GO <> JOY_PAUSE_BUTTON_GO Then
'            Call WinAmpCommands(2)
'        End If
'    End If
'    O_JOY_PAUSE_BUTTON_GO = JOY_PAUSE_BUTTON_GO

'    If JOY_CONTROLLER = "" Then
'        If be.B(3) Then
'            'Pause
'            Call WinAmpCommands(2)
'        End If
'    End If

'    If JOY_CONTROLLER = "Controller (Wireless Gamepad F710)" Then
'        If be.B(2) Then
'            'Pause
'            Call WinAmpCommands(2)
'        End If
'    End If
'End If


If 1 = 2 Then
    If JoyPress = True And be1.B(10) = True Then
    
        If OKJ <> be1.B(10) Then
            'Call CenterCursor
        End If
        
        'If GF = False Then Call CenterBot
        'GF = True
        
        If JJ = 4 Then
            Call CenterBot
        End If
        If JJ = 3 Then
            Call CenterTop
        End If
        If JJ = 5 Then
            Call CenterRight
        End If
        If JJ = 2 Then
            Call CenterLeft
        End If
        '----------------- Breathe
        If JJ = 8 Then
            Call DiagBotLeft
        End If
        If JJ = 6 Then
            Call DiagTopLeft
        End If
        If JJ = 7 Then
            Call DiagTopRight
        End If
        If JJ = 9 Then
            Call DiagBotRight
        End If
    End If
End If

If be1.B(10) = False Then GF = False

For n = 0 To 9
    be1.B(n + 1) = be.B(n + 1)
Next

Exit Sub



If WAFastFF = True Then
    'problem here i think when keys use callback stop even this buffer results show the key stopped
    'notepad dont like f4 we block that dont see a buffer over run in this anyhow
    'okay we need tis in diff place bit more befor hand
    'bdf1 = GetAsyncKeyState(115)
    'If bdf1 = 0 Then WAFastFF = 0: Debug.Print Str(Now)
    If WAFastFF = True Then WAFastFF = False: Call WinAmpCommands(4)
End If



If Sluty3 > 0 And OldSluty3 = 0 Then TimeWait = 0
If Sluty3 > 0 And OldSluty3 > 0 Then
    If Sluty3 <> OldSluty3 Then TimeWait = 0
End If
        
OldSluty3 = Sluty3

If (Sluty3 > 0 And Val(ATidalX.LblVol) < 8) And TimeWait > Timer Then
    Sluty3 = 0
End If

If Sluty3 > 0 Then
    TimeWait = Timer + 0.1
    Call ATidalX.VolumeLevelSet
End If


Exit Sub



    If j.x = 0 And j.Y = 32511 Then
        '#JoyInfo Left
        'Close Winamp
        'use whatever one on top for close video wanted
        WINAMP2 = FindWindow("Winamp v1.x", vbNullString)
        If WINAMP2 > 0 Then
            'Close WinAmp
            'MsgResult = SendMessage(WinAmp2, WM_CLOSE, , ByVal XcNul)
            'Tx = CloseWindow(WinAmp2)
            'Tx = CloseHandle(WinAmp2)
            'Call SendKey(70, 4) '# send Alt F4 get rigth keys this is not
    
        End If
    End If
            
    

End Sub


Sub None()
Exit Sub
If vcode = 121 Then Sluty3 = 4    'F10 = 'Vol Up
If vcode = 120 Then Sluty3 = 5  'F9 = 'Vol Down
If vcode = 119 Then Sluty3 = 2 'F8 = Pause
If vcode = 118 Then Sluty3 = 3 'F7 = Next Track


'on Mobile Remote
If Actionz = 2 And Statez = 6 And vcode = 50 Then Sluty3 = 3 'F7 = Next Track
If Actionz = 2 And Statez = 6 And vcode = 49 Then Sluty3 = 2 'F8 = Pause???? no 118=f8

 'left 0 ---  32511
 'up 32511 ---  0
 'down 32511 ---  65535
 'right 65535 ---  32511
 'none 32511 ---  32511



End Sub


Sub CenterCursor()
If PY > 4 Then PY = 1

Call CenterCur

'If PY = 1 Then Call CenterCur

'If PY = 2 Then Call CenterBot
'If PY = 3 Then Call CenterTop
'If PY = 4 Then Call BottomLeft
'If PY = 5 Then Call BottomMid
'If PY = 6 Then Call BottomRight
'If PY = 7 Then Call TopLeft
'If PY = 8 Then Call TopMid
'If PY = 9 Then Call TopRight

PY = PY + 1
PY = 1


End Sub
Public Sub CenterCur()
y1 = Screen.Height
x1 = Screen.Width
'FrmJoy.Show
'Me.DrawWidth = 10
'On Error Resume Next
'For Each Control In Controls
'Control.Visible = False
'Control.Top = 20000
'Next
'On Error GoTo 0
'Me.Height = Y1
'Me.Width = X1
'Me.Top = 0
'Me.Left = 0

'Do
'Me.Cls
yy1 = 0: XX1 = 0

Dim P As POINTAPI

SMX = (GetSystemMetrics(0) / Screen.Width)
SMY = (GetSystemMetrics(1) / Screen.Height)

step1 = 10
For rx = 0 To y1 Step step1 '(Y1 / 1.7) / step1
    
    OX1 = x1 / 2
    OY1 = y1 / 2
    ry = step1 * (y1 / x1)
    XX1 = XX1 + step1
    yy1 = (yy1 + ry)
    
    A1 = OX1 + XX1: B1 = OY1 + yy1
'    Me.PSet (a1, b1)
    A1 = A1 * SMX: B1 = B1 * SMY
    dar = SetCursorPos(A1, B1)
    
    A1 = OX1 + XX1: B1 = OY1 + yy1 * -1
'    Me.PSet (a1, b1)
    A1 = A1 * SMX: B1 = B1 * SMY
    dar = SetCursorPos(A1, B1)
    
    A1 = OX1 + XX1 * -1: B1 = OY1 + yy1 * -1
'    Me.PSet (a1, b1)
    A1 = A1 * SMX: B1 = B1 * SMY
    dar = SetCursorPos(A1, B1)
    If A1 < 125 Then
    Exit For
    End If
    
    A1 = OX1 + XX1 * -1: B1 = OY1 + yy1
'    Me.PSet (a1, b1)
    A1 = A1 * SMX: B1 = B1 * SMY
    dar = SetCursorPos(A1, B1)
Next
'Exit Sub
'DoEvents
'Sleep 1000
'Loop Until 1 = 2
    
End Sub
Public Sub CenterBot()
    A1 = (x1 / 2) * SMX: B1 = y1 * SMY
    dar = SetCursorPos(A1, B1)
End Sub
Public Sub CenterTop()
    A1 = (x1 / 2) * SMX: B1 = 0 * SMY
    dar = SetCursorPos(A1, B1)
End Sub
Public Sub CenterRight()
    A1 = x1 * SMX: B1 = (y1 / 2) * SMY
    dar = SetCursorPos(A1, B1)
End Sub
Public Sub CenterLeft()
    A1 = 0 * SMX: B1 = (y1 / 2) * SMY
    dar = SetCursorPos(A1, B1)
End Sub
Public Sub DiagBotRight()
    A1 = x1 * SMX: B1 = y1 * SMY
    dar = SetCursorPos(A1, B1)
End Sub
Public Sub DiagTopRight()
    A1 = x1 * SMX: B1 = 0 * SMY
    dar = SetCursorPos(A1, B1)
End Sub
Public Sub DiagTopLeft()
    A1 = 0 * SMX: B1 = 0 * SMY
    dar = SetCursorPos(A1, B1)
End Sub
Public Sub DiagBotLeft()
    A1 = 0 * SMX: B1 = y1 * SMY
    dar = SetCursorPos(A1, B1)
End Sub

