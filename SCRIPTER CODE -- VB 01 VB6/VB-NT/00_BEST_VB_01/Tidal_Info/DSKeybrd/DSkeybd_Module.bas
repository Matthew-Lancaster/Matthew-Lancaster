Attribute VB_Name = "DSkeybd_M"
'Timer13_Timer() ------- Look for this one for sendkeys on progs

Option Explicit

Public MAIN_FORM_READY_YET

Dim TXXRRS
Dim TXXRRS_CHANGE_VAR

Dim TEXTISVB
Dim FwnD


Public SayTimeVolAdjustLateNite
Public SayTime3
Public MEHwnD, TestHookinVB


Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Public KeyIdleTime As Date, Xxr5, Flag22

Public LockDontAllowHotKeyforPassword
Public KeyLoggDate
Public TestKeyLoggOff

Public RecordWindow, HittMan, HittMan2, HottMini
Public Tiger As Long
Public KeyLog2
Public GoThrou, VolumeSet1, Go8
Public Action, State, vcode, Scode
Public Actionz, Statez, Vcodez, Scodez
Public KeyPlus
Public SndKey
Public RKeyPlus ', KeyPlus2
Public TimeTrigSay

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, _
     ByVal lpWindowName As String) _
    As Long



'***********************************************
'     Sample for using the "dskbhook.dll"
'        (c) 2002 by Delphin Software
'***********************************************
'Option Explicit
'***********************************************
Public Const KEYEVENTF_KEYUP& = &H2
Public Const KEYEVENTF_EXTENDEDKEY& = &H1
'-----------------------------------------------
Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags&, ByVal dwExtraInfo&)
'-----------------------------------------------
Declare Function SetKeyboardHook& Lib "dskbhook" (ByVal hwnd&, ByVal Callback&, ByVal Adr&)
Declare Function SetGlobalParams& Lib "dskbhook" (ByVal Repeat&, ByVal Discard&)
Declare Function SetKeyParams& Lib "dskbhook" (ByVal KeyCode&, ByVal SCAE&, ByVal RDC&)
Declare Function UseSendMessage& Lib "dskbhook" (ByVal Send&)
Declare Function SkipInjected& Lib "dskbhook" (ByVal skip&)
Declare Function NoLLHook& Lib "dskbhook" (ByVal LLHook&)
Declare Function SetStop& Lib "dskbhook" (ByVal Stp&)
Declare Function HookVB& Lib "dskbhook" (ByVal Hook&)
Declare Function NoTL& Lib "dskbhook" (ByVal tl&)
Declare Function IsNT& Lib "dskbhook" ()
'***********************************************
Public mode%, IsHook As Boolean
Public Shift As Boolean, Ctrl As Boolean, LeftAlt As Boolean, RightAlt As Boolean, AnyAlt As Boolean
Public Ext As Boolean, Up As Boolean, Down As Boolean, NoStateKey


' Virtual Keys, Standard Set
Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2
Public Const VK_CANCEL = &H3
Public Const VK_MBUTTON = &H4             '  NOT contiguous with L RBUTTON

Public Const VK_BACK = &H8
Public Const VK_TAB = &H9

Public Const VK_CLEAR = &HC
Public Const VK_RETURN = &HD

Public Const VK_SHIFT = &H10
Public Const VK_CONTROL = &H11
Public Const VK_MENU = &H12
Public Const VK_PAUSE = &H13
'Public Const VK_CAPITAL = &H14

Public Const VK_ESCAPE = &H1B

'###
Public Const VK_SPACE = &H20
Public Const VK_PRIOR = &H21
Public Const VK_NEXT = &H22
Public Const VK_END = &H23
Public Const VK_HOME = &H24
Public Const VK_LEFT = &H25
Public Const VK_UP = &H26
Public Const VK_RIGHT = &H27
Public Const VK_DOWN = &H28
Public Const VK_SELECT = &H29
Public Const VK_PRINT = &H2A
Public Const VK_EXECUTE = &H2B
Public Const VK_SNAPSHOT = &H2C
Public Const VK_INSERT = &H2D
Public Const VK_DELETE = &H2E
Public Const VK_HELP = &H2F

' VK_A thru VK_Z are the same as their ASCII equivalents: 'A' thru 'Z'
' VK_0 thru VK_9 are the same as their ASCII equivalents: '0' thru '9'

Public Const VK_NUMPAD0 = &H60
Public Const VK_NUMPAD1 = &H61
Public Const VK_NUMPAD2 = &H62
Public Const VK_NUMPAD3 = &H63
Public Const VK_NUMPAD4 = &H64
Public Const VK_NUMPAD5 = &H65
Public Const VK_NUMPAD6 = &H66
Public Const VK_NUMPAD7 = &H67
Public Const VK_NUMPAD8 = &H68
Public Const VK_NUMPAD9 = &H69
Public Const VK_MULTIPLY = &H6A
Public Const VK_ADD = &H6B
Public Const VK_SEPARATOR = &H6C
Public Const VK_SUBTRACT = &H6D
Public Const VK_DECIMAL = &H6E
Public Const VK_DIVIDE = &H6F
Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73 '115
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B
Public Const VK_F13 = &H7C
Public Const VK_F14 = &H7D
Public Const VK_F15 = &H7E
Public Const VK_F16 = &H7F
Public Const VK_F17 = &H80
Public Const VK_F18 = &H81
Public Const VK_F19 = &H82
Public Const VK_F20 = &H83
Public Const VK_F21 = &H84
Public Const VK_F22 = &H85
Public Const VK_F23 = &H86
Public Const VK_F24 = &H87

Public Const VK_NUMLOCK = &H90
Public Const VK_SCROLL = &H91
Public Const VK_CAPITAL = &H14

'
'   VK_L VK_R - left and right Alt, Ctrl and Shift virtual keys.
'   Used only as parameters to GetAsyncKeyState() and GetKeyState().
'   No other API or message will distinguish left and right keys in this way.
'  /
Public Const VK_LSHIFT = &HA0
Public Const VK_RSHIFT = &HA1
Public Const VK_LCONTROL = &HA2
Public Const VK_RCONTROL = &HA3
Public Const VK_LMENU = &HA4
Public Const VK_RMENU = &HA5

Public Const VK_ATTN = &HF6
Public Const VK_CRSEL = &HF7
Public Const VK_EXSEL = &HF8
Public Const VK_EREOF = &HF9
Public Const VK_PLAY = &HFA
Public Const VK_ZOOM = &HFB
Public Const VK_NONAME = &HFC
Public Const VK_PA1 = &HFD
Public Const VK_OEM_CLEAR = &HFE




'***********************************************
'# This function is called from the DLL in general
Public Function Callback&(ByVal Action&, ByVal State&, ByVal vcode&, ByVal Scode&)

'Don't get caught in the IDE Design Code
If IsVBrunningForCode = True Then
    
    'GO CODE WRONG THIS NEED RUN ALL THE TIME
    'NOT JUST ONLY AT FIRST KEY PRESS
    'WILL RUN ALWAYS WHEN A KEY PRESS
    
    ATidalX.Timer_REINTRODUCE_DSkeybd_m.Enabled = True
    
    'WRONG DON'T WANT CALLBACK VAR = 2 BECAUSE THAT BLOCK KEY
    'Callback = 2:
    'TRY AND UNDERSTAND WHY WAS OKAY IN IS-IDE
    
    Exit Function
    
End If

On Local Error GoTo ErrTrap

'Check = SlowKeyPress()

If vcode = 0 Then vcode = 93: Action = 2 'Test


If KeyClean = True Then
    Action = 0
    State = 0
    vcode = 0
    Scode = 0
    Callback = 2: Exit Function
End If



'If TestKeyLoggOff = False Then
'# Display the passed values
With DSkeybd_F.List1
  'If .ListCount > 500 Then .Clear
  '# Add the new entry at the first position and highlight it
'  .AddItem "   " & Action & Chr$(9) & State & Chr$(9) & VCode & Chr$(9) & Scode, 0
  .AddItem Format$(Action, "000 ") & Format$(State, "000 ") & Format$(vcode, "000 ") & Format$(Scode, "000 ")
  .Selected(.NewIndex) = True
    DSkeybd_F.Label5 = .ListCount

End With
'End If

'KeyLog2 = KeyLog2 + Format$(Now, "DDD DD-MMM-YY HH:MM:SS") + " - KEY =" + Str(Action) + Chr$(9) + Str(State) + Chr$(9) + Str(Vcode) + Chr$(9) + Str(Scode) + vbCrLf

'DPrint action, state, vcode, scode
Down = Action And 2: Up = Action And 4: Ext = Action And 1
Ctrl = False: LeftAlt = False: RightAlt = False: AnyAlt = False: NoStateKey = False
Shift = State And 1


'Call KeyStateTest(vcode)

'If CapsStat = False Then If (State And 1) = 1 Then Shift = False: State = 0


If (State And 2) = 2 Then Ctrl = True
If (State And 3) = 3 Then Ctrl = True
If (State And 4) = 4 Then LeftAlt = True: AnyAlt = True
If (State And 5) = 5 Then LeftAlt = True: AnyAlt = True
If (State And 6) = 6 Then RightAlt = True: AnyAlt = True
If (State And 7) = 7 Then RightAlt = True: AnyAlt = True

If AnyAlt = False And Ctrl = False And Shift = False Then NoStateKey = True Else NoStateKey = False


Callback = 0 '#never suppress any keys yet
    


'If SndKey = True Then SndKey = False: Callback = 2: Exit Function
    
KeyIdleTime = Now + TimeSerial(0, 0, 30)
    
Actionz = Action
Statez = State
Vcodez = vcode
Scodez = Scode

'Kbbdf5 = False

'If KeyPlus2 = True Then KeyPlus = True

'If Not (Action = 4 Or Action = 5) And IsVBrunning = 0 Then
'    If State = 2 And vcode = 38 Then  'Up Arrow Mobile
'        If Menu.Slider1.Value - 1 > 0 Then
'            Menu.Slider1.Value = Menu.Slider1.Value - 1
'        End If
'        Menu.Show
'        'Exit Function
'    End If
'    If State = 2 And vcode = 40 Then  'Down Arrow Mobile
'        If Menu.Slider1.Value - 1 < 100 Then
'            Menu.Slider1.Value = Menu.Slider1.Value + 1
'        End If
'        Menu.Show
'        'Exit Function
'    End If
'End If

ATidalX.Label15.Caption = Action & "-" & State & "-" & vcode & "-" & Scode

'If SndKey = True Then SndKey = False: Callback = 2: Exit Function
'If (action = 4 Or action = 5) And vcode = 107 Then Callback = 2: Exit Function

'Why stop this key
'This is the Key bottom right night to control inbetween windows key an control
'If vcode = 93 Then
'    Callback = 2
'End If

If vcode = 123 Then
    Callback = 2
End If

If Action = 4 Or Action = 5 Then Exit Function 'key release

'F4 = Fast F
'Timer
'This code Bring up th ebuffer of keys but if winamp is paused in song an FF then dont
'shows winamp is taking CPU - even happens with our timer
If vcode = 115 Then
    WAFastFF = True
    ATidalX.WinampFastForwardTimer.Enabled = True
End If

'#### DONT ALLOW F12'S ON ANYTHING NO MORE



'Call ATidalX.TimerKeyInterceptSlow_Timer

'Call KeyStateTest(Vcode)

' VK_A thru VK_Z are the same as their ASCII equivalents: 'A' thru 'Z'
' VK_0 thru VK_9 are the same as their ASCII equivalents: '0' thru '9'

VCodeTT_ = VCodeTT_ + Chr$(vcode)

Dim IsVBrunningForCode2
IsVBrunningForCode2 = IsVBrunningForCode

If IsVBrunningForCode2 = True And vcode = 115 Then Callback = 2

If Statez <> 2 And IsVBrunningForCode2 = True Then
    If vcode = 119 Or vcode = 120 Or vcode = 121 Then vcode = 0: Vcodez = 0
End If

If vcode = 145 Then
    'Call ATidalX.LABEL_TIME_Click
    'SayTime3 = True
    'Call ATidalX.LABEL_TIME_Click
End If









'If Actionz = 3 And (Statez = 0 Or Statez = 2) And Vcodez = 38 Then 'Up Arrow Mobile
'    Call KnockerLogg
'End If

'If Actionz = 2 And Vcodez = 106 Then '* Num Pad
    'Call KnockerLogg
'End If

'If kbi = 117  -- 'F6 = 1 Min delay Pause
'If Actionz = 2 And Statez = 0 And Vcodez = 117 Then MinDelayPause = Now + TimeSerial(0, 30, 0)

'If Actionz = 2 And Statez = 0 And Vcodez = 122 Then Sluty4 = 1      'F11 = Mark Messages By Conversation
Tiger = FindWindow("Outlook Express Browser Class", vbNullString)
If Tiger > 0 And GetWindowState(Tiger) <> vbMinimized Then
'   If Vcodez = 109 Then - minus key num pad
    If Vcodez = 122 Then 'f11
        Sluty4 = 1      '- Num Pad = Mark Messages By Conversation
'        sdhc = 1
    End If
End If


'F11
If vcode = 122 Then
    TYYG2 = Now - 1
End If

'F7
If IsVBrunningForCode2 = True And Vcodez = 118 Then Callback = 2: Exit Function


Dim TigerX, XTiger

TigerX = FindWindow("ATH_Note", vbNullString)
'F7
If TigerX <> 0 And Vcodez = 118 Then Callback = 2: Exit Function




'TestKeyLoggOff = True
'search

'This Is Now Ran From Inet Explore pass word Entry If Not Used Then This Called as Normal

'F12 Timer
If vcode = 123 And Sluty = 0 Then
   'ATidalX.ShellKeyLoad.Enabled = True
End If
        
        
        
'F12 Timer
If Ctrl = True And vcode = 123 Then
    ATidalX.ShellKeyLoad.Enabled = True
End If
        
'Dictation
'F12
'Timer
If vcode = 123 Then
    Tiger = FindWindow("#32770", "Recording Window")
    If Tiger = 0 Then
        Tiger = FindWindow("Digital Wave Player", vbNullString)
        If Tiger > 0 And WindowVisible(Tiger) = True Then
            WindowVisible(Tiger) = False
            'ShowWindow tiger, SW_MINIMIZE
            Sluty = 0
            ATidalX.ShellKeyLoad.Enabled = False
        End If
    End If
End If
        
'F12
'Timer
If vcode = 123 Then
    Tiger = FindWindow("Dictation", vbNullString)
    If Tiger > 0 And GetWindowState(Tiger) <> vbMinimized Then
        ShowWindow Tiger, SW_MINIMIZE
        Sluty = 0
        ATidalX.ShellKeyLoad.Enabled = False
    End If
End If
        
'F12
'Timer
If vcode = 123 And App.title = "Extreme Lock Build: 2011" Then
    Sluty = 0
    ATidalX.ShellKeyLoad.Enabled = False
End If

'F12 New Ting about ?? or f7 f7 carte browsing
If Ctrl = False And vcode = 123 Or vcode = 118 Then
    If FindWindow("IEFrame", vbNullString) = GetForegroundWindow And _
        GetWindowState(FindWindow("IEFrame", vbNullString)) <> vbMinimized Then
        'run auto entry passwords
        'keywindowcheck
        If vcode = 123 Then Sluty = 1   ' F12
        Callback = 2: Exit Function
    End If
End If

Dim HH$

HH$ = GetWindowTitle(GetForegroundWindow)
Dim XX22
Dim PIP
XX22 = 0
PIP = 0
Dim XX5
If InStr(HH$, "IceView (Unregistered)") > 0 Then XX22 = True: PIP = 1
If InStr(HH$, "Notepad2") > 0 Then XX22 = True
If InStr(HH$, "Microsoft Excel") > 0 Then XX22 = True
If InStr(HH$, "Microsoft Word") > 0 Then XX22 = True ': xx2 = True

If XX22 = True Then
    If vcode = 123 Or vcode = 118 Or vcode = 119 Or vcode = 120 Or vcode = 115 Then
        
        Callback = 2
        'If pip = 1 And State <> 0 And vcode = 115 Then Callback = 0
        If Callback = 2 Then Exit Function
    End If

    'F1 save file
    'Timer
    If vcode = 112 Then
        ATidalX.FileSaveFileExit.Enabled = True
        Callback = 2: Exit Function
    End If

    'If vcode = 107 Then
    '    If RKeyPlus >= 21 Then KeyPlus = True
    '    Callback = 2: Exit Function
    'End If
    
    'If vcode = 107 And Action = 0 Then
    '    Callback = 2: Exit Function
    'End If
End If
        
        
        
        
'# key ' state cntr
'Timer
If vcode = 222 And State = 2 And App.title <> "Extreme Lock Build: 2011" Then
    Sluty = 0
    ATidalX.ShellWinampLoggLoader.Enabled = True
    Callback = 2: Exit Function
End If
        
Dim edge
        
        
If vcode = 223 And IsVBrunningForCode2 = 0 Then 'esc
    Tiger = FindWindow("Notepad2", vbNullString)
    HH$ = GetWindowTitle(FindWindow("Notepad2", vbNullString))
    If InStr(HH$, "- Notepad2") > 0 Then XTiger = 1 Else XTiger = 0
    edge = 0
    If XTiger = 1 Then
        Callback = 2: Exit Function
    End If
End If

'# Swap keys /=\
If vcode = 111 And Action = 3 Then
    Call SendKey(220, 0) '# Swap keys /=\
    Callback = 2: Exit Function
End If

'If vcode = 107 And (IsVBrunningCode2 = True) Then  'Or ATidalX.hWnd = GetForegroundWindow) Then
'    Call SendKey(3, 2) '# Swap NumPad + to Break
'    Callback = 2: Exit Function
'End If

HH$ = GetWindowTitle(GetForegroundWindow)
If InStr(HH$, "IceView (Unregistered)") > 0 Then
    If vcode = 118 Then Callback = 2: Exit Function
End If

'---------
'Counters

If (State = 0 And vcode = 106 And Action = 2) Or (State = 0 And vcode = 145) Then
    If FindWindow("TfrmWinDowse", vbNullString) = 0 Then
        If vcode <> 145 Then Callback = 2: Exit Function
    End If
End If

If AnyAlt = True And vcode = 106 Then
    Callback = 2: Exit Function
End If
If Ctrl = True And vcode = 106 Then
    Callback = 2: Exit Function
End If

'If NoStateKey = True And vcode = 109 And Action = 2 Then
'    Callback = 2: Exit Function
'End If
'If AnyAlt = True And vcode = 109 Then
'    Callback = 2: Exit Function
'End If


FwnD = GetForegroundWindow
'VB code window
XX5 = 0

'If Ctrl = True And vcode = 109 Then
'If FwnD = FindWindow("TextviewWndClass60", vbNullString) Then xx5 = 1
'    If xx5 = 0 Then Callback = 2: Exit Function
'End If

If NoStateKey = True And vcode = 12 Then
    Callback = 2: Exit Function
End If
If AnyAlt = True And vcode = 12 Then
    Callback = 2: Exit Function
End If
If Ctrl = True And vcode = 12 Then
    Callback = 2: Exit Function
End If

If NoStateKey = 0 And vcode = 93 Then
    Callback = 2: Exit Function
End If
If AnyAlt = True And vcode = 93 Then
    Callback = 2: Exit Function
End If
If Ctrl = True And vcode = 93 Then
    Callback = 2: Exit Function
End If

'----------End Counters




'try put date in outlook e fuck up
'findwindow ("ATH_Note",vbnullstring) 'just outlook e
'If Ctrl = True And vcode = 120 Then 'Cntl Shit f9
    'SendKeys "-- Edit - " + Format$(Now, "DDD-DD-MMM HH:MM:SSa/p"), False
'End If

'Dodge Ctrl again
If Ctrl = True And Vcodez = 82 Then 'Ctrl R Dictor Up Down
    'Callback = 2: Exit Function
End If

'--------------------------------------------
'f4 stop winamp fastforward form dragging down box in explorers an inet explorer
If vcode = VK_F4 Then
    'WAFastFF = True 'F4 = Fast F '115

    If FindWindow("IEFrame", vbNullString) = GetForegroundWindow And _
        GetWindowState(FindWindow("IEFrame", vbNullString)) <> vbMinimized Then
            Callback = 2: Exit Function
    End If
End If

If vcode = 115 Then
    If FindWindow("ExploreWClass", vbNullString) = GetForegroundWindow And _
        GetWindowState(FindWindow("ExploreWClass", vbNullString)) <> vbMinimized Then
            Callback = 2: Exit Function
    End If
End If
'--------------------------------------------



If vcode = 117 Then
    Callback = 2: Exit Function
End If

If vcode = 117 And Ctrl = True Then
    Callback = 2: Exit Function
End If


'77 Ctrl M Hott Mini Foreground Window
'78 Ctrl N Hott Mini All Window at Cursor
'76 Ctrl L Hott Mini All Windows Bar My Favs
If App.title = "Tidal Information..." Then
    'more case of conrtol not work proper
    'If Ctrl = True And (vcode = 76 Or vcode = 77 Or vcode = 78) Then Callback = 2
    
End If


Exit Function
ErrTrap:

Dim A55, B55$

A55 = Err.Number
B55$ = Err.Description

Call DSkeybd_F.UnHookCommand_Click

MsgBox ("Error in Subroutine CallBack " + str(A55) + vbCrLf + B55$)
If IsIDE Then End: Stop
Resume Next

End Function



'***********************************************
'# Check, whether we are in the IDE
Function IsIDE() As Boolean
  Debug.Assert Not TestIDE(IsIDE)
  'TEST
  'IsIDE = False
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
  ' Test = False
End Function
'***********************************************
'***********************************************
'# Generates a keystroke


Sub SendKey(ByVal vk%, ByVal mSCAE%)
Dim flags&
'DPrint vk, mSCAE
Call SetStop(1)
DoEvents
Call SetStop(1)
DoEvents
Call SetStop(1)

If mSCAE And 1 Then Call keybd_event(16, 0, 0, 0) '# Shift
If mSCAE And 2 Then Call keybd_event(17, 0, 0, 0) '# Ctrl
If mSCAE And 4 Then Call keybd_event(18, 0, 0, 0) '# Alt
If mSCAE And 8 Then flags = KEYEVENTF_EXTENDEDKEY '# Extended

Call keybd_event(vk, 0, flags, 0)
Call keybd_event(vk, 0, flags Or KEYEVENTF_KEYUP, 0)

flags = KEYEVENTF_KEYUP
If mSCAE And 4 Then Call keybd_event(18, 0, flags, 0)
If mSCAE And 2 Then Call keybd_event(17, 0, flags, 0)
If mSCAE And 1 Then Call keybd_event(16, 0, flags, 0)

If SndKey = False Then Call SetStop(0)

End Sub

'***********************************************
Sub UnHook()
Call SetKeyboardHook(0, 0, 0)
IsHook = False
End Sub


Public Sub KeyStateTest(vcode)

    'Call CapsNumLockStatus(vcode)
 
End Sub

Sub CapsNumLockStatus(vcode)

'GetKeyState

ATidalX.Label14 = "-Caps-" + str(CapsStat) + "-Num-" + str(NumsStat) + "-Scr-" + str(ScrosStat)

'Turn caps off if shift pressed Sweet
'CapsFlag = False
'If Vcode = 160 Or Vcode = 161 Then
'    If CapsStat = True Then
'        Call CapsLockOFF
'        CapsFlag = True
'        State = 0
'    End If
'End If

'GetKeyState
'Label14 = "-Caps-" + str(CapsStat) + "-Num-" + str(NumsStat) + "-Scr-" + str(ScrosStat)

End Sub




Function GetKeyState()
    'Public CapsStat, NumsStat, ScrosStat
    Exit Function
    
    Dim abytBuffer(0 To 255) As Byte
    GetKeyboardState abytBuffer(0)

    CapsStat = abytBuffer(vbKeyCapital) > 0
    NumsStat = abytBuffer(vbKeyNumlock) > 0
    ScrosStat = abytBuffer(vbKeyScrollLock) > 0

End Function

Function SetKeyState(ByVal intKey As Integer, ByVal fTurnOn As Boolean)
    Dim abytBuffer(0 To 255) As Byte
    GetKeyboardState abytBuffer(0)
    abytBuffer(intKey) = CByte(Abs(fTurnOn))
    SetKeyboardState abytBuffer(0)
End Function
Function CapsLockON()
    SetKeyState vbKeyCapital, True
End Function
Private Function CapsLockOFF()
    SetKeyState vbKeyCapital, False
    'rr = SetKeys("CAPSLOCK_OFF")
End Function
Function NumLockON()
    SetKeyState vbKeyNumlock, True
End Function
Function NumLockOFF()
    SetKeyState vbKeyNumlock, False
End Function
Function ScrollLockON()
    SetKeyState vbKeyScrollLock, True
End Function
Function ScrollLockOFF()
    SetKeyState vbKeyScrollLock, False
End Function


Public Function IsVBrunningForMusick()
'WE COULD USE OUR ISIDE BUT MAYBE WORKING ON ANOTHER PROGAM CODE AND
'DONT WANT F8 START STOP PAUSING


Tiger = FindWindow("Winamp PE", vbNullString)
If Tiger = 0 Then Tiger = FindWindow("Winamp v1.x", vbNullString)
If Tiger = 0 Then Tiger = FindWindow("WinampVideoChild", vbNullString)

Dim FwnD
Dim Nx As Long
Dim Ny As Long

IsVBrunningForMusick = False

'ATidalX.FindCursor Nx, Ny
'Mwnd = WindowFromPoint(Nx, Ny)
'VB code window
'If Mwnd = FindWindow("wndclass_desked_gsk", vbNullString) Then IsVBrunningForMusick = True
'If Mwnd = FindWindow("#32770", "Microsoft Visual Basic") Then IsVBrunningForMusick = True
'If Mwnd = FindWindow("VbaWindow", vbNullString) Then IsVBrunningForMusick = True
'EVEN HOOVER OVER IF WANT
'LATER ON ANOTHER SUBJECT OF UNLOAD KEY HOOKER IN VB CODE MODE

FwnD = GetForegroundWindow
'VB code window

'If FwnD = FindWindow("wndclass_desked_gsk", vbNullString) Then IsVBrunningForMusick = True
'If FwnD = FindWindow("#32770", "Microsoft Visual Basic") Then IsVBrunningForMusick = True
'If FwnD = FindWindow("VbaWindow", vbNullString) Then IsVBrunningForMusick = True

'If FwnD = MEHwnD Then IsVBrunningForMusick = true


TEXTISVB = GetWindowTitle(FindWindow("wndclass_desked_gsk", vbNullString))
If InStr(TEXTISVB, " - Microsoft Visual Basic [") > 0 Then
    If FwnD = FindWindow("wndclass_desked_gsk", vbNullString) Then
        IsVBrunningForMusick = True
    End If
End If


'xxd = FindWindow("wndclass_desked_gsk", vbNullString)
'If GetWindowState(xxd) = vbMaximized Or GetWindowState(xxd) = -1 Then IsVBrunningForMusick = 1
'If GetWindowState(xxd) = vbMinimized Then
'IsVBrunningForMusick = false
'End If

'MsgBox (GetWindowState(xxd))
'IsVBrunning = False



'If TestHookinVB = True Then IsVBrunning = False
'NOT HERE WITH MUSIC BUT ABLE IN TEST KEYHOOK CODE MODE UNLOAD ITSELF


'ATidalX.Label37 = Trim(Str(Val(IsVBrunning)))



End Function


Public Function IsVBrunningForCode()
'WE COULD USE OUR ISIDE BUT MAYBE WORKING ON ANOTHER PROGAM CODE AND
'DONT WANT F8 START STOP PAUSING

'Tiger = FindWindow("Winamp PE", vbNullString)
'If Tiger = 0 Then Tiger = FindWindow("Winamp v1.x", vbNullString)
'If Tiger = 0 Then Tiger = FindWindow("WinampVideoChild", vbNullString)

Dim FwnD As Long, FwnD2 As Long, Mwnd As Long
Dim Nx As Long
Dim Ny As Long


ATidalX.FindCursor Nx, Ny
Mwnd = WindowFromPoint(Nx, Ny)
IsVBrunningForCode = IsVBrunningForCodeVAR


If Mwnd = Mwnd_OLD Then Exit Function
Mwnd_OLD = Mwnd


'VB code window
'If Mwnd = FindWindow("wndclass_desked_gsk", vbNullString) Then IsVBrunningForCode = True
'If Mwnd = FindWindow("#32770", "Microsoft Visual Basic") Then IsVBrunningForCode = True
'If Mwnd = FindWindow("VbaWindow", vbNullString) Then IsVBrunningForCode = True
'EVEN HOOVER OVER IF WANT
'LATER ON ANOTHER SUBJECT OF UNLOAD KEY HOOKER IN VB CODE MODE

FwnD = GetForegroundWindow
'VB code window

'If FwnD = FindWindow("wndclass_desked_gsk", vbNullString) Then IsVBrunningForCode = True
'If FwnD = FindWindow("#32770", "Microsoft Visual Basic") Then IsVBrunningForCode = True
'If FwnD = FindWindow("VbaWindow", vbNullString) Then IsVBrunningForCode = True

IsVBrunningForCodeVAR = False

FwnD2 = FindWindow("wndclass_desked_gsk", vbNullString)

'LOOKING FOR OUR OWN PROGRAM - BUT WE NEED ALL
'WRONG
TEXTISVB = GetWindowTitle(FwnD2)
'If InStr(TEXTISVB, " - Microsoft Visual Basic [") > 0 Then
If InStr(TEXTISVB, " - Microsoft Visual Basic [") > 0 Then
    'NEED LOST FOCUS THINK
    If FwnD = FwnD2 Then
        IsVBrunningForCode = True
        IsVBrunningForCodeVAR = True
        
        Exit Function
    End If
    
    If Mwnd = FwnD2 Then
        IsVBrunningForCode = True
        IsVBrunningForCodeVAR = True
        
        Exit Function
    End If
    
    'WITH CURSOR FROM POINT - IT DONT FIND PARENT
    'SO MORE CHECKING
    
    
    'DOUBLE DECODE WORK TRACE
    TXXRRS_CHANGE_VAR = False
    If GetActiveWindowPARENT_HWND(Mwnd, FwnD2) = True Then
        If Mwnd + FwnD2 <> TXXRRS Then
            TXXRRS = Mwnd + FwnD2
            TXXRRS_CHANGE_VAR = True
        End If
    End If
    
    
    
    
    If GetActiveWindowPARENT_HWND(Mwnd, FwnD2) = True And TXXRRS_CHANGE_VAR = True Then
        IsVBrunningForCode = True
        IsVBrunningForCodeVAR = True

    End If


End If

'xxd = FindWindow("wndclass_desked_gsk", vbNullString)
'If GetWindowState(xxd) = vbMaximized Or GetWindowState(xxd) = -1 Then IsVBrunningForCode = 1
'If GetWindowState(xxd) = vbMinimized Then


'MsgBox (GetWindowState(xxd))


'If TestHookinVB = True Then IsVBrunning = False
'NOT HERE WITH MUSIC BUT ABLE IN TEST KEYHOOK CODE MODE UNLOAD ITSELF

'ATidalX.Label37 = Trim(Str(Val(IsVBrunning)))


End Function

Public Function GetActiveWindowPARENT_HWND(ReturnParent As Long, MATCHParent As Long) As Long
    Dim i As Long, i2 As Long
    Dim j As Long
    GetActiveWindowPARENT_HWND = False
    i = ReturnParent
    If ReturnParent Then
        Do While i <> 0
            j = i
            i = GetParent(i)
            If i = MATCHParent Then
                GetActiveWindowPARENT_HWND = True
                Exit Do
            End If
        Loop
    End If
'    If i = 0 Then i = j
'    GetActiveWindowPARENT_HWND = i

End Function

