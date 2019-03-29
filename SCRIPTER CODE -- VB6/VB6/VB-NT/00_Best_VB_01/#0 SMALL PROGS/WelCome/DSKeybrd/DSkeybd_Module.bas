Attribute VB_Name = "DSkeybd_M"
'Timer13_Timer() ------- Look for this one for sendkeys on progs

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
Public action, state, vcode, scode
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
Declare Function SkipInjected& Lib "dskbhook" (ByVal Skip&)
Declare Function NoLLHook& Lib "dskbhook" (ByVal LLHook&)
Declare Function SetStop& Lib "dskbhook" (ByVal Stp&)
Declare Function HookVB& Lib "dskbhook" (ByVal Hook&)
Declare Function NoTL& Lib "dskbhook" (ByVal TL&)
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
Public Function Callback&(ByVal action&, ByVal state&, ByVal vcode&, ByVal scode&)

On Local Error GoTo ErrTrap

'Check = SlowKeyPress()

If vcode = 0 Then vcode = 93: action = 2 'Test

'If TestKeyLoggOff = False Then
'# Display the passed values
With DSkeybd_F.List1
  'If .ListCount > 500 Then .Clear
  '# Add the new entry at the first position and highlight it
'  .AddItem "   " & Action & Chr$(9) & State & Chr$(9) & VCode & Chr$(9) & Scode, 0
  .AddItem Format$(action, "000 ") & Format$(state, "000 ") & Format$(vcode, "000 ") & Format$(scode, "000 ")
  .Selected(.NewIndex) = True
    DSkeybd_F.Label5 = .ListCount

End With
'End If

'KeyLog2 = KeyLog2 + Format$(Now, "DDD DD-MMM-YY HH:MM:SS") + " - KEY =" + Str(Action) + Chr$(9) + Str(State) + Chr$(9) + Str(Vcode) + Chr$(9) + Str(Scode) + vbCrLf

'DPrint action, state, vcode, scode
Down = action And 2: Up = action And 4: Ext = action And 1
Ctrl = False: LeftAlt = False: RightAlt = False: AnyAlt = False: NoStateKey = False
Shift = state And 1


Call KeyStateTest(vcode)

'If CapsStat = False Then If (State And 1) = 1 Then Shift = False: State = 0


If (state And 2) = 2 Then Ctrl = True
If (state And 3) = 3 Then Ctrl = True
If (state And 4) = 4 Then LeftAlt = True: AnyAlt = True
If (state And 5) = 5 Then LeftAlt = True: AnyAlt = True
If (state And 6) = 6 Then RightAlt = True: AnyAlt = True
If (state And 7) = 7 Then RightAlt = True: AnyAlt = True

If AnyAlt = False And Ctrl = False And Shift = False Then NoStateKey = True Else NoStateKey = False


Callback = 0 '#never suppress any keys yet







Exit Function
ErrTrap:

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
'IsIDE = False
End Function
Private Function TestIDE(Test As Boolean) As Boolean
  Test = True
End Function

'***********************************************
'***********************************************
'# Generates a keystroke
Sub SendKey(ByVal vk%, ByVal mSCAE%)
Dim Flags&
'DPrint vk, mSCAE
Call SetStop(1)
DoEvents
Call SetStop(1)
DoEvents
Call SetStop(1)

If mSCAE And 1 Then Call keybd_event(16, 0, 0, 0) '# Shift
If mSCAE And 2 Then Call keybd_event(17, 0, 0, 0) '# Ctrl
If mSCAE And 4 Then Call keybd_event(18, 0, 0, 0) '# Alt
If mSCAE And 8 Then Flags = KEYEVENTF_EXTENDEDKEY '# Extended

Call keybd_event(vk, 0, Flags, 0)
Call keybd_event(vk, 0, Flags Or KEYEVENTF_KEYUP, 0)

Flags = KEYEVENTF_KEYUP
If mSCAE And 4 Then Call keybd_event(18, 0, Flags, 0)
If mSCAE And 2 Then Call keybd_event(17, 0, Flags, 0)
If mSCAE And 1 Then Call keybd_event(16, 0, Flags, 0)

If SndKey = False Then Call SetStop(0)

End Sub

'***********************************************
Sub UnHook()
Call SetKeyboardHook(0, 0, 0)
IsHook = False
End Sub


Public Sub KeyStateTest(vcode)

    Call CapsNumLockStatus(vcode)
 
End Sub

Sub CapsNumLockStatus(vcode)

GetKeyState

'ATidalX.Label14 = "-Caps-" + str(CapsStat) + "-Num-" + str(NumsStat) + "-Scr-" + str(ScrosStat)

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
Label14 = "-Caps-" + str(CapsStat) + "-Num-" + str(NumsStat) + "-Scr-" + str(ScrosStat)

End Sub




Function GetKeyState()
    'Public CapsStat, NumsStat, ScrosStat
    
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
Function CapsLockOFF()
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



